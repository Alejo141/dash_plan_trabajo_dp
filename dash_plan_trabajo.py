"""
DISPOWER — Tablero Dirección Comercial ZNI
==========================================
Conectado a SharePoint / OneDrive for Business via Microsoft Graph API.
Lee Y escribe directamente en el Excel compartido — los cambios se ven en tiempo real.

CONFIGURACIÓN REQUERIDA (.streamlit/secrets.toml)
──────────────────────────────────────────────────
[graph]
tenant_id     = "tu-tenant-id"          # ID del directorio Azure AD
client_id     = "tu-client-id"          # ID de la aplicación registrada
client_secret = "tu-client-secret"      # Secreto de la app
sharepoint_url = "https://suncompanycol-my.sharepoint.com"
share_link    = "https://suncompanycol-my.sharepoint.com/personal/javier_agudelo_dispower_co/_layouts/15/guestaccess.aspx?share=IQAR5H9drEd6RJJ06JnXkTVAAUi79Eg4HXu6KxcLFNlLfPk&e=jGJRo3"

VER README_GRAPH.md para instrucciones detalladas de configuración.
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime
import io, warnings, time
import requests

warnings.filterwarnings("ignore")

st.set_page_config(
    page_title="DISPOWER · ZNI",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── PALETA Y CONSTANTES ────────────────────────────────────────────────────
AREA_COLOR = {
    "Gestión Social": "#1A5C3A", "SAC": "#2E6DA4",
    "BI & Analítica": "#E67E22", "SUI & Subsidios": "#7D3C98", "Dirección": "#3D3D3D",
}
AREA_BG = {
    "Gestión Social": "#D6F0E3", "SAC": "#D6E8F7",
    "BI & Analítica": "#FDEBD0", "SUI & Subsidios": "#EAD8F5", "Dirección": "#F5F5F5",
}
ESTADOS    = ["Sin iniciar", "En curso", "Cerrada", "Bloqueada", "Cancelada"]
PRIORIDADES = ["CRÍTICA", "ALTA", "MEDIA", "BAJA"]
ESTADO_COL = {"Sin iniciar": "#E74C3C", "En curso": "#F39C12", "Cerrada": "#27AE60", "Bloqueada": "#D85A30", "Cancelada": "#888"}
ESTADO_BG  = {"Sin iniciar": "#FADBD8", "En curso": "#FEF9E7", "Cerrada": "#D6F0E3", "Bloqueada": "#FDEBD0", "Cancelada": "#F5F5F5"}
PRIO_COL   = {"CRÍTICA": "#C0392B", "ALTA": "#E67E22", "MEDIA": "#27AE60", "BAJA": "#2E6DA4"}
PRIO_BG    = {"CRÍTICA": "#FADBD8", "ALTA": "#FDEBD0", "MEDIA": "#D6F0E3", "BAJA": "#D6E8F7"}

# Columnas exactas del Excel (hoja ✅ Tareas)
COLS = ["ID", "ÁREA", "PROYECTO", "TAREA", "CATEGORÍA", "RESPONSABLE",
        "FECHA INICIO", "FECHA LÍMITE", "PRIORIDAD", "AVANCE", "ESTADO",
        "OBSERVACIÓN", "FECHA CIERRE", "CERRADA POR"]

st.markdown("""<style>
.kpi{background:white;border-radius:10px;padding:16px 18px;
     border-left:5px solid #2E6DA4;box-shadow:0 1px 5px rgba(0,0,0,.07);margin-bottom:10px}
.kv{font-size:2rem;font-weight:700;margin:4px 0}
.kl{font-size:.7rem;color:#888;text-transform:uppercase;letter-spacing:.05em}
.badge{display:inline-block;border-radius:20px;padding:3px 10px;font-size:.72rem;font-weight:600}
.info{background:#EBF5FB;border-left:4px solid #2E6DA4;border-radius:4px;
      padding:10px 14px;margin-bottom:10px;font-size:.86rem;color:#1B3A5C}
.warn{background:#FEF9E7;border-left:4px solid #F39C12;border-radius:4px;
      padding:10px 14px;margin-bottom:10px;font-size:.86rem;color:#7D6608}
.ok  {background:#D6F0E3;border-left:4px solid #27AE60;border-radius:4px;
      padding:10px 14px;margin-bottom:10px;font-size:.86rem;color:#1A5C3A}
</style>""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# MICROSOFT GRAPH API — AUTENTICACIÓN Y OPERACIONES
# ════════════════════════════════════════════════════════════════════════════

def get_graph_token() -> str | None:
    """Obtiene token de acceso OAuth2 para Microsoft Graph."""
    try:
        cfg = st.secrets.get("graph", {})
        tenant_id    = cfg.get("tenant_id", "")
        client_id    = cfg.get("client_id", "")
        client_secret = cfg.get("client_secret", "")
        if not all([tenant_id, client_id, client_secret]):
            return None
        url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        data = {
            "grant_type":    "client_credentials",
            "client_id":     client_id,
            "client_secret": client_secret,
            "scope":         "https://graph.microsoft.com/.default",
        }
        r = requests.post(url, data=data, timeout=15)
        if r.status_code == 200:
            return r.json().get("access_token")
        return None
    except Exception:
        return None

@st.cache_data(ttl=3300, show_spinner=False)
def get_cached_token() -> str | None:
    return get_graph_token()

def graph_get(path: str, token: str) -> dict | None:
    """GET a Microsoft Graph."""
    try:
        r = requests.get(
            f"https://graph.microsoft.com/v1.0/{path}",
            headers={"Authorization": f"Bearer {token}"},
            timeout=20,
        )
        return r.json() if r.status_code == 200 else None
    except Exception:
        return None

def graph_patch(path: str, token: str, body: dict) -> bool:
    """PATCH a Microsoft Graph (actualiza valores en Excel)."""
    try:
        r = requests.patch(
            f"https://graph.microsoft.com/v1.0/{path}",
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type":  "application/json",
            },
            json=body,
            timeout=20,
        )
        return r.status_code in [200, 204]
    except Exception:
        return False

def get_drive_item_id(token: str) -> tuple[str | None, str | None]:
    """
    Obtiene el driveId y itemId del archivo Excel.
    Intenta múltiples estrategias en orden.
    """
    import base64
    cfg = st.secrets.get("graph", {})

    # ── Estrategia 1: item_id y drive_id explícitos en secrets (más confiable) ──
    explicit_item  = cfg.get("item_id", "")
    explicit_drive = cfg.get("drive_id", "")
    if explicit_item and explicit_drive:
        return explicit_drive, explicit_item

    # ── Estrategia 2: /me/drive — buscar por nombre de archivo ──
    try:
        filename = cfg.get("filename", "DISPOWER_Tareas_Streamlit.xlsx")
        me_result = graph_get(f"me/drive/root/search(q='{filename}')", token)
        if me_result and me_result.get("value"):
            item = me_result["value"][0]
            drive_id = item.get("parentReference", {}).get("driveId")
            item_id  = item.get("id")
            if drive_id and item_id:
                return drive_id, item_id
    except Exception:
        pass

    # ── Estrategia 3: /shares con el share_link ──
    share_url = cfg.get("share_link", "")
    if share_url:
        try:
            encoded = base64.urlsafe_b64encode(share_url.encode()).decode().rstrip("=")
            share_token = f"u!{encoded}"
            result = graph_get(f"shares/{share_token}/driveItem", token)
            if result and result.get("id"):
                item_id  = result.get("id")
                drive_id = result.get("parentReference", {}).get("driveId")
                if drive_id and item_id:
                    return drive_id, item_id
        except Exception:
            pass

        # ── Estrategia 4: /shares con solo el token del enlace (parámetro share=...) ──
        try:
            import re
            match = re.search(r"share=([^&]+)", share_url)
            if match:
                share_token2 = f"u!{match.group(1)}"
                result2 = graph_get(f"shares/{share_token2}/driveItem", token)
                if result2 and result2.get("id"):
                    item_id2  = result2.get("id")
                    drive_id2 = result2.get("parentReference", {}).get("driveId")
                    if drive_id2 and item_id2:
                        return drive_id2, item_id2
        except Exception:
            pass

    # ── Estrategia 5: Listar archivos recientes del usuario ──
    try:
        recent = graph_get("me/drive/recent", token)
        if recent and recent.get("value"):
            filename_cfg = cfg.get("filename", "DISPOWER")
            for item in recent["value"]:
                if filename_cfg.lower() in item.get("name","").lower():
                    drive_id = item.get("parentReference", {}).get("driveId")
                    item_id  = item.get("id")
                    if drive_id and item_id:
                        return drive_id, item_id
    except Exception:
        pass

    return None, None

@st.cache_data(ttl=25, show_spinner=False)
def load_excel_from_graph(token: str) -> pd.DataFrame | None:
    """
    Descarga el Excel desde SharePoint vía Graph y lo parsea.
    Cache de 25 segundos — balance entre frescura y velocidad.
    """
    try:
        drive_id, item_id = get_drive_item_id(token)
        if not drive_id or not item_id:
            return None

        # Descargar el contenido del archivo
        r = requests.get(
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content",
            headers={"Authorization": f"Bearer {token}"},
            timeout=30,
        )
        if r.status_code != 200:
            return None

        raw = r.content
        xl  = pd.ExcelFile(io.BytesIO(raw))

        # Buscar la hoja de tareas
        target = None
        for sn in xl.sheet_names:
            if any(k in sn.lower() for k in ["tarea", "cierre", "task"]):
                target = sn; break
        if not target:
            target = xl.sheet_names[0]

        df = xl.parse(target, dtype=str).fillna("")
        # Normalizar columnas
        df.columns = [str(c).strip().replace("\n", " ") for c in df.columns]
        col_map = {
            "ÁREA": "ÁREA", "AREA": "ÁREA",
            "PROYECTO / MUNICIPIO": "PROYECTO", "PROYECTO": "PROYECTO",
            "CATEGORÍA": "CATEGORÍA", "CATEGORIA": "CATEGORÍA",
            "% AVANCE": "AVANCE", "AVANCE": "AVANCE",
            "OBSERVACIÓN / BLOQUEO": "OBSERVACIÓN", "OBSERVACION": "OBSERVACIÓN",
            "FECHA LÍMITE": "FECHA LÍMITE", "FECHA LIMITE": "FECHA LÍMITE",
        }
        rename = {}
        for col in df.columns:
            for alias, canon in col_map.items():
                if alias.upper() == col.upper():
                    rename[col] = canon; break
        df = df.rename(columns=rename)

        # Filtrar solo filas de tareas (ID = T-XXX)
        if "ID" in df.columns:
            df = df[df["ID"].astype(str).str.match(r"T-\d+", na=False)].copy()

        for c in COLS:
            if c not in df.columns:
                df[c] = ""

        return df[COLS].reset_index(drop=True) if len(df) > 0 else None

    except Exception as e:
        st.error(f"Error leyendo el Excel: {e}")
        return None

def save_excel_to_sharepoint(df: pd.DataFrame, token: str) -> bool:
    """
    Guarda el DataFrame completo de vuelta en SharePoint sobrescribiendo el Excel.
    Usa el endpoint /content con PUT para subir el archivo actualizado.
    """
    try:
        drive_id, item_id = get_drive_item_id(token)
        if not drive_id or not item_id:
            return False

        # 1. Descargar el archivo original para preservar formato
        r = requests.get(
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content",
            headers={"Authorization": f"Bearer {token}"},
            timeout=30,
        )
        if r.status_code != 200:
            return False

        raw = r.content

        # 2. Abrir con openpyxl y actualizar la hoja de tareas
        from openpyxl import load_workbook
        wb = load_workbook(io.BytesIO(raw))

        target = None
        for sn in wb.sheetnames:
            if any(k in sn.lower() for k in ["tarea", "cierre", "task"]):
                target = sn; break
        if not target:
            target = wb.sheetnames[0]

        ws = wb[target]

        # Encontrar la fila de encabezado
        hrow = None
        for ri in range(1, min(12, ws.max_row + 1)):
            vals = [str(ws.cell(ri, c).value or "").upper() for c in range(1, ws.max_column + 1)]
            if any("TAREA" in v or "ÁREA" in v or "AREA" in v for v in vals):
                hrow = ri; break

        if hrow is None:
            return False

        # Mapear columnas del Excel a índices
        hdrs = [str(ws.cell(hrow, c).value or "").strip().replace("\n", " ")
                for c in range(1, ws.max_column + 1)]
        col_map2 = {}
        alias_back = {
            "PROYECTO / MUNICIPIO": "PROYECTO",
            "% AVANCE": "AVANCE",
            "OBSERVACIÓN / BLOQUEO": "OBSERVACIÓN",
            "FECHA LÍMITE": "FECHA LÍMITE",
        }
        for ci, h in enumerate(hdrs, 1):
            canon = alias_back.get(h.upper(), h.upper())
            # Buscar en COLS
            for c in COLS:
                if c.upper() == canon or c.upper() == h.upper():
                    col_map2[c] = ci; break

        # Limpiar datos existentes
        for rr in range(hrow + 1, ws.max_row + 5):
            for cc in range(1, ws.max_column + 1):
                ws.cell(rr, cc).value = None

        # Escribir datos actualizados
        for ri2, (_, row) in enumerate(df.iterrows(), hrow + 1):
            for col_name, ci2 in col_map2.items():
                ws.cell(ri2, ci2).value = str(row.get(col_name, ""))

        # 3. Subir el archivo actualizado a SharePoint
        buf = io.BytesIO()
        wb.save(buf)
        new_content = buf.getvalue()

        upload_r = requests.put(
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content",
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            },
            data=new_content,
            timeout=60,
        )
        return upload_r.status_code in [200, 201]

    except Exception as e:
        st.error(f"Error guardando en SharePoint: {e}")
        return False

# ════════════════════════════════════════════════════════════════════════════
# ESTADO DE SESIÓN
# ════════════════════════════════════════════════════════════════════════════
if "hist" not in st.session_state:   st.session_state["hist"] = []
if "df"   not in st.session_state:   st.session_state["df"]   = None
if "token" not in st.session_state:  st.session_state["token"] = None
if "last_load" not in st.session_state: st.session_state["last_load"] = 0

def add_hist(msg: str):
    st.session_state["hist"].insert(0, {
        "ts":  datetime.now().strftime("%d/%m/%Y %H:%M"),
        "msg": msg,
    })

def badge(text, fg, bg):
    return f'<span class="badge" style="background:{bg};color:{fg}">{text}</span>'
def estado_badge(e):  return badge(e, ESTADO_COL.get(e,"#888"), ESTADO_BG.get(e,"#eee"))
def prio_badge(p):    return badge(p, PRIO_COL.get(p,"#888"),   PRIO_BG.get(p,"#eee"))

def next_id(df: pd.DataFrame) -> str:
    nums = pd.to_numeric(df["ID"].str.extract(r"T-(\d+)")[0], errors="coerce").dropna()
    return f"T-{int(nums.max())+1:03d}" if len(nums) else "T-001"

def get_token() -> str | None:
    if st.session_state["token"]:
        return st.session_state["token"]
    token = get_cached_token()
    st.session_state["token"] = token
    return token

def load_data(force: bool = False) -> pd.DataFrame | None:
    now = time.time()
    since = now - st.session_state["last_load"]
    # Recargar si fuerza o si pasaron más de 30s
    if force or st.session_state["df"] is None or since > 30:
        token = get_token()
        if token:
            with st.spinner("Sincronizando con SharePoint..."):
                df = load_excel_from_graph(token)
            if df is not None:
                st.session_state["df"] = df
                st.session_state["last_load"] = now
    return st.session_state.get("df")

def commit(df_new: pd.DataFrame, msg: str):
    """Guarda cambios localmente Y en SharePoint."""
    st.session_state["df"] = df_new.copy()
    token = get_token()
    if token:
        with st.spinner("Guardando en SharePoint..."):
            ok = save_excel_to_sharepoint(df_new, token)
        if ok:
            add_hist(f"✅ {msg}")
            st.success("Guardado en SharePoint — el Excel se actualizó en tiempo real.")
        else:
            add_hist(f"⚠️ {msg} (error al guardar en SharePoint)")
            st.warning("Los cambios se aplicaron en el tablero, pero no se pudieron guardar en el Excel. Revisa la configuración.")
    else:
        add_hist(f"💾 {msg} (sin conexión Graph)")
        st.info("Cambios guardados en sesión. Configura Graph API para persistir en SharePoint.")

# ── Cargar datos ──────────────────────────────────────────────────────────
token = get_token()
graph_ok = token is not None
df = load_data()

# ── SIDEBAR ────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚡ DISPOWER")
    st.markdown("**Dirección Comercial ZNI**")
    st.markdown("---")

    pagina = st.radio("Navegación", [
        "🏠 Dashboard", "📋 Lista de Tareas", "🗂️ Kanban",
        "📊 Proyectos ZNI", "➕ Nueva Tarea", "📜 Historial", "⚙️ Configuración",
    ])

    st.markdown("---")
    if graph_ok:
        st.markdown("🟢 **SharePoint conectado**")
        st.caption("Cambios se guardan en el Excel automáticamente.")
    else:
        st.markdown("🔴 **SharePoint no configurado**")
        st.caption("Ve a ⚙️ Configuración para conectar.")

    if df is not None:
        since_sec = int(time.time() - st.session_state["last_load"])
        st.caption(f"{len(df)} tareas · hace {since_sec}s")

    col_r1, col_r2 = st.columns(2)
    with col_r1:
        if st.button("🔄 Recargar"):
            load_excel_from_graph.clear()
            load_data(force=True)
            st.rerun()
    with col_r2:
        if graph_ok and st.button("📤 Forzar sync"):
            if df is not None:
                with st.spinner("Subiendo..."):
                    ok = save_excel_to_sharepoint(df, token)
                st.success("OK" if ok else "Error")

# ════════════════════════════════════════════════════════════════════════════
# MODO SIN CONFIGURACIÓN — Mostrar guía
# ════════════════════════════════════════════════════════════════════════════
if not graph_ok:
    if pagina != "⚙️ Configuración":
        st.warning("⚠️ Graph API no configurada — el tablero no puede conectarse a SharePoint. Ve a **⚙️ Configuración** para ver las instrucciones.")

if df is None and pagina not in ["⚙️ Configuración", "📜 Historial"]:
    st.info("No hay datos cargados. Si Graph API está configurada, haz clic en 🔄 Recargar en el sidebar.")
    st.stop()

# ════════════════════════════════════════════════════════════════════════════
# DASHBOARD
# ════════════════════════════════════════════════════════════════════════════
if pagina == "🏠 Dashboard":
    st.markdown("## ⚡ Dashboard Ejecutivo")
    st.caption("Director: Javier Alejandro Agudelo · Plan Estratégico 6 Meses · ZNI / SISFV")

    total    = len(df)
    cerradas = (df["ESTADO"] == "Cerrada").sum()
    criticas = ((df["PRIORIDAD"] == "CRÍTICA") & (df["ESTADO"] != "Cerrada")).sum()
    en_curso = (df["ESTADO"] == "En curso").sum()
    bloq     = (df["ESTADO"] == "Bloqueada").sum()
    try:    avg_av = round(pd.to_numeric(df["AVANCE"], errors="coerce").fillna(0).mean())
    except: avg_av = 0

    cols_m = st.columns(6)
    for col, (val, lbl, sub, cc) in zip(cols_m, [
        (total,    "Total tareas",      f"{cerradas} cerradas",              "#2E6DA4"),
        (criticas, "Críticas pend.",    "atención urgente",                  "#C0392B"),
        (en_curso, "En curso",          "activas",                           "#E67E22"),
        (bloq,     "Bloqueadas",        "requieren decisión",                "#D85A30"),
        (f"{avg_av}%","Avance global",  "promedio plan",                     "#27AE60"),
        (cerradas, "Completadas",       f"{round(cerradas/total*100) if total else 0}% del plan", "#27AE60"),
    ]):
        col.markdown(f'<div class="kpi" style="border-left-color:{cc}">'
                     f'<div class="kl">{lbl}</div><div class="kv" style="color:{cc}">{val}</div>'
                     f'<div style="font-size:.7rem;color:#aaa">{sub}</div></div>', unsafe_allow_html=True)

    st.markdown("---")
    col_left, col_right = st.columns([2, 1])

    with col_left:
        st.markdown("**Avance por área**")
        rows_a = []
        for area in AREA_COLOR:
            s = df[df["ÁREA"] == area]
            if not len(s): continue
            try:    av = round(pd.to_numeric(s["AVANCE"], errors="coerce").fillna(0).mean())
            except: av = 0
            rows_a.append({"Área": area, "Avance": av})
        if rows_a:
            fig = go.Figure(go.Bar(
                x=[r["Área"] for r in rows_a], y=[r["Avance"] for r in rows_a],
                marker_color=[AREA_COLOR[r["Área"]] for r in rows_a], opacity=.85,
                text=[f"{r['Avance']}%" for r in rows_a], textposition="outside",
            ))
            fig.add_hline(y=100, line_dash="dash", line_color="#27AE60", line_width=1.5)
            fig.update_layout(height=260, margin=dict(l=0,r=0,t=10,b=0),
                plot_bgcolor="white", paper_bgcolor="white",
                yaxis=dict(title="% avance", range=[0, 120]), showlegend=False,
                font=dict(family="Arial", size=11))
            st.plotly_chart(fig, use_container_width=True)

    with col_right:
        st.markdown("**Distribución de estados**")
        ec = df["ESTADO"].value_counts()
        if len(ec):
            fig2 = go.Figure(go.Pie(
                labels=ec.index.tolist(), values=ec.values.tolist(),
                marker_colors=[ESTADO_COL.get(e, "#888") for e in ec.index],
                hole=.5, textinfo="value+percent",
            ))
            fig2.update_layout(height=260, margin=dict(l=0,r=0,t=10,b=10),
                paper_bgcolor="white", showlegend=True,
                legend=dict(orientation="h", y=-.15), font=dict(family="Arial", size=11))
            st.plotly_chart(fig2, use_container_width=True)

    st.markdown("---")
    st.markdown("**Tareas críticas pendientes:**")
    crit = df[(df["PRIORIDAD"] == "CRÍTICA") & (df["ESTADO"].isin(["Sin iniciar", "Bloqueada"]))]
    if len(crit):
        for _, r in crit.iterrows():
            c1, c2, c3, c4 = st.columns([1, 5, 2, 2])
            c1.markdown(f'<b style="color:#C0392B">{r["ID"]}</b>', unsafe_allow_html=True)
            c2.write(r["TAREA"][:80])
            c3.markdown(f'<small style="color:{AREA_COLOR.get(r["ÁREA"],"#888")}">{r["ÁREA"]}</small>', unsafe_allow_html=True)
            c4.markdown(estado_badge(r["ESTADO"]), unsafe_allow_html=True)
    else:
        st.success("Sin tareas críticas pendientes.")

# ════════════════════════════════════════════════════════════════════════════
# LISTA DE TAREAS
# ════════════════════════════════════════════════════════════════════════════
elif pagina == "📋 Lista de Tareas":
    st.markdown("## 📋 Lista de Tareas")

    fa_v = st.selectbox("Área", ["Todas"] + sorted(df["ÁREA"].unique().tolist()), key="la")
    c1f, c2f, c3f = st.columns(3)
    with c1f: fe_v = st.selectbox("Estado",    ["Todos"]  + ESTADOS,    key="le")
    with c2f: fp_v = st.selectbox("Prioridad", ["Todas"]  + PRIORIDADES, key="lp")
    with c3f: fb_v = st.text_input("Buscar", placeholder="Tarea, proyecto, responsable...", key="lb")

    filt = df.copy()
    if fa_v != "Todas":  filt = filt[filt["ÁREA"]      == fa_v]
    if fe_v != "Todos":  filt = filt[filt["ESTADO"]    == fe_v]
    if fp_v != "Todas":  filt = filt[filt["PRIORIDAD"] == fp_v]
    if fb_v:
        mask = (filt["TAREA"].str.contains(fb_v, case=False, na=False) |
                filt["PROYECTO"].str.contains(fb_v, case=False, na=False) |
                filt["RESPONSABLE"].str.contains(fb_v, case=False, na=False))
        filt = filt[mask]

    st.caption(f"{len(filt)} de {len(df)} tareas · Último sync: {datetime.fromtimestamp(st.session_state['last_load']).strftime('%H:%M:%S') if st.session_state['last_load'] else '—'}")
    st.markdown("---")

    for idx, row in filt.iterrows():
        try:    av_v = int(float(row.get("AVANCE", "0") or 0))
        except: av_v = 0
        area_c = AREA_COLOR.get(row["ÁREA"], "#888")

        with st.expander(f"**{row['ID']}** — {row['TAREA'][:85]}{'…' if len(row['TAREA']) > 85 else ''}"):

            h1c, h2c, h3c, h4c = st.columns(4)
            h1c.markdown(estado_badge(row["ESTADO"]),    unsafe_allow_html=True)
            h2c.markdown(prio_badge(row["PRIORIDAD"]),   unsafe_allow_html=True)
            h3c.markdown(f'<span style="color:{area_c};font-weight:600;font-size:.8rem">{row["ÁREA"]}</span>', unsafe_allow_html=True)
            h4c.markdown(f'<span style="color:#888;font-size:.8rem">{row["RESPONSABLE"]}</span>', unsafe_allow_html=True)

            prog_col = "#27AE60" if av_v >= 80 else ("#F39C12" if av_v >= 40 else "#E74C3C")
            st.markdown(
                f'<div style="height:4px;background:#eee;border-radius:2px;margin:6px 0">'
                f'<div style="height:4px;background:{prog_col};width:{av_v}%;border-radius:2px"></div></div>'
                f'<div style="font-size:.72rem;color:#aaa">Avance: {av_v}% · Proyecto: {row["PROYECTO"]} · Cat: {row["CATEGORÍA"]}</div>',
                unsafe_allow_html=True)

            if row.get("OBSERVACIÓN", "").strip():
                st.markdown(f'<div class="warn">{row["OBSERVACIÓN"]}</div>', unsafe_allow_html=True)

            st.markdown("---")
            e1, e2, e3 = st.columns(3)
            with e1:
                new_est = st.selectbox("Estado", ESTADOS,
                    index=ESTADOS.index(row["ESTADO"]) if row["ESTADO"] in ESTADOS else 0,
                    key=f"est_{row['ID']}")
            with e2:
                new_av = st.slider("Avance %", 0, 100, av_v, key=f"av_{row['ID']}")
            with e3:
                new_resp = st.text_input("Responsable", value=row["RESPONSABLE"], key=f"rsp_{row['ID']}")

            new_obs = st.text_area("Observación / bloqueo", value=row.get("OBSERVACIÓN", ""), height=65, key=f"obs_{row['ID']}")
            try:
                fl_val = datetime.strptime(row.get("FECHA LÍMITE", ""), "%d/%m/%Y").date() if row.get("FECHA LÍMITE", "").strip() else None
            except: fl_val = None
            new_fl = st.date_input("Fecha límite", value=fl_val, key=f"fl_{row['ID']}")

            b1, b2, b3, b4 = st.columns(4)
            with b1:
                if st.button("💾 Guardar cambios", key=f"s_{row['ID']}"):
                    df_new = st.session_state["df"].copy()
                    m = df_new["ID"] == row["ID"]
                    df_new.loc[m, "ESTADO"]      = new_est
                    df_new.loc[m, "AVANCE"]      = str(new_av)
                    df_new.loc[m, "RESPONSABLE"] = new_resp
                    df_new.loc[m, "OBSERVACIÓN"] = new_obs
                    if new_fl:
                        df_new.loc[m, "FECHA LÍMITE"] = new_fl.strftime("%d/%m/%Y")
                    commit(df_new, f"{row['ID']} actualizada: estado={new_est}, avance={new_av}%")
                    st.rerun()

            with b2:
                if row["ESTADO"] != "Cerrada":
                    if st.button("✅ Cerrar tarea", key=f"c_{row['ID']}"):
                        df_new = st.session_state["df"].copy()
                        m = df_new["ID"] == row["ID"]
                        df_new.loc[m, "ESTADO"]      = "Cerrada"
                        df_new.loc[m, "AVANCE"]      = "100"
                        df_new.loc[m, "FECHA CIERRE"] = datetime.now().strftime("%d/%m/%Y")
                        commit(df_new, f"{row['ID']} CERRADA: {row['TAREA'][:45]}")
                        st.rerun()
                else:
                    if st.button("🔄 Reabrir", key=f"r_{row['ID']}"):
                        df_new = st.session_state["df"].copy()
                        m = df_new["ID"] == row["ID"]
                        df_new.loc[m, "ESTADO"]      = "En curso"
                        df_new.loc[m, "FECHA CIERRE"] = ""
                        commit(df_new, f"{row['ID']} reabierta")
                        st.rerun()

            with b3:
                if st.button("🟠 Bloquear", key=f"bl_{row['ID']}"):
                    df_new = st.session_state["df"].copy()
                    df_new.loc[df_new["ID"] == row["ID"], "ESTADO"] = "Bloqueada"
                    commit(df_new, f"{row['ID']} marcada BLOQUEADA")
                    st.rerun()

            with b4:
                if st.button("🗑️ Eliminar", key=f"d_{row['ID']}"):
                    df_new = st.session_state["df"].copy()
                    df_new = df_new[df_new["ID"] != row["ID"]].reset_index(drop=True)
                    commit(df_new, f"{row['ID']} ELIMINADA: {row['TAREA'][:45]}")
                    st.rerun()

# ════════════════════════════════════════════════════════════════════════════
# KANBAN
# ════════════════════════════════════════════════════════════════════════════
elif pagina == "🗂️ Kanban":
    st.markdown("## 🗂️ Kanban")
    fa_k = st.selectbox("Filtrar área", ["Todas"] + sorted(df["ÁREA"].unique().tolist()), key="ka")
    filt_k = df if fa_k == "Todas" else df[df["ÁREA"] == fa_k]

    cols_k   = st.columns(4)
    estados_k = {"Sin iniciar": "🔴", "En curso": "🟡", "Bloqueada": "🟠", "Cerrada": "🟢"}

    for col, (est, icon) in zip(cols_k, estados_k.items()):
        sub = filt_k[filt_k["ESTADO"] == est]
        col.markdown(f"**{icon} {est}** `{len(sub)}`")
        col.markdown('<hr style="margin:5px 0;border-color:#eee">', unsafe_allow_html=True)
        for _, r in sub.iterrows():
            try:    av = int(float(r.get("AVANCE", "0") or 0))
            except: av = 0
            ac = AREA_COLOR.get(r["ÁREA"], "#888")
            col.markdown(f"""
            <div style="background:white;border:0.5px solid #ddd;border-radius:8px;
                 padding:10px;margin-bottom:8px;border-top:3px solid {ac}">
              <div style="font-size:.68rem;color:#999">{r['ID']} · {r['ÁREA']}</div>
              <div style="font-size:.8rem;font-weight:600;color:#1B3A5C;margin:3px 0">
                {r['TAREA'][:60]}{'…' if len(r['TAREA']) > 60 else ''}</div>
              <div style="font-size:.68rem;color:#aaa">{r['RESPONSABLE']}</div>
              <div style="height:3px;background:#eee;border-radius:2px;margin-top:5px">
                <div style="height:3px;background:{ac};width:{av}%;border-radius:2px"></div>
              </div>
              <div style="font-size:.65rem;color:#bbb;text-align:right">{av}%</div>
            </div>""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# PROYECTOS
# ════════════════════════════════════════════════════════════════════════════
elif pagina == "📊 Proyectos ZNI":
    st.markdown("## 📊 Proyectos ZNI — 35 Municipios")
    pdata = [
        ("Dispac","Cartagena del Chaira",1118,"CRÍTICA"),("Dispac","Puerto Leguizamo",1154,"CRÍTICA"),
        ("Dispac","Miraflores",1057,"CRÍTICA"),("Sunco Energy","Linea Colectora",1122,"CRÍTICA"),
        ("Dispac","Tierralta",976,"ALTA"),("Dispac","Tolima",858,"ALTA"),("Dispac","Guainia",798,"ALTA"),
        ("IPSE","Michikai Norte y Centro",749,"ALTA"),("IPSE","Michikai Sur y Noreste",660,"ALTA"),
        ("Dispac","Unguia",651,"ALTA"),("Gensa","Solano",503,"ALTA"),("Gensa","Amazonas 659",491,"ALTA"),
        ("Dispac","Cumaribo",485,"ALTA"),("Gensa","Puerto Carreño",355,"MEDIA"),
        ("Dispac","Morales",279,"MEDIA"),("Gensa","Fundacion",274,"MEDIA"),
        ("Dispac","Milan",254,"MEDIA"),("Dispac","Barrancas",235,"MEDIA"),
        ("Dispac","Riohacha",220,"MEDIA"),("Gensa","Amazonas 653",157,"MEDIA"),
        ("Cedenar","Puerto Asis",154,"MEDIA"),("Cedenar","Mocoa",151,"MEDIA"),
        ("CENS","Tibú",130,"MEDIA"),("Dispac","Paz de Ariporo",131,"MEDIA"),
        ("Dispac","Maicao",153,"MEDIA"),("Dispac","Zambrano",104,"BAJA"),
        ("Cedenar","Orito",105,"BAJA"),("Cedenar","Valle del Guamuez",78,"BAJA"),
        ("ISA","El Copey",58,"BAJA"),("CENS","El Carmen",58,"BAJA"),
        ("CENS","Sardinata",33,"BAJA"),("CENS","El Tarra",35,"BAJA"),
        ("CENS","Abrego",19,"BAJA"),("CENS","Teorama",3,"BAJA"),("CENS","Convención",4,"BAJA"),
    ]
    dp = pd.DataFrame(pdata, columns=["Proyecto","Municipio","Usuarios","Prioridad"])
    m1,m2,m3,m4 = st.columns(4)
    m1.metric("Proyectos","7"); m2.metric("Municipios","35")
    m3.metric("Usuarios totales",f"{dp['Usuarios'].sum():,}")
    m4.metric("Proyectos críticos",str(len(dp[dp["Prioridad"]=="CRÍTICA"])))
    st.markdown("---")
    cl, cr = st.columns(2)
    with cl:
        fig_p = px.bar(dp.sort_values("Usuarios",ascending=True).tail(15),
            x="Usuarios",y="Municipio",orientation="h",color="Prioridad",
            color_discrete_map={"CRÍTICA":"#C0392B","ALTA":"#E67E22","MEDIA":"#27AE60","BAJA":"#2E6DA4"},
            title="Top 15 municipios")
        fig_p.update_layout(height=380,margin=dict(l=0,r=0,t=40,b=0),
            plot_bgcolor="white",paper_bgcolor="white",font=dict(family="Arial",size=11))
        st.plotly_chart(fig_p,use_container_width=True)
    with cr:
        res = dp.groupby("Proyecto").agg(Municipios=("Municipio","count"),Usuarios=("Usuarios","sum")).reset_index().sort_values("Usuarios",ascending=False)
        st.dataframe(res,use_container_width=True,hide_index=True)

# ════════════════════════════════════════════════════════════════════════════
# NUEVA TAREA
# ════════════════════════════════════════════════════════════════════════════
elif pagina == "➕ Nueva Tarea":
    st.markdown("## ➕ Nueva Tarea")
    with st.form("nf", clear_on_submit=True):
        r1, r2 = st.columns(2)
        with r1:
            t    = st.text_area("Título / Descripción *", height=80)
            area = st.selectbox("Área *", list(AREA_COLOR.keys()))
            proy = st.text_input("Proyecto / Municipio", placeholder="Todos / Cartagena del Chaira...")
        with r2:
            resp = st.text_input("Responsable *")
            prio = st.selectbox("Prioridad *", PRIORIDADES)
            cat  = st.text_input("Categoría", placeholder="Jornada / Cartera / BBDD...")
            flim = st.date_input("Fecha límite", value=None)
            av_i = st.slider("Avance inicial %", 0, 100, 0)
        obs = st.text_area("Observaciones / contexto", height=70)

        if st.form_submit_button("✅ Crear tarea", type="primary"):
            if not t.strip():    st.error("El título es obligatorio.")
            elif not resp.strip(): st.error("El responsable es obligatorio.")
            else:
                df_new = st.session_state["df"].copy()
                nid = next_id(df_new)
                nr = {c: "" for c in COLS}
                nr.update({
                    "ID": nid, "ÁREA": area, "PROYECTO": proy or "Todos",
                    "TAREA": t.strip(), "CATEGORÍA": cat or "General",
                    "RESPONSABLE": resp.strip(),
                    "FECHA INICIO": datetime.now().strftime("%d/%m/%Y"),
                    "FECHA LÍMITE": flim.strftime("%d/%m/%Y") if flim else "",
                    "PRIORIDAD": prio, "AVANCE": str(av_i), "ESTADO": "Sin iniciar",
                    "OBSERVACIÓN": obs.strip(),
                })
                df_new = pd.concat([df_new, pd.DataFrame([nr])], ignore_index=True)
                commit(df_new, f"{nid} creada: {t[:45]} [{area}]")
                st.rerun()

# ════════════════════════════════════════════════════════════════════════════
# HISTORIAL
# ════════════════════════════════════════════════════════════════════════════
elif pagina == "📜 Historial":
    st.markdown("## 📜 Historial de actividad")
    hist = st.session_state.get("hist", [])
    if not hist:
        st.info("Sin actividad registrada aún en esta sesión.")
    else:
        for h in hist:
            st.markdown(
                f'<div style="padding:7px 0;border-bottom:0.5px solid #eee;font-size:.86rem">'
                f'<span style="color:#aaa;margin-right:10px">{h["ts"]}</span>{h["msg"]}</div>',
                unsafe_allow_html=True)
    if df is not None:
        st.markdown("---")
        ec1, ec2 = st.columns(2)
        with ec1:
            csv = df.to_csv(index=False).encode("utf-8-sig")
            st.download_button("⬇️ Descargar CSV", csv,
                               f"DISPOWER_Tareas_{datetime.now().strftime('%Y%m%d')}.csv", "text/csv")
        with ec2:
            buf2 = io.BytesIO()
            df.to_excel(buf2, index=False, engine="openpyxl")
            st.download_button("⬇️ Descargar Excel", buf2.getvalue(),
                f"DISPOWER_Tareas_{datetime.now().strftime('%Y%m%d')}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ════════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ════════════════════════════════════════════════════════════════════════════
elif pagina == "⚙️ Configuración":
    st.markdown("## Configuración y diagnóstico")

    cfg_now = {}
    try:
        cfg_now = st.secrets.get("graph", {})
    except Exception:
        pass
    token_ok = get_graph_token() is not None

    if token_ok:
        st.markdown('<div class="ok">Token Microsoft Graph OK. Azure AD bien configurado.</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="warn">No se pudo obtener el token. Verifica tenant_id, client_id y client_secret.</div>', unsafe_allow_html=True)

    st.markdown("### Paso 1 — Encontrar el archivo en tu OneDrive")
    st.markdown(
        '<div class="info">El token funciona pero el app necesita el drive_id y el item_id exactos de tu archivo. ' +
        "Haz clic para buscar automaticamente todos los Excel en tu OneDrive.</div>",
        unsafe_allow_html=True
    )

    if st.button("Buscar archivo Excel en mi OneDrive", type="primary", disabled=not token_ok):
        t2 = get_graph_token()
        found_items = []

        with st.spinner("Buscando archivos Excel..."):
            try:
                res1 = graph_get("me/drive/root/search(q='.xlsx')", t2)
                if res1 and res1.get("value"):
                    for item in res1["value"]:
                        name = item.get("name", "")
                        if ".xlsx" in name.lower() or ".xls" in name.lower():
                            found_items.append({
                                "Nombre": name,
                                "drive_id": item.get("parentReference", {}).get("driveId", ""),
                                "item_id": item.get("id", ""),
                                "Ruta": item.get("parentReference", {}).get("path", ""),
                                "Modificado": (item.get("lastModifiedDateTime") or "")[:10],
                            })
            except Exception as ex1:
                st.warning(f"Busqueda fallida: {ex1}")

            try:
                res2 = graph_get("me/drive/recent", t2)
                if res2 and res2.get("value"):
                    existing = {i["item_id"] for i in found_items}
                    for item in res2["value"]:
                        name = item.get("name", "")
                        iid  = item.get("id", "")
                        if (".xlsx" in name.lower() or ".xls" in name.lower()) and iid not in existing:
                            found_items.append({
                                "Nombre": name,
                                "drive_id": item.get("parentReference", {}).get("driveId", ""),
                                "item_id": iid,
                                "Ruta": item.get("parentReference", {}).get("path", ""),
                                "Modificado": (item.get("lastModifiedDateTime") or "")[:10],
                            })
            except Exception:
                pass

        if found_items:
            st.success(f"Se encontraron {len(found_items)} archivos Excel en tu OneDrive:")
            st.dataframe(pd.DataFrame(found_items), use_container_width=True, hide_index=True)

            best = next((i for i in found_items if "dispower" in i["Nombre"].lower()), found_items[0])
            st.markdown(f"**Archivo sugerido:** {best['Nombre']}")
            st.markdown("### Paso 2 — Copia este secrets.toml actualizado:")

            tid = cfg_now.get("tenant_id", "TU-TENANT-ID")
            cid = cfg_now.get("client_id", "TU-CLIENT-ID")
            sec = cfg_now.get("client_secret", "TU-SECRET")
            did = best["drive_id"]
            iid = best["item_id"]
            fnm = best["Nombre"]

            toml_txt = f"""[graph]
tenant_id     = "{tid}"
client_id     = "{cid}"
client_secret = "{sec}"
drive_id      = "{did}"
item_id       = "{iid}"
filename      = "{fnm}"
"""
            st.code(toml_txt, language="toml")
            st.markdown(
                '<div class="ok">Copia este bloque en .streamlit/secrets.toml, ' +
                "reemplaza el contenido anterior y reinicia el app con: streamlit run streamlit_app.py</div>",
                unsafe_allow_html=True
            )
        else:
            me_info = graph_get("me", t2)
            if me_info:
                user_str = f"{me_info.get('displayName','?')} ({me_info.get('userPrincipalName','?')})"
            else:
                user_str = "desconocido"
            st.error(f"No se encontraron archivos Excel para: {user_str}")
            st.markdown(
                '<div class="warn">Verifica que: el archivo este en el OneDrive correcto, ' +
                "el permiso Files.ReadWrite.All este concedido, y se haya pulsado Conceder consentimiento de administrador.</div>",
                unsafe_allow_html=True
            )

    st.markdown("---")
    st.markdown("### Estructura del secrets.toml (despues de ejecutar el diagnostico)")
    st.code("""[graph]
tenant_id     = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
client_id     = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
client_secret = "tu-secreto-aqui"
drive_id      = "b!xxxxxxxxxxxx"
item_id       = "xxxxxxxxxxxx"
filename      = "DISPOWER_Tareas_Streamlit.xlsx"
""", language="toml")

    st.markdown("**Permisos requeridos en Azure AD:**")
    st.markdown("- `Files.ReadWrite.All` Leer y escribir el Excel **CRITICO**")
    st.markdown("- `User.Read` Identificar el usuario para el diagnostico **Recomendado**")

    st.markdown("---")
    if st.button("Resetear sesion y caches"):
        st.session_state.clear()
        load_excel_from_graph.clear()
        get_cached_token.clear()
        st.rerun()
