"""
DISPOWER — Tablero Dirección Comercial ZNI
==========================================
Conectado a OneDrive for Business via Microsoft Graph API (client_credentials).

CONFIGURACION secrets.toml:
────────────────────────────
[graph]
tenant_id     = "tu-tenant-id"
client_id     = "tu-client-id"
client_secret = "tu-client-secret"
user_upn      = "javier_agudelo@dispower.co"   # email del dueño del archivo
filename      = "DISPOWER_Tareas_Streamlit.xlsx"

# Opcional — si ya conoces los IDs exactos (más rápido):
# drive_id  = "b!xxxx"
# item_id   = "xxxx"
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime
import io, warnings, time
import requests
from openpyxl import load_workbook

warnings.filterwarnings("ignore")

st.set_page_config(
    page_title="DISPOWER · ZNI",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Paleta ──────────────────────────────────────────────────────────────────
AREA_COLOR = {
    "Gestión Social": "#1A5C3A", "SAC": "#2E6DA4",
    "BI & Analítica": "#E67E22", "SUI & Subsidios": "#7D3C98", "Dirección": "#3D3D3D",
}
ESTADOS     = ["Sin iniciar", "En curso", "Cerrada", "Bloqueada", "Cancelada"]
PRIORIDADES = ["CRÍTICA", "ALTA", "MEDIA", "BAJA"]
ESTADO_COL  = {"Sin iniciar":"#E74C3C","En curso":"#F39C12","Cerrada":"#27AE60","Bloqueada":"#D85A30","Cancelada":"#888"}
ESTADO_BG   = {"Sin iniciar":"#FADBD8","En curso":"#FEF9E7","Cerrada":"#D6F0E3","Bloqueada":"#FDEBD0","Cancelada":"#F5F5F5"}
PRIO_COL    = {"CRÍTICA":"#C0392B","ALTA":"#E67E22","MEDIA":"#27AE60","BAJA":"#2E6DA4"}
PRIO_BG     = {"CRÍTICA":"#FADBD8","ALTA":"#FDEBD0","MEDIA":"#D6F0E3","BAJA":"#D6E8F7"}
COLS = ["ID","ÁREA","PROYECTO","TAREA","CATEGORÍA","RESPONSABLE",
        "FECHA INICIO","FECHA LÍMITE","PRIORIDAD","AVANCE","ESTADO",
        "OBSERVACIÓN","FECHA CIERRE","CERRADA POR"]

st.markdown("""<style>
.kpi{background:white;border-radius:10px;padding:16px 18px;
     border-left:5px solid #2E6DA4;box-shadow:0 1px 5px rgba(0,0,0,.07);margin-bottom:10px}
.kv{font-size:2rem;font-weight:700;margin:4px 0}
.kl{font-size:.7rem;color:#888;text-transform:uppercase;letter-spacing:.05em}
.badge{display:inline-block;border-radius:20px;padding:3px 10px;font-size:.72rem;font-weight:600}
.info{background:#EBF5FB;border-left:4px solid #2E6DA4;border-radius:4px;padding:10px 14px;margin:8px 0;font-size:.86rem;color:#1B3A5C}
.warn{background:#FEF9E7;border-left:4px solid #F39C12;border-radius:4px;padding:10px 14px;margin:8px 0;font-size:.86rem;color:#7D6608}
.ok  {background:#D6F0E3;border-left:4px solid #27AE60;border-radius:4px;padding:10px 14px;margin:8px 0;font-size:.86rem;color:#1A5C3A}
.err {background:#FADBD8;border-left:4px solid #E74C3C;border-radius:4px;padding:10px 14px;margin:8px 0;font-size:.86rem;color:#922B21}
</style>""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# MICROSOFT GRAPH — AUTENTICACIÓN
# ════════════════════════════════════════════════════════════════════════════

def get_cfg():
    try:    return st.secrets.get("graph", {})
    except: return {}

@st.cache_data(ttl=3200, show_spinner=False)
def get_token() -> str | None:
    cfg = get_cfg()
    tid = cfg.get("tenant_id","")
    cid = cfg.get("client_id","")
    sec = cfg.get("client_secret","")
    if not all([tid, cid, sec]):
        return None
    try:
        r = requests.post(
            f"https://login.microsoftonline.com/{tid}/oauth2/v2.0/token",
            data={"grant_type":"client_credentials","client_id":cid,
                  "client_secret":sec,"scope":"https://graph.microsoft.com/.default"},
            timeout=15
        )
        return r.json().get("access_token") if r.status_code == 200 else None
    except Exception:
        return None

def gget(path: str, token: str) -> dict | None:
    """GET Microsoft Graph."""
    try:
        r = requests.get(f"https://graph.microsoft.com/v1.0/{path}",
                         headers={"Authorization":f"Bearer {token}"}, timeout=25)
        if r.status_code == 200:
            return r.json()
        return None
    except Exception:
        return None

def gget_raw(path: str, token: str) -> bytes | None:
    """GET Microsoft Graph — retorna bytes (para descargar archivos)."""
    try:
        r = requests.get(f"https://graph.microsoft.com/v1.0/{path}",
                         headers={"Authorization":f"Bearer {token}"}, timeout=40)
        if r.status_code == 200 and len(r.content) > 500:
            return r.content
        return None
    except Exception:
        return None

def gput(path: str, token: str, data: bytes, content_type: str) -> bool:
    """PUT Microsoft Graph — sube archivo."""
    try:
        r = requests.put(
            f"https://graph.microsoft.com/v1.0/{path}",
            headers={"Authorization":f"Bearer {token}","Content-Type":content_type},
            data=data, timeout=60
        )
        return r.status_code in [200, 201]
    except Exception:
        return False

# ════════════════════════════════════════════════════════════════════════════
# LOCALIZACIÓN DEL ARCHIVO EN ONEDRIVE
# ════════════════════════════════════════════════════════════════════════════

@st.cache_data(ttl=86400, show_spinner=False)
def find_file(token: str) -> tuple[str | None, str | None]:
    """
    Encuentra el drive_id e item_id del Excel.
    Con client_credentials debemos usar /users/{upn}/drive, NO /me/drive.
    """
    cfg = get_cfg()

    # 1. IDs explícitos en secrets (más rápido y confiable)
    if cfg.get("drive_id") and cfg.get("item_id"):
        return cfg["drive_id"], cfg["item_id"]

    upn      = cfg.get("user_upn", "javier_agudelo@dispower.co")
    filename = cfg.get("filename", "DISPOWER_Tareas_Streamlit.xlsx")

    # 2. Buscar por nombre en el drive del usuario (usando UPN, no /me)
    try:
        fname_enc = filename.replace(" ", "%20").replace("&", "%26")
        result = gget(f"users/{upn}/drive/root/search(q='{fname_enc}')", token)
        if result and result.get("value"):
            for item in result["value"]:
                if filename.lower() in item.get("name","").lower():
                    did = item.get("parentReference",{}).get("driveId")
                    iid = item.get("id")
                    if did and iid:
                        return did, iid
    except Exception:
        pass

    # 3. Buscar en archivos recientes del usuario
    try:
        recent = gget(f"users/{upn}/drive/recent", token)
        if recent and recent.get("value"):
            for item in recent["value"]:
                if filename.lower() in item.get("name","").lower():
                    did = item.get("parentReference",{}).get("driveId")
                    iid = item.get("id")
                    if did and iid:
                        return did, iid
    except Exception:
        pass

    # 4. Listar todos los xlsx del drive y buscar por nombre
    try:
        search_all = gget(f"users/{upn}/drive/root/search(q='.xlsx')", token)
        if search_all and search_all.get("value"):
            for item in search_all["value"]:
                if filename.lower() in item.get("name","").lower():
                    did = item.get("parentReference",{}).get("driveId")
                    iid = item.get("id")
                    if did and iid:
                        return did, iid
    except Exception:
        pass

    return None, None

# ════════════════════════════════════════════════════════════════════════════
# LEER Y ESCRIBIR EXCEL
# ════════════════════════════════════════════════════════════════════════════

def parse_excel_bytes(raw: bytes) -> pd.DataFrame | None:
    try:
        xl = pd.ExcelFile(io.BytesIO(raw))
        target = next((s for s in xl.sheet_names
                       if any(k in s.lower() for k in ["tarea","cierre","task"])),
                      xl.sheet_names[0] if xl.sheet_names else None)
        if not target: return None
        df = xl.parse(target, dtype=str).fillna("")
        df.columns = [str(c).strip().replace("\n"," ") for c in df.columns]
        aliases = {
            "ÁREA":"ÁREA","AREA":"ÁREA",
            "PROYECTO / MUNICIPIO":"PROYECTO","PROYECTO /  MUNICIPIO":"PROYECTO","PROYECTO":"PROYECTO",
            "CATEGORÍA":"CATEGORÍA","CATEGORIA":"CATEGORÍA",
            "% AVANCE":"AVANCE","AVANCE %":"AVANCE","AVANCE":"AVANCE",
            "OBSERVACIÓN / BLOQUEO":"OBSERVACIÓN","OBSERVACION":"OBSERVACIÓN",
            "FECHA LÍMITE":"FECHA LÍMITE","FECHA LIMITE":"FECHA LÍMITE",
        }
        df = df.rename(columns={c: aliases.get(c.upper().strip(), c) for c in df.columns})
        if "ID" in df.columns:
            df = df[df["ID"].astype(str).str.match(r"T-\d+", na=False)].copy()
        for c in COLS:
            if c not in df.columns: df[c] = ""
        return df[COLS].reset_index(drop=True) if len(df) > 0 else None
    except Exception:
        return None

@st.cache_data(ttl=30, show_spinner=False)
def download_excel(token: str, drive_id: str, item_id: str) -> bytes | None:
    raw = gget_raw(f"drives/{drive_id}/items/{item_id}/content", token)
    if raw and raw[:4] == b'PK\x03\x04':
        return raw
    return None

def upload_excel(df: pd.DataFrame, token: str, drive_id: str, item_id: str,
                 original_raw: bytes | None) -> bool:
    """
    Sube el DataFrame actualizado al Excel en OneDrive.
    Preserva el formato original si está disponible.
    """
    try:
        if original_raw:
            wb = load_workbook(io.BytesIO(original_raw))
            target = next((s for s in wb.sheetnames
                           if any(k in s.lower() for k in ["tarea","cierre","task"])),
                          wb.sheetnames[0])
            ws = wb[target]
            # Localizar fila de encabezado
            hrow = None
            for ri in range(1, min(15, ws.max_row + 1)):
                vals = [str(ws.cell(ri,c).value or "").upper() for c in range(1, ws.max_column+1)]
                if any("TAREA" in v or "ÁREA" in v or "AREA" in v for v in vals):
                    hrow = ri; break
            if hrow:
                # Mapear columnas
                hdrs = [str(ws.cell(hrow,c).value or "").strip().replace("\n"," ")
                        for c in range(1, ws.max_column+1)]
                rev = {"PROYECTO / MUNICIPIO":"PROYECTO","PROYECTO /  MUNICIPIO":"PROYECTO",
                       "% AVANCE":"AVANCE","OBSERVACIÓN / BLOQUEO":"OBSERVACIÓN",
                       "FECHA LÍMITE":"FECHA LÍMITE"}
                col_idx = {}
                for ci, h in enumerate(hdrs, 1):
                    canon = rev.get(h.upper().strip(), h.upper().strip())
                    for c in COLS:
                        if c.upper() == canon:
                            col_idx[c] = ci; break
                # Limpiar filas de datos
                for rr in range(hrow+1, ws.max_row+10):
                    for cc in range(1, ws.max_column+1):
                        ws.cell(rr, cc).value = None
                # Escribir datos
                for ri2, (_, row) in enumerate(df.iterrows(), hrow+1):
                    for col_name, ci2 in col_idx.items():
                        ws.cell(ri2, ci2).value = str(row.get(col_name, ""))
                buf = io.BytesIO(); wb.save(buf)
                new_bytes = buf.getvalue()
            else:
                # Sin encabezado encontrado, guardar como nuevo
                buf = io.BytesIO()
                df.to_excel(buf, index=False, sheet_name="Tareas", engine="openpyxl")
                new_bytes = buf.getvalue()
        else:
            buf = io.BytesIO()
            df.to_excel(buf, index=False, sheet_name="Tareas", engine="openpyxl")
            new_bytes = buf.getvalue()

        ct = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        return gput(f"drives/{drive_id}/items/{item_id}/content", token, new_bytes, ct)

    except Exception as e:
        st.error(f"Error al subir: {e}")
        return False

# ════════════════════════════════════════════════════════════════════════════
# ESTADO DE SESIÓN E INICIALIZACIÓN
# ════════════════════════════════════════════════════════════════════════════
for k, v in [("df",None),("raw",None),("drive_id",None),("item_id",None),
              ("hist",[]),("loaded",False),("src","")]:
    if k not in st.session_state: st.session_state[k] = v

def add_hist(msg: str):
    st.session_state["hist"].insert(0, {"ts": datetime.now().strftime("%d/%m/%Y %H:%M"), "msg": msg})

def badge(text, fg, bg):
    return f'<span class="badge" style="background:{bg};color:{fg}">{text}</span>'
def estado_badge(e): return badge(e, ESTADO_COL.get(e,"#888"), ESTADO_BG.get(e,"#eee"))
def prio_badge(p):   return badge(p, PRIO_COL.get(p,"#888"),   PRIO_BG.get(p,"#eee"))

def next_id(df: pd.DataFrame) -> str:
    nums = pd.to_numeric(df["ID"].str.extract(r"T-(\d+)")[0], errors="coerce").dropna()
    return f"T-{int(nums.max())+1:03d}" if len(nums) else "T-001"

# ── Carga inicial ─────────────────────────────────────────────────────────
if not st.session_state["loaded"]:
    token = get_token()
    if token:
        did, iid = find_file(token)
        if did and iid:
            raw = download_excel(token, did, iid)
            if raw:
                parsed = parse_excel_bytes(raw)
                if parsed is not None:
                    st.session_state.update({"df":parsed,"raw":raw,"drive_id":did,
                                             "item_id":iid,"src":"onedrive","loaded":True})
    if not st.session_state["loaded"]:
        st.session_state.update({"loaded":True,"src":"sin_conexion"})

token    = get_token()
df       = st.session_state["df"]
drive_id = st.session_state["drive_id"]
item_id  = st.session_state["item_id"]
src      = st.session_state["src"]

def commit(df_new: pd.DataFrame, msg: str):
    """Guarda en sesión y sube a OneDrive."""
    st.session_state["df"] = df_new.copy()
    tok = get_token()
    did = st.session_state["drive_id"]
    iid = st.session_state["item_id"]
    raw = st.session_state["raw"]
    if tok and did and iid:
        with st.spinner("Guardando en OneDrive..."):
            ok = upload_excel(df_new, tok, did, iid, raw)
            # Invalidar caché de descarga
            download_excel.clear()
        if ok:
            add_hist(f"✅ {msg}")
            st.success("Guardado en OneDrive. El Excel se actualizó.")
        else:
            add_hist(f"⚠️ {msg} (error al subir)")
            st.warning("No se pudo guardar en OneDrive. El cambio queda en memoria.")
    else:
        add_hist(f"💾 {msg} (sin conexión OneDrive)")

# ════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### ⚡ DISPOWER")
    st.markdown("**Dirección Comercial ZNI**")
    st.markdown("---")
    pagina = st.radio("Navegación", [
        "🏠 Dashboard","📋 Lista de Tareas","🗂️ Kanban",
        "📊 Proyectos ZNI","➕ Nueva Tarea","📜 Historial","⚙️ Configuración",
    ])
    st.markdown("---")
    src_info = {
        "onedrive":    ("🟢","OneDrive conectado","Cambios se guardan automáticamente"),
        "sin_conexion":("🔴","Sin conexión","Ve a Configuración"),
        "":            ("⚪","Iniciando...",""),
    }
    em, lbl, sub = src_info.get(src, ("⚪","—",""))
    st.markdown(f"**{em} {lbl}**")
    if sub: st.caption(sub)
    if df is not None: st.caption(f"{len(df)} tareas")

    col_r1, col_r2 = st.columns(2)
    with col_r1:
        if st.button("🔄 Recargar", use_container_width=True):
            download_excel.clear()
            find_file.clear()
            st.session_state["loaded"] = False
            st.rerun()
    with col_r2:
        if df is not None:
            buf_dl = io.BytesIO()
            df.to_excel(buf_dl, index=False, sheet_name="Tareas", engine="openpyxl")
            st.download_button("⬇️ Excel", buf_dl.getvalue(),
                f"DISPOWER_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

# ─── Aviso si no hay datos ───────────────────────────────────────────────────
if df is None and pagina not in ["⚙️ Configuración"]:
    st.markdown('<div class="warn">No hay datos cargados. Ve a ⚙️ Configuración para diagnosticar la conexión, o sube el Excel manualmente.</div>', unsafe_allow_html=True)
    upl = st.file_uploader("Cargar Excel manualmente", type=["xlsx","xls"])
    if upl:
        parsed = parse_excel_bytes(upl.read())
        if parsed is not None:
            st.session_state["df"] = parsed
            st.session_state["src"] = "local"
            st.success(f"Cargadas {len(parsed)} tareas.")
            st.rerun()
    st.stop()

# ════════════════════════════════════════════════════════════════════════════
# DASHBOARD
# ════════════════════════════════════════════════════════════════════════════
if pagina == "🏠 Dashboard":
    st.markdown("## ⚡ Dashboard Ejecutivo")
    st.caption("Director: Javier Alejandro Agudelo · Plan Estratégico 6 Meses · ZNI / SISFV")

    total    = len(df)
    cerradas = (df["ESTADO"]=="Cerrada").sum()
    criticas = ((df["PRIORIDAD"]=="CRÍTICA")&(df["ESTADO"]!="Cerrada")).sum()
    en_curso = (df["ESTADO"]=="En curso").sum()
    bloq     = (df["ESTADO"]=="Bloqueada").sum()
    try:    avg_av = round(pd.to_numeric(df["AVANCE"],errors="coerce").fillna(0).mean())
    except: avg_av = 0

    cols_m = st.columns(6)
    for col,(val,lbl,sub2,cc) in zip(cols_m,[
        (total,    "Total tareas",    f"{cerradas} cerradas",           "#2E6DA4"),
        (criticas, "Críticas pend.",  "atención urgente",               "#C0392B"),
        (en_curso, "En curso",        "activas",                        "#E67E22"),
        (bloq,     "Bloqueadas",      "requieren decisión",             "#D85A30"),
        (f"{avg_av}%","Avance global","promedio plan",                  "#27AE60"),
        (cerradas, "Completadas",     f"{round(cerradas/total*100) if total else 0}%","#27AE60"),
    ]):
        col.markdown(
            f'<div class="kpi" style="border-left-color:{cc}">'
            f'<div class="kl">{lbl}</div><div class="kv" style="color:{cc}">{val}</div>'
            f'<div style="font-size:.7rem;color:#aaa">{sub2}</div></div>',
            unsafe_allow_html=True)

    st.markdown("---")
    cl,cr = st.columns([2,1])
    with cl:
        st.markdown("**Avance por área**")
        rows_a=[]
        for area in AREA_COLOR:
            s=df[df["ÁREA"]==area]
            if not len(s): continue
            try:    av=round(pd.to_numeric(s["AVANCE"],errors="coerce").fillna(0).mean())
            except: av=0
            rows_a.append({"Área":area,"Avance":av})
        if rows_a:
            fig=go.Figure(go.Bar(
                x=[r["Área"] for r in rows_a],y=[r["Avance"] for r in rows_a],
                marker_color=[AREA_COLOR[r["Área"]] for r in rows_a],opacity=.85,
                text=[f"{r['Avance']}%" for r in rows_a],textposition="outside"))
            fig.add_hline(y=100,line_dash="dash",line_color="#27AE60",line_width=1.5)
            fig.update_layout(height=260,margin=dict(l=0,r=0,t=10,b=0),
                plot_bgcolor="white",paper_bgcolor="white",
                yaxis=dict(title="% avance",range=[0,120]),showlegend=False,
                font=dict(family="Arial",size=11))
            st.plotly_chart(fig,use_container_width=True)
    with cr:
        st.markdown("**Estados**")
        ec=df["ESTADO"].value_counts()
        if len(ec):
            fig2=go.Figure(go.Pie(
                labels=ec.index.tolist(),values=ec.values.tolist(),
                marker_colors=[ESTADO_COL.get(e,"#888") for e in ec.index],
                hole=.5,textinfo="value+percent"))
            fig2.update_layout(height=260,margin=dict(l=0,r=0,t=10,b=10),
                paper_bgcolor="white",showlegend=True,
                legend=dict(orientation="h",y=-.15),font=dict(family="Arial",size=11))
            st.plotly_chart(fig2,use_container_width=True)

    # Resumen por área
    st.markdown("---")
    res_cols=st.columns(len(AREA_COLOR))
    for col,(area,color) in zip(res_cols,AREA_COLOR.items()):
        s=df[df["ÁREA"]==area]
        cerr=(s["ESTADO"]=="Cerrada").sum()
        crit=((s["PRIORIDAD"]=="CRÍTICA")&(s["ESTADO"]!="Cerrada")).sum()
        try:    av2=round(pd.to_numeric(s["AVANCE"],errors="coerce").fillna(0).mean())
        except: av2=0
        alerta=""
        if crit: alerta=f'<div style="font-size:.65rem;color:#C0392B">⚠ {crit} críticas</div>'
        col.markdown(
            f'<div style="background:white;border-radius:8px;padding:12px;border-top:3px solid {color};'
            f'box-shadow:0 1px 4px rgba(0,0,0,.07);text-align:center">'
            f'<div style="font-size:.68rem;color:{color};font-weight:600;margin-bottom:4px">{area}</div>'
            f'<div style="font-size:1.5rem;font-weight:700;color:{color}">{av2}%</div>'
            f'<div style="font-size:.68rem;color:#aaa">{len(s)} tareas · {cerr} cerradas</div>'
            f'{alerta}</div>', unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("**Tareas críticas pendientes:**")
    crit_df=df[(df["PRIORIDAD"]=="CRÍTICA")&(df["ESTADO"].isin(["Sin iniciar","Bloqueada"]))]
    if len(crit_df):
        for _,r in crit_df.iterrows():
            c1,c2,c3,c4=st.columns([1,5,2,2])
            c1.markdown(f'<b style="color:#C0392B">{r["ID"]}</b>',unsafe_allow_html=True)
            c2.write(r["TAREA"][:80])
            c3.markdown(f'<small style="color:{AREA_COLOR.get(r["ÁREA"],"#888")}">{r["ÁREA"]}</small>',unsafe_allow_html=True)
            c4.markdown(estado_badge(r["ESTADO"]),unsafe_allow_html=True)
    else:
        st.success("Sin tareas críticas pendientes.")

# ════════════════════════════════════════════════════════════════════════════
# LISTA DE TAREAS
# ════════════════════════════════════════════════════════════════════════════
elif pagina == "📋 Lista de Tareas":
    st.markdown("## 📋 Lista de Tareas")
    fa_v=st.selectbox("Área",["Todas"]+sorted(df["ÁREA"].unique().tolist()),key="la")
    c1f,c2f,c3f=st.columns(3)
    with c1f: fe_v=st.selectbox("Estado",["Todos"]+ESTADOS,key="le")
    with c2f: fp_v=st.selectbox("Prioridad",["Todas"]+PRIORIDADES,key="lp")
    with c3f: fb_v=st.text_input("Buscar",placeholder="Tarea, proyecto, responsable...",key="lb")

    filt=df.copy()
    if fa_v!="Todas": filt=filt[filt["ÁREA"]==fa_v]
    if fe_v!="Todos": filt=filt[filt["ESTADO"]==fe_v]
    if fp_v!="Todas": filt=filt[filt["PRIORIDAD"]==fp_v]
    if fb_v:
        mask=(filt["TAREA"].str.contains(fb_v,case=False,na=False)|
              filt["PROYECTO"].str.contains(fb_v,case=False,na=False)|
              filt["RESPONSABLE"].str.contains(fb_v,case=False,na=False))
        filt=filt[mask]

    st.caption(f"{len(filt)} de {len(df)} tareas")
    st.markdown("---")

    for idx,row in filt.iterrows():
        try:    av_v=int(float(row.get("AVANCE","0") or 0))
        except: av_v=0
        area_c=AREA_COLOR.get(row["ÁREA"],"#888")

        with st.expander(f"**{row['ID']}** — {row['TAREA'][:85]}{'…' if len(row['TAREA'])>85 else ''}"):
            h1c,h2c,h3c,h4c=st.columns(4)
            h1c.markdown(estado_badge(row["ESTADO"]),unsafe_allow_html=True)
            h2c.markdown(prio_badge(row["PRIORIDAD"]),unsafe_allow_html=True)
            h3c.markdown(f'<span style="color:{area_c};font-weight:600;font-size:.8rem">{row["ÁREA"]}</span>',unsafe_allow_html=True)
            h4c.markdown(f'<span style="color:#888;font-size:.8rem">{row["RESPONSABLE"]}</span>',unsafe_allow_html=True)

            prog_c="#27AE60" if av_v>=80 else ("#F39C12" if av_v>=40 else "#E74C3C")
            st.markdown(
                f'<div style="height:4px;background:#eee;border-radius:2px;margin:6px 0">'
                f'<div style="height:4px;background:{prog_c};width:{av_v}%;border-radius:2px"></div></div>'
                f'<div style="font-size:.72rem;color:#aaa">Avance: {av_v}% · {row["PROYECTO"]} · {row["CATEGORÍA"]}</div>',
                unsafe_allow_html=True)
            if row.get("OBSERVACIÓN","").strip():
                st.markdown(f'<div class="warn">{row["OBSERVACIÓN"]}</div>',unsafe_allow_html=True)

            st.markdown("---")
            e1,e2,e3=st.columns(3)
            with e1:
                new_est=st.selectbox("Estado",ESTADOS,
                    index=ESTADOS.index(row["ESTADO"]) if row["ESTADO"] in ESTADOS else 0,
                    key=f"est_{row['ID']}")
            with e2:
                new_av=st.slider("Avance %",0,100,av_v,key=f"av_{row['ID']}")
            with e3:
                new_resp=st.text_input("Responsable",value=row["RESPONSABLE"],key=f"rsp_{row['ID']}")

            new_obs=st.text_area("Observación",value=row.get("OBSERVACIÓN",""),height=65,key=f"obs_{row['ID']}")
            try:    fl_val=datetime.strptime(row.get("FECHA LÍMITE",""),"%d/%m/%Y").date() if row.get("FECHA LÍMITE","").strip() else None
            except: fl_val=None
            new_fl=st.date_input("Fecha límite",value=fl_val,key=f"fl_{row['ID']}")

            b1,b2,b3,b4=st.columns(4)
            with b1:
                if st.button("💾 Guardar",key=f"s_{row['ID']}"):
                    df_new=st.session_state["df"].copy()
                    m=df_new["ID"]==row["ID"]
                    df_new.loc[m,"ESTADO"]=new_est
                    df_new.loc[m,"AVANCE"]=str(new_av)
                    df_new.loc[m,"RESPONSABLE"]=new_resp
                    df_new.loc[m,"OBSERVACIÓN"]=new_obs
                    if new_fl: df_new.loc[m,"FECHA LÍMITE"]=new_fl.strftime("%d/%m/%Y")
                    commit(df_new,f"{row['ID']} actualizada: {new_est}, {new_av}%")
                    st.rerun()
            with b2:
                if row["ESTADO"]!="Cerrada":
                    if st.button("✅ Cerrar",key=f"c_{row['ID']}"):
                        df_new=st.session_state["df"].copy()
                        m=df_new["ID"]==row["ID"]
                        df_new.loc[m,"ESTADO"]="Cerrada"
                        df_new.loc[m,"AVANCE"]="100"
                        df_new.loc[m,"FECHA CIERRE"]=datetime.now().strftime("%d/%m/%Y")
                        commit(df_new,f"{row['ID']} CERRADA: {row['TAREA'][:40]}")
                        st.rerun()
                else:
                    if st.button("🔄 Reabrir",key=f"r_{row['ID']}"):
                        df_new=st.session_state["df"].copy()
                        m=df_new["ID"]==row["ID"]
                        df_new.loc[m,"ESTADO"]="En curso"
                        df_new.loc[m,"FECHA CIERRE"]=""
                        commit(df_new,f"{row['ID']} reabierta")
                        st.rerun()
            with b3:
                if st.button("🟠 Bloquear",key=f"bl_{row['ID']}"):
                    df_new=st.session_state["df"].copy()
                    df_new.loc[df_new["ID"]==row["ID"],"ESTADO"]="Bloqueada"
                    commit(df_new,f"{row['ID']} BLOQUEADA")
                    st.rerun()
            with b4:
                if st.button("🗑️ Eliminar",key=f"d_{row['ID']}"):
                    df_new=st.session_state["df"].copy()
                    df_new=df_new[df_new["ID"]!=row["ID"]].reset_index(drop=True)
                    commit(df_new,f"{row['ID']} ELIMINADA")
                    st.rerun()

# ════════════════════════════════════════════════════════════════════════════
# KANBAN
# ════════════════════════════════════════════════════════════════════════════
elif pagina == "🗂️ Kanban":
    st.markdown("## 🗂️ Kanban")
    fa_k=st.selectbox("Filtrar área",["Todas"]+sorted(df["ÁREA"].unique().tolist()),key="ka")
    filt_k=df if fa_k=="Todas" else df[df["ÁREA"]==fa_k]
    cols_k=st.columns(4)
    for col,(est,icon) in zip(cols_k,{"Sin iniciar":"🔴","En curso":"🟡","Bloqueada":"🟠","Cerrada":"🟢"}.items()):
        sub_k=filt_k[filt_k["ESTADO"]==est]
        col.markdown(f"**{icon} {est}** `{len(sub_k)}`")
        col.markdown('<hr style="margin:5px 0;border-color:#eee">',unsafe_allow_html=True)
        for _,r in sub_k.iterrows():
            try:    av=int(float(r.get("AVANCE","0") or 0))
            except: av=0
            ac=AREA_COLOR.get(r["ÁREA"],"#888")
            col.markdown(f"""
            <div style="background:white;border:0.5px solid #ddd;border-radius:8px;
                 padding:10px;margin-bottom:8px;border-top:3px solid {ac}">
              <div style="font-size:.68rem;color:#999">{r['ID']} · {r['ÁREA']}</div>
              <div style="font-size:.8rem;font-weight:600;color:#1B3A5C;margin:3px 0">
                {r['TAREA'][:60]}{'…' if len(r['TAREA'])>60 else ''}</div>
              <div style="font-size:.68rem;color:#aaa">{r['RESPONSABLE']}</div>
              <div style="height:3px;background:#eee;border-radius:2px;margin-top:5px">
                <div style="height:3px;background:{ac};width:{av}%;border-radius:2px"></div></div>
              <div style="font-size:.65rem;color:#bbb;text-align:right">{av}%</div>
            </div>""",unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# PROYECTOS
# ════════════════════════════════════════════════════════════════════════════
elif pagina == "📊 Proyectos ZNI":
    st.markdown("## 📊 Proyectos ZNI — 35 Municipios")
    pdata=[
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
        ("CENS","Tibú",130,"MEDIA"),("Dispac","Paz de Ariporo",131,"MEDIA"),("Dispac","Maicao",153,"MEDIA"),
        ("Dispac","Zambrano",104,"BAJA"),("Cedenar","Orito",105,"BAJA"),
        ("Cedenar","Valle del Guamuez",78,"BAJA"),("ISA","El Copey",58,"BAJA"),
        ("CENS","El Carmen",58,"BAJA"),("CENS","Sardinata",33,"BAJA"),("CENS","El Tarra",35,"BAJA"),
        ("CENS","Abrego",19,"BAJA"),("CENS","Teorama",3,"BAJA"),("CENS","Convención",4,"BAJA"),
    ]
    dp=pd.DataFrame(pdata,columns=["Proyecto","Municipio","Usuarios","Prioridad"])
    m1,m2,m3,m4=st.columns(4)
    m1.metric("Proyectos","7"); m2.metric("Municipios","35")
    m3.metric("Usuarios totales",f"{dp['Usuarios'].sum():,}")
    m4.metric("Críticos",str(len(dp[dp["Prioridad"]=="CRÍTICA"])))
    cl2,cr2=st.columns(2)
    with cl2:
        fig_p=px.bar(dp.sort_values("Usuarios",ascending=True).tail(15),
            x="Usuarios",y="Municipio",orientation="h",color="Prioridad",
            color_discrete_map={"CRÍTICA":"#C0392B","ALTA":"#E67E22","MEDIA":"#27AE60","BAJA":"#2E6DA4"},
            title="Top 15 municipios")
        fig_p.update_layout(height=380,margin=dict(l=0,r=0,t=40,b=0),
            plot_bgcolor="white",paper_bgcolor="white",font=dict(family="Arial",size=11))
        st.plotly_chart(fig_p,use_container_width=True)
    with cr2:
        res=dp.groupby("Proyecto").agg(Municipios=("Municipio","count"),Usuarios=("Usuarios","sum")).reset_index().sort_values("Usuarios",ascending=False)
        st.dataframe(res,use_container_width=True,hide_index=True)

# ════════════════════════════════════════════════════════════════════════════
# NUEVA TAREA
# ════════════════════════════════════════════════════════════════════════════
elif pagina == "➕ Nueva Tarea":
    st.markdown("## ➕ Nueva Tarea")
    with st.form("nf",clear_on_submit=True):
        r1,r2=st.columns(2)
        with r1:
            t_v   = st.text_area("Título / Descripción *",height=80)
            area_v= st.selectbox("Área *",list(AREA_COLOR.keys()))
            proy_v= st.text_input("Proyecto / Municipio",placeholder="Todos / Cartagena del Chaira...")
        with r2:
            resp_v= st.text_input("Responsable *")
            prio_v= st.selectbox("Prioridad *",PRIORIDADES)
            cat_v = st.text_input("Categoría",placeholder="Jornada / Cartera / BBDD...")
            flim_v= st.date_input("Fecha límite",value=None)
            av_i  = st.slider("Avance inicial %",0,100,0)
        obs_v=st.text_area("Observaciones",height=70)
        if st.form_submit_button("✅ Crear tarea",type="primary"):
            if not t_v.strip():    st.error("El título es obligatorio.")
            elif not resp_v.strip(): st.error("El responsable es obligatorio.")
            else:
                df_new=st.session_state["df"].copy()
                nid=next_id(df_new)
                nr={c:"" for c in COLS}
                nr.update({"ID":nid,"ÁREA":area_v,"PROYECTO":proy_v or "Todos","TAREA":t_v.strip(),
                    "CATEGORÍA":cat_v or "General","RESPONSABLE":resp_v.strip(),
                    "FECHA INICIO":datetime.now().strftime("%d/%m/%Y"),
                    "FECHA LÍMITE":flim_v.strftime("%d/%m/%Y") if flim_v else "",
                    "PRIORIDAD":prio_v,"AVANCE":str(av_i),"ESTADO":"Sin iniciar","OBSERVACIÓN":obs_v.strip()})
                df_new=pd.concat([df_new,pd.DataFrame([nr])],ignore_index=True)
                commit(df_new,f"{nid} creada: {t_v[:45]} [{area_v}]")
                st.rerun()

# ════════════════════════════════════════════════════════════════════════════
# HISTORIAL
# ════════════════════════════════════════════════════════════════════════════
elif pagina == "📜 Historial":
    st.markdown("## 📜 Historial")
    hist=st.session_state.get("hist",[])
    if not hist: st.info("Sin actividad registrada aún.")
    else:
        for h in hist:
            st.markdown(
                f'<div style="padding:7px 0;border-bottom:0.5px solid #eee;font-size:.86rem">'
                f'<span style="color:#aaa;margin-right:10px">{h["ts"]}</span>{h["msg"]}</div>',
                unsafe_allow_html=True)
    st.markdown("---")
    ec1,ec2=st.columns(2)
    with ec1:
        csv=df.to_csv(index=False).encode("utf-8-sig")
        st.download_button("⬇️ CSV",csv,f"DISPOWER_{datetime.now().strftime('%Y%m%d')}.csv","text/csv")
    with ec2:
        buf2=io.BytesIO(); df.to_excel(buf2,index=False,engine="openpyxl")
        st.download_button("⬇️ Excel",buf2.getvalue(),
            f"DISPOWER_{datetime.now().strftime('%Y%m%d')}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ════════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN — DIAGNÓSTICO COMPLETO
# ════════════════════════════════════════════════════════════════════════════
elif pagina == "⚙️ Configuración":
    st.markdown("## Configuración y diagnóstico")
    cfg_now=get_cfg()

    # Estado del token
    tok2=get_token()
    if tok2:
        st.markdown('<div class="ok">Token Microsoft Graph OK. Azure AD bien configurado.</div>',unsafe_allow_html=True)
    else:
        st.markdown('<div class="err">No se pudo obtener el token. Verifica tenant_id, client_id y client_secret.</div>',unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### Diagnosticar y encontrar el archivo")
    st.markdown(
        '<div class="info">Con <b>client_credentials</b> (flujo máquina-a-máquina) no se puede usar /me/drive. '
        "El app usará <b>/users/{user_upn}/drive</b> para acceder al OneDrive de Javier. "
        "Asegúrate de tener <b>user_upn</b> en secrets.toml y el permiso <b>Files.ReadWrite.All</b> concedido.</div>",
        unsafe_allow_html=True)

    upn_input = st.text_input(
        "Email / UPN del propietario del archivo",
        value=cfg_now.get("user_upn","javier_agudelo@dispower.co"),
        help="El correo corporativo de quien tiene el Excel en su OneDrive"
    )
    fname_input = st.text_input(
        "Nombre del archivo",
        value=cfg_now.get("filename","DISPOWER_Tareas_Streamlit.xlsx")
    )

    if st.button("Buscar archivo con el UPN indicado", type="primary", disabled=not tok2):
        found=[]
        with st.spinner(f"Buscando en el OneDrive de {upn_input}..."):
            # Búsqueda principal
            try:
                r1=gget(f"users/{upn_input}/drive/root/search(q='.xlsx')",tok2)
                if r1 and r1.get("value"):
                    for item in r1["value"]:
                        nm=item.get("name","")
                        if ".xlsx" in nm.lower() or ".xls" in nm.lower():
                            found.append({
                                "Nombre":nm,
                                "drive_id":item.get("parentReference",{}).get("driveId",""),
                                "item_id":item.get("id",""),
                                "Modificado":(item.get("lastModifiedDateTime") or "")[:10],
                                "Ruta":item.get("parentReference",{}).get("path",""),
                            })
            except Exception as ex:
                st.warning(f"Búsqueda fallida: {ex}")

            # Recientes
            try:
                r2=gget(f"users/{upn_input}/drive/recent",tok2)
                if r2 and r2.get("value"):
                    ex_ids={i["item_id"] for i in found}
                    for item in r2["value"]:
                        nm=item.get("name",""); iid=item.get("id","")
                        if (".xlsx" in nm.lower() or ".xls" in nm.lower()) and iid not in ex_ids:
                            found.append({
                                "Nombre":nm,
                                "drive_id":item.get("parentReference",{}).get("driveId",""),
                                "item_id":iid,
                                "Modificado":(item.get("lastModifiedDateTime") or "")[:10],
                                "Ruta":item.get("parentReference",{}).get("path",""),
                            })
            except Exception:
                pass

        if found:
            st.success(f"Se encontraron {len(found)} archivos Excel:")
            st.dataframe(pd.DataFrame(found),use_container_width=True,hide_index=True)
            best=next((i for i in found if fname_input.lower() in i["Nombre"].lower()),found[0])
            st.markdown(f"**Archivo seleccionado:** `{best['Nombre']}`")
            st.markdown("### Copia este secrets.toml y reinicia el app:")
            toml_out=(
                f"[graph]\n"
                f"tenant_id     = \"{cfg_now.get('tenant_id','')}\"\n"
                f"client_id     = \"{cfg_now.get('client_id','')}\"\n"
                f"client_secret = \"{cfg_now.get('client_secret','')}\"\n"
                f"user_upn      = \"{upn_input}\"\n"
                f"filename      = \"{best['Nombre']}\"\n"
                f"drive_id      = \"{best['drive_id']}\"\n"
                f"item_id       = \"{best['item_id']}\"\n"
            )
            st.code(toml_out,language="toml")
            st.markdown('<div class="ok">Copia este bloque en .streamlit/secrets.toml y reinicia: <b>streamlit run streamlit_app.py</b></div>',unsafe_allow_html=True)
        else:
            st.error(f"No se encontraron archivos Excel para {upn_input}")
            st.markdown('<div class="warn">Posibles causas:<br>'
                '1. El UPN no es correcto (debe ser el email corporativo exacto)<br>'
                '2. El permiso <b>Files.ReadWrite.All</b> no tiene consentimiento de administrador<br>'
                '3. El archivo no está en el OneDrive de ese usuario (puede estar en SharePoint de sitio)<br>'
                '4. El directorio no tiene usuarios licenciados con OneDrive</div>',unsafe_allow_html=True)

            # Intentar listar usuarios para diagnóstico
            st.markdown("**Diagnóstico adicional — verificar acceso:**")
            try:
                usr=gget(f"users/{upn_input}",tok2)
                if usr and usr.get("id"):
                    st.success(f"Usuario encontrado: {usr.get('displayName')} ({usr.get('userPrincipalName')})")
                    st.info("El usuario existe pero no se pudo acceder a su OneDrive. Verifica el permiso Files.ReadWrite.All con consentimiento de administrador.")
                else:
                    st.error("El usuario no se encontró en el directorio. Verifica el UPN.")
            except Exception as ex3:
                st.error(f"Error al buscar usuario: {ex3}")

    st.markdown("---")
    st.markdown("### secrets.toml requerido")
    st.code("""[graph]
tenant_id     = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
client_id     = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
client_secret = "tu-secreto-aqui"
user_upn      = "javier_agudelo@dispower.co"
filename      = "DISPOWER_Tareas_Streamlit.xlsx"

# Estos se agregan automaticamente despues del diagnostico:
# drive_id  = "b!xxxx"
# item_id   = "xxxx"
""",language="toml")

    st.markdown("### Permiso requerido en Azure AD")
    st.markdown("En **Permisos de API → Permisos de aplicación**, agrega:")
    st.markdown("- `Files.ReadWrite.All` — **CRÍTICO**: leer y escribir archivos de cualquier usuario")
    st.markdown("- Luego pulsar **Conceder consentimiento de administrador para [tu organización]**")

    st.markdown("---")
    if st.button("Resetear sesion y caches"):
        st.session_state.clear()
        get_token.clear()
        find_file.clear()
        download_excel.clear()
        st.rerun()
