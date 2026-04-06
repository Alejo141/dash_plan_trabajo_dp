"""
DISPOWER — Tablero de Seguimiento | Dirección Comercial ZNI
============================================================
Lee y escribe el Excel directamente en OneDrive for Business
usando Microsoft Graph API (client_credentials).

REPO EN GITHUB:
  dash_plan_trabajo.py            ← este archivo
  DISPOWER_Tareas_Streamlit.xlsx  ← Excel (respaldo local)
  requirements.txt
  .streamlit/secrets.toml         ← credenciales (en Streamlit Cloud Secrets)

secrets.toml:
──────────────────────────────────────────────────────────
[graph]
tenant_id     = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
client_id     = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
client_secret = "tu-secreto"
user_upn      = "javier_agudelo@dispower.co"
filename      = "DISPOWER_Tareas_Streamlit.xlsx"
drive_id      = ""   # se obtiene desde la página Configuración
item_id       = ""   # se obtiene desde la página Configuración
──────────────────────────────────────────────────────────
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime
import io, os, warnings
import requests
from openpyxl import load_workbook

warnings.filterwarnings("ignore")

st.set_page_config(
    page_title="DISPOWER · Dirección Comercial ZNI",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ══════════════════════════════════════════════════════════════════════════
# CONSTANTES
# ══════════════════════════════════════════════════════════════════════════
EXCEL_LOCAL = "DISPOWER_Tareas_Streamlit.xlsx"
HEADER_ROW  = 3   # encabezado en fila 4 del Excel (0-indexed = 3)

AREA_COLOR = {
    "Gestión Social":  "#1A5C3A",
    "SAC":             "#2E6DA4",
    "BI & Analítica":  "#E67E22",
    "SUI & Subsidios": "#7D3C98",
    "Dirección":       "#3D3D3D",
}
AREA_BG = {
    "Gestión Social":  "#D6F0E3", "SAC": "#D6E8F7",
    "BI & Analítica":  "#FDEBD0", "SUI & Subsidios": "#EAD8F5", "Dirección": "#F5F5F5",
}
ESTADOS     = ["Sin iniciar", "En curso", "Cerrada", "Bloqueada", "Cancelada"]
PRIORIDADES = ["CRÍTICA", "ALTA", "MEDIA", "BAJA"]
ESTADO_COL  = {"Sin iniciar":"#E74C3C","En curso":"#F39C12","Cerrada":"#27AE60","Bloqueada":"#D85A30","Cancelada":"#888888"}
ESTADO_BG   = {"Sin iniciar":"#FADBD8","En curso":"#FEF9E7","Cerrada":"#D6F0E3","Bloqueada":"#FDEBD0","Cancelada":"#F5F5F5"}
PRIO_COL    = {"CRÍTICA":"#C0392B","ALTA":"#E67E22","MEDIA":"#27AE60","BAJA":"#2E6DA4"}
PRIO_BG     = {"CRÍTICA":"#FADBD8","ALTA":"#FDEBD0","MEDIA":"#D6F0E3","BAJA":"#D6E8F7"}

COLS = [
    "ID","ÁREA","PROYECTO","TAREA","CATEGORÍA","RESPONSABLE",
    "FECHA INICIO","FECHA LÍMITE","PRIORIDAD","AVANCE","ESTADO",
    "OBSERVACIÓN","FECHA CIERRE","CERRADA POR"
]

st.markdown("""<style>
.kpi{background:white;border-radius:10px;padding:16px 18px;
     border-left:5px solid #2E6DA4;box-shadow:0 1px 5px rgba(0,0,0,.08);margin-bottom:10px}
.kv{font-size:2rem;font-weight:700;margin:4px 0}
.kl{font-size:.7rem;color:#888;text-transform:uppercase;letter-spacing:.05em}
.ks{font-size:.7rem;color:#aaa;margin-top:2px}
.badge{display:inline-block;border-radius:20px;padding:3px 10px;font-size:.72rem;font-weight:600}
.binfo{background:#EBF5FB;border-left:4px solid #2E6DA4;border-radius:4px;padding:10px 14px;margin:8px 0;font-size:.86rem;color:#1B3A5C}
.bok  {background:#D6F0E3;border-left:4px solid #27AE60;border-radius:4px;padding:10px 14px;margin:8px 0;font-size:.86rem;color:#1A5C3A}
.bwarn{background:#FEF9E7;border-left:4px solid #F39C12;border-radius:4px;padding:10px 14px;margin:8px 0;font-size:.86rem;color:#7D6608}
.berr {background:#FADBD8;border-left:4px solid #E74C3C;border-radius:4px;padding:10px 14px;margin:8px 0;font-size:.86rem;color:#922B21}
</style>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════
# GRAPH API
# ══════════════════════════════════════════════════════════════════════════
def get_cfg():
    try:    return dict(st.secrets.get("graph", {}))
    except: return {}

@st.cache_data(ttl=3200, show_spinner=False)
def get_token():
    cfg = get_cfg()
    tid, cid, sec = cfg.get("tenant_id",""), cfg.get("client_id",""), cfg.get("client_secret","")
    if not all([tid, cid, sec]): return None
    try:
        r = requests.post(
            f"https://login.microsoftonline.com/{tid}/oauth2/v2.0/token",
            data={"grant_type":"client_credentials","client_id":cid,
                  "client_secret":sec,"scope":"https://graph.microsoft.com/.default"},
            timeout=15
        )
        return r.json().get("access_token") if r.status_code == 200 else None
    except Exception: return None

def g_get(path, token):
    try:
        r = requests.get(f"https://graph.microsoft.com/v1.0/{path}",
                         headers={"Authorization":f"Bearer {token}"}, timeout=25)
        return r.json() if r.status_code == 200 else None
    except Exception: return None

def g_get_bytes(path, token):
    try:
        r = requests.get(f"https://graph.microsoft.com/v1.0/{path}",
                         headers={"Authorization":f"Bearer {token}"}, timeout=40)
        return r.content if r.status_code == 200 and len(r.content) > 100 else None
    except Exception: return None

def g_put(path, token, data):
    try:
        ct = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        r = requests.put(f"https://graph.microsoft.com/v1.0/{path}",
                         headers={"Authorization":f"Bearer {token}","Content-Type":ct},
                         data=data, timeout=60)
        return r.status_code in [200, 201]
    except Exception: return False

@st.cache_data(ttl=3600*8, show_spinner=False)
def resolve_file(token):
    cfg      = get_cfg()
    drive_id = cfg.get("drive_id","")
    item_id  = cfg.get("item_id","")
    if drive_id and item_id: return drive_id, item_id

    upn      = cfg.get("user_upn","")
    filename = cfg.get("filename", EXCEL_LOCAL)
    if not upn: return None, None

    for ep in [
        f"users/{upn}/drive/root/search(q='{filename}')",
        f"users/{upn}/drive/root/search(q='.xlsx')",
        f"users/{upn}/drive/recent",
    ]:
        try:
            res = g_get(ep, token)
            if res and res.get("value"):
                for item in res["value"]:
                    if filename.lower() in item.get("name","").lower():
                        did = item.get("parentReference",{}).get("driveId")
                        iid = item.get("id")
                        if did and iid: return did, iid
        except Exception: continue
    return None, None

# ══════════════════════════════════════════════════════════════════════════
# EXCEL — LEER Y ESCRIBIR
# ══════════════════════════════════════════════════════════════════════════
def parse_excel(raw):
    try:
        xl = pd.ExcelFile(io.BytesIO(raw))
        target = next((s for s in xl.sheet_names if any(k in s.lower() for k in ["tarea","cierre","task"])), xl.sheet_names[0])
        df = xl.parse(target, header=HEADER_ROW, dtype=str).fillna("")
        df.columns = [str(c).strip().replace("\n"," ") for c in df.columns]
        aliases = {
            "ÁREA":"ÁREA","AREA":"ÁREA","CATEGORÍA":"CATEGORÍA","CATEGORIA":"CATEGORÍA",
            "OBSERVACIÓN":"OBSERVACIÓN","OBSERVACION":"OBSERVACIÓN",
            "OBSERVACIÓN / BLOQUEO":"OBSERVACIÓN",
            "FECHA LÍMITE":"FECHA LÍMITE","FECHA LIMITE":"FECHA LÍMITE",
            "% AVANCE":"AVANCE","AVANCE %":"AVANCE",
        }
        df = df.rename(columns={c: aliases.get(c.upper().strip(), c) for c in df.columns})
        if "ID" not in df.columns: return None
        df = df[df["ID"].astype(str).str.match(r"T-\d+", na=False)].copy()
        for c in COLS:
            if c not in df.columns: df[c] = ""
        df = df[COLS].reset_index(drop=True)
        return df if len(df) > 0 else None
    except Exception as e:
        return None

@st.cache_data(ttl=30, show_spinner=False)
def download_excel(token, drive_id, item_id):
    raw = g_get_bytes(f"drives/{drive_id}/items/{item_id}/content", token)
    return raw if (raw and raw[:4] == b"PK\x03\x04") else None

def build_bytes(df, original_raw):
    if original_raw:
        try:
            wb = load_workbook(io.BytesIO(original_raw))
            target = next((s for s in wb.sheetnames if any(k in s.lower() for k in ["tarea","cierre","task"])), wb.sheetnames[0])
            ws = wb[target]
            hrow = None
            for ri in range(1, min(10, ws.max_row+1)):
                vals = [str(ws.cell(ri,c).value or "").upper().strip() for c in range(1, ws.max_column+1)]
                if "ID" in vals and "TAREA" in vals:
                    hrow = ri; break
            if hrow:
                hdrs = [str(ws.cell(hrow,c).value or "").strip().replace("\n"," ") for c in range(1, ws.max_column+1)]
                rev  = {"OBSERVACIÓN / BLOQUEO":"OBSERVACIÓN","% AVANCE":"AVANCE","FECHA LÍMITE":"FECHA LÍMITE"}
                col_idx = {}
                for ci, h in enumerate(hdrs, 1):
                    canon = rev.get(h.upper().strip(), h.upper().strip())
                    for c in COLS:
                        if c.upper() == canon:
                            col_idx[c] = ci; break
                for rr in range(hrow+1, ws.max_row+10):
                    for cc in range(1, ws.max_column+1): ws.cell(rr,cc).value = None
                for ri2, (_, row) in enumerate(df.iterrows(), hrow+1):
                    for col_name, ci2 in col_idx.items():
                        val = row.get(col_name,"")
                        ws.cell(ri2, ci2).value = str(val) if val != "" else None
                buf = io.BytesIO(); wb.save(buf); return buf.getvalue()
        except Exception: pass
    buf2 = io.BytesIO(); df.to_excel(buf2, index=False, sheet_name="Tareas", engine="openpyxl"); return buf2.getvalue()

# ══════════════════════════════════════════════════════════════════════════
# ESTADO DE SESIÓN
# ══════════════════════════════════════════════════════════════════════════
for k, v in [("df",None),("raw",None),("drive_id",None),("item_id",None),("hist",[]),("loaded",False),("src","")]:
    if k not in st.session_state: st.session_state[k] = v

def badge(text, fg, bg): return f'<span class="badge" style="background:{bg};color:{fg}">{text}</span>'
def ebadge(e): return badge(e, ESTADO_COL.get(e,"#888"), ESTADO_BG.get(e,"#eee"))
def pbadge(p): return badge(p, PRIO_COL.get(p,"#888"),   PRIO_BG.get(p,"#eee"))
def add_hist(msg): st.session_state["hist"].insert(0,{"ts":datetime.now().strftime("%d/%m/%Y %H:%M"),"msg":msg})
def next_id(df):
    nums = pd.to_numeric(df["ID"].str.extract(r"T-(\d+)")[0], errors="coerce").dropna()
    return f"T-{int(nums.max())+1:03d}" if len(nums) else "T-001"

def commit(df_new, msg):
    st.session_state["df"] = df_new.copy()
    tok = get_token()
    did = st.session_state["drive_id"]
    iid = st.session_state["item_id"]
    raw = st.session_state["raw"]
    if tok and did and iid:
        nb = build_bytes(df_new, raw)
        with st.spinner("Guardando en OneDrive..."):
            ok = g_put(f"drives/{did}/items/{iid}/content", tok, nb)
        download_excel.clear()
        add_hist(("✅ " if ok else "⚠️ ") + msg)
        st.toast("Guardado en OneDrive ✅" if ok else "Guardado localmente (error al subir)", icon="✅" if ok else "⚠️")
    else:
        add_hist(f"💾 {msg}")
        st.toast("Guardado en sesión", icon="💾")

# ── Carga inicial ──────────────────────────────────────────────────────────
if not st.session_state["loaded"]:
    tok_i = get_token()
    if tok_i:
        did_i, iid_i = resolve_file(tok_i)
        if did_i and iid_i:
            raw_i = download_excel(tok_i, did_i, iid_i)
            if raw_i:
                df_i = parse_excel(raw_i)
                if df_i is not None:
                    st.session_state.update({"df":df_i,"raw":raw_i,"drive_id":did_i,"item_id":iid_i,"src":"onedrive"})
    if st.session_state["df"] is None and os.path.exists(EXCEL_LOCAL):
        with open(EXCEL_LOCAL,"rb") as f: raw_l = f.read()
        df_l = parse_excel(raw_l)
        if df_l is not None:
            st.session_state.update({"df":df_l,"raw":raw_l,"src":"local"})
    st.session_state["loaded"] = True

df  = st.session_state["df"]
src = st.session_state["src"]

# ══════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### ⚡ DISPOWER")
    st.markdown("**Dirección Comercial ZNI**")
    st.markdown("---")
    pagina = st.radio("Navegación",[
        "🏠 Dashboard","📋 Lista de Tareas","🗂️ Kanban",
        "📊 Proyectos ZNI","➕ Nueva Tarea","📜 Historial","⚙️ Configuración",
    ])
    st.markdown("---")
    icons  = {"onedrive":"🟢","local":"🟡","":"🔴"}
    labels = {"onedrive":"OneDrive conectado","local":"Excel local","":"Sin conexión"}
    subs   = {"onedrive":"Cambios en tiempo real","local":"Conecta OneDrive en ⚙️","":"Ve a ⚙️ Configuración"}
    st.markdown(f"**{icons.get(src,'🔴')} {labels.get(src,'—')}**")
    st.caption(subs.get(src,""))
    if df is not None: st.caption(f"{len(df)} tareas")
    c1,c2 = st.columns(2)
    with c1:
        if st.button("🔄 Recargar", use_container_width=True):
            download_excel.clear(); resolve_file.clear()
            st.session_state["loaded"] = False; st.rerun()
    with c2:
        if df is not None:
            buf_dl = io.BytesIO(); df.to_excel(buf_dl, index=False, sheet_name="Tareas", engine="openpyxl")
            st.download_button("⬇️ Excel", buf_dl.getvalue(),
                f"DISPOWER_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

# ── Sin datos ─────────────────────────────────────────────────────────────
if df is None and pagina != "⚙️ Configuración":
    st.markdown('<div class="bwarn">No hay datos. Sube el Excel manualmente o configura OneDrive en ⚙️ Configuración.</div>', unsafe_allow_html=True)
    upl = st.file_uploader("📂 Subir DISPOWER_Tareas_Streamlit.xlsx", type=["xlsx","xls"])
    if upl:
        raw_u = upl.read(); df_u = parse_excel(raw_u)
        if df_u is not None:
            st.session_state.update({"df":df_u,"raw":raw_u,"src":"local"})
            st.success(f"✅ {len(df_u)} tareas cargadas."); st.rerun()
        else: st.error("No se pudo leer el archivo.")
    st.stop()

# ══════════════════════════════════════════════════════════════════════════
# DASHBOARD
# ══════════════════════════════════════════════════════════════════════════
if pagina == "🏠 Dashboard":
    st.markdown("## ⚡ Dashboard Ejecutivo")
    st.caption("Director: Javier Alejandro Agudelo · Plan Estratégico 6 Meses · ZNI / SISFV")

    total    = len(df)
    cerradas = int((df["ESTADO"]=="Cerrada").sum())
    criticas = int(((df["PRIORIDAD"]=="CRÍTICA")&(df["ESTADO"]!="Cerrada")).sum())
    en_curso = int((df["ESTADO"]=="En curso").sum())
    bloq     = int((df["ESTADO"]=="Bloqueada").sum())
    try:    avg_av = round(pd.to_numeric(df["AVANCE"],errors="coerce").fillna(0).mean())
    except: avg_av = 0

    for col,(val,lbl,sub2,cc) in zip(st.columns(6),[
        (total,"Total tareas",f"{cerradas} cerradas","#2E6DA4"),
        (criticas,"Críticas pend.","requieren atención","#C0392B"),
        (en_curso,"En curso","activas ahora","#E67E22"),
        (bloq,"Bloqueadas","requieren decisión","#D85A30"),
        (f"{avg_av}%","Avance global","promedio plan","#27AE60"),
        (cerradas,"Completadas",f"{round(cerradas/total*100) if total else 0}% del plan","#27AE60"),
    ]):
        col.markdown(f'<div class="kpi" style="border-left-color:{cc}"><div class="kl">{lbl}</div>'
                     f'<div class="kv" style="color:{cc}">{val}</div><div class="ks">{sub2}</div></div>',
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
            fig.add_hline(y=100,line_dash="dash",line_color="#27AE60",line_width=1.5,annotation_text="Meta")
            fig.update_layout(height=260,margin=dict(l=0,r=0,t=10,b=0),
                plot_bgcolor="white",paper_bgcolor="white",
                yaxis=dict(title="% avance",range=[0,120]),showlegend=False,font=dict(family="Arial",size=11))
            st.plotly_chart(fig,use_container_width=True)
    with cr:
        st.markdown("**Estados**")
        ec=df["ESTADO"].value_counts()
        if len(ec):
            fig2=go.Figure(go.Pie(labels=ec.index.tolist(),values=ec.values.tolist(),
                marker_colors=[ESTADO_COL.get(e,"#888") for e in ec.index],hole=.5,textinfo="value+percent"))
            fig2.update_layout(height=260,margin=dict(l=0,r=0,t=10,b=10),paper_bgcolor="white",
                showlegend=True,legend=dict(orientation="h",y=-.15),font=dict(family="Arial",size=11))
            st.plotly_chart(fig2,use_container_width=True)

    st.markdown("---")
    area_cols=st.columns(len(AREA_COLOR))
    for col,(area,color) in zip(area_cols,AREA_COLOR.items()):
        s=df[df["ÁREA"]==area]; cerr=int((s["ESTADO"]=="Cerrada").sum())
        crit=int(((s["PRIORIDAD"]=="CRÍTICA")&(s["ESTADO"]!="Cerrada")).sum())
        try:    av2=round(pd.to_numeric(s["AVANCE"],errors="coerce").fillna(0).mean())
        except: av2=0
        alerta=f'<div style="font-size:.64rem;color:#C0392B">⚠ {crit} críticas</div>' if crit else ""
        col.markdown(f'<div style="background:white;border-radius:8px;padding:12px;border-top:3px solid {color};'
                     f'box-shadow:0 1px 4px rgba(0,0,0,.07);text-align:center">'
                     f'<div style="font-size:.68rem;color:{color};font-weight:600;margin-bottom:4px">{area}</div>'
                     f'<div style="font-size:1.6rem;font-weight:700;color:{color}">{av2}%</div>'
                     f'<div style="font-size:.68rem;color:#aaa">{len(s)} tareas · {cerr} cerradas</div>'
                     f'{alerta}</div>',unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("**⚠️ Tareas críticas pendientes:**")
    crit_df=df[(df["PRIORIDAD"]=="CRÍTICA")&(df["ESTADO"].isin(["Sin iniciar","Bloqueada"]))]
    if len(crit_df):
        for _,r in crit_df.iterrows():
            c1,c2,c3,c4=st.columns([1,5,2,2])
            c1.markdown(f'<b style="color:#C0392B">{r["ID"]}</b>',unsafe_allow_html=True)
            c2.write(r["TAREA"][:80]+("…" if len(r["TAREA"])>80 else ""))
            c3.markdown(f'<small style="color:{AREA_COLOR.get(r["ÁREA"],"#888")}">{r["ÁREA"]}</small>',unsafe_allow_html=True)
            c4.markdown(ebadge(r["ESTADO"]),unsafe_allow_html=True)
    else:
        st.success("Sin tareas críticas pendientes.")

# ══════════════════════════════════════════════════════════════════════════
# LISTA DE TAREAS
# ══════════════════════════════════════════════════════════════════════════
elif pagina == "📋 Lista de Tareas":
    st.markdown("## 📋 Lista de Tareas")
    fa=st.selectbox("Área",["Todas"]+sorted(df["ÁREA"].unique().tolist()),key="la")
    c1f,c2f,c3f=st.columns(3)
    with c1f: fe=st.selectbox("Estado",["Todos"]+ESTADOS,key="le")
    with c2f: fp=st.selectbox("Prioridad",["Todas"]+PRIORIDADES,key="lp")
    with c3f: fb=st.text_input("Buscar",placeholder="Tarea, proyecto, responsable...",key="lb")
    filt=df.copy()
    if fa!="Todas":  filt=filt[filt["ÁREA"]==fa]
    if fe!="Todos":  filt=filt[filt["ESTADO"]==fe]
    if fp!="Todas":  filt=filt[filt["PRIORIDAD"]==fp]
    if fb:
        mask=(filt["TAREA"].str.contains(fb,case=False,na=False)|
              filt["PROYECTO"].str.contains(fb,case=False,na=False)|
              filt["RESPONSABLE"].str.contains(fb,case=False,na=False))
        filt=filt[mask]
    st.caption(f"{len(filt)} de {len(df)} tareas")
    st.markdown("---")

    for pos,(idx,row) in enumerate(filt.iterrows()):
        try:    av_v=int(float(row.get("AVANCE","0") or 0))
        except: av_v=0
        area_c=AREA_COLOR.get(row["ÁREA"],"#888")
        wk=f"t{pos}"  # clave única por posición — evita DuplicateElementKey

        with st.expander(f"**{row['ID']}** — {row['TAREA'][:85]}{'…' if len(row['TAREA'])>85 else ''}",key=f"exp_{wk}"):
            h1,h2,h3,h4=st.columns(4)
            h1.markdown(ebadge(row["ESTADO"]),unsafe_allow_html=True)
            h2.markdown(pbadge(row["PRIORIDAD"]),unsafe_allow_html=True)
            h3.markdown(f'<span style="color:{area_c};font-weight:600;font-size:.8rem">{row["ÁREA"]}</span>',unsafe_allow_html=True)
            h4.markdown(f'<span style="color:#888;font-size:.8rem">{row["RESPONSABLE"]}</span>',unsafe_allow_html=True)

            prog_c="#27AE60" if av_v>=80 else ("#F39C12" if av_v>=40 else "#E74C3C")
            st.markdown(
                f'<div style="height:4px;background:#eee;border-radius:2px;margin:6px 0">'
                f'<div style="height:4px;background:{prog_c};width:{av_v}%;border-radius:2px"></div></div>'
                f'<div style="font-size:.72rem;color:#aaa">Avance: {av_v}% · {row["PROYECTO"]} · {row["CATEGORÍA"]}</div>',
                unsafe_allow_html=True)
            if row.get("OBSERVACIÓN","").strip():
                st.markdown(f'<div class="bwarn">{row["OBSERVACIÓN"]}</div>',unsafe_allow_html=True)

            st.markdown("---")
            e1,e2,e3=st.columns(3)
            with e1:
                new_est=st.selectbox("Estado",ESTADOS,
                    index=ESTADOS.index(row["ESTADO"]) if row["ESTADO"] in ESTADOS else 0,key=f"est_{wk}")
            with e2: new_av=st.slider("Avance %",0,100,av_v,key=f"av_{wk}")
            with e3: new_resp=st.text_input("Responsable",value=row["RESPONSABLE"],key=f"rsp_{wk}")
            new_obs=st.text_area("Observación",value=row.get("OBSERVACIÓN",""),height=65,key=f"obs_{wk}")
            try:    fl_val=datetime.strptime(row.get("FECHA LÍMITE",""),"%d/%m/%Y").date() if row.get("FECHA LÍMITE","").strip() else None
            except: fl_val=None
            new_fl=st.date_input("Fecha límite",value=fl_val,key=f"fl_{wk}")

            b1,b2,b3,b4=st.columns(4)
            with b1:
                if st.button("💾 Guardar",key=f"sv_{wk}"):
                    df_new=st.session_state["df"].copy(); m=df_new["ID"]==row["ID"]
                    df_new.loc[m,"ESTADO"]=new_est; df_new.loc[m,"AVANCE"]=str(new_av)
                    df_new.loc[m,"RESPONSABLE"]=new_resp; df_new.loc[m,"OBSERVACIÓN"]=new_obs
                    if new_fl: df_new.loc[m,"FECHA LÍMITE"]=new_fl.strftime("%d/%m/%Y")
                    commit(df_new,f"{row['ID']} actualizada: {new_est}, {new_av}%"); st.rerun()
            with b2:
                if row["ESTADO"]!="Cerrada":
                    if st.button("✅ Cerrar",key=f"cl_{wk}"):
                        df_new=st.session_state["df"].copy(); m=df_new["ID"]==row["ID"]
                        df_new.loc[m,"ESTADO"]="Cerrada"; df_new.loc[m,"AVANCE"]="100"
                        df_new.loc[m,"FECHA CIERRE"]=datetime.now().strftime("%d/%m/%Y")
                        commit(df_new,f"{row['ID']} CERRADA"); st.rerun()
                else:
                    if st.button("🔄 Reabrir",key=f"ro_{wk}"):
                        df_new=st.session_state["df"].copy(); m=df_new["ID"]==row["ID"]
                        df_new.loc[m,"ESTADO"]="En curso"; df_new.loc[m,"FECHA CIERRE"]=""
                        commit(df_new,f"{row['ID']} reabierta"); st.rerun()
            with b3:
                if st.button("🟠 Bloquear",key=f"bl_{wk}"):
                    df_new=st.session_state["df"].copy()
                    df_new.loc[df_new["ID"]==row["ID"],"ESTADO"]="Bloqueada"
                    commit(df_new,f"{row['ID']} BLOQUEADA"); st.rerun()
            with b4:
                if st.button("🗑️ Eliminar",key=f"dl_{wk}"):
                    df_new=st.session_state["df"].copy()
                    df_new=df_new[df_new["ID"]!=row["ID"]].reset_index(drop=True)
                    commit(df_new,f"{row['ID']} ELIMINADA"); st.rerun()

# ══════════════════════════════════════════════════════════════════════════
# KANBAN
# ══════════════════════════════════════════════════════════════════════════
elif pagina == "🗂️ Kanban":
    st.markdown("## 🗂️ Tablero Kanban")
    fa_k=st.selectbox("Filtrar área",["Todas"]+sorted(df["ÁREA"].unique().tolist()),key="ka")
    fk=df if fa_k=="Todas" else df[df["ÁREA"]==fa_k]
    for col,(est,icon) in zip(st.columns(4),{"Sin iniciar":"🔴","En curso":"🟡","Bloqueada":"🟠","Cerrada":"🟢"}.items()):
        sub=fk[fk["ESTADO"]==est]
        col.markdown(f"**{icon} {est}** `{len(sub)}`")
        col.markdown('<hr style="margin:5px 0;border-color:#eee">',unsafe_allow_html=True)
        for _,r in sub.iterrows():
            try:    avk=int(float(r.get("AVANCE","0") or 0))
            except: avk=0
            ac=AREA_COLOR.get(r["ÁREA"],"#888")
            col.markdown(f'<div style="background:white;border:0.5px solid #e0e0e0;border-radius:8px;'
                         f'padding:10px;margin-bottom:8px;border-top:3px solid {ac}">'
                         f'<div style="font-size:.68rem;color:#999">{r["ID"]} · {r["ÁREA"]}</div>'
                         f'<div style="font-size:.8rem;font-weight:600;color:#1B3A5C;margin:3px 0 4px">'
                         f'{r["TAREA"][:60]}{"…" if len(r["TAREA"])>60 else ""}</div>'
                         f'<div style="font-size:.68rem;color:#aaa">{r["RESPONSABLE"]}</div>'
                         f'<div style="height:3px;background:#eee;border-radius:2px;margin-top:5px">'
                         f'<div style="height:3px;background:{ac};width:{avk}%;border-radius:2px"></div></div>'
                         f'<div style="font-size:.64rem;color:#bbb;text-align:right">{avk}%</div></div>',
                         unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════
# PROYECTOS ZNI
# ══════════════════════════════════════════════════════════════════════════
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
    st.markdown("---")
    cl2,cr2=st.columns(2)
    with cl2:
        fig_p=px.bar(dp.sort_values("Usuarios",ascending=True).tail(15),x="Usuarios",y="Municipio",orientation="h",
            color="Prioridad",color_discrete_map={"CRÍTICA":"#C0392B","ALTA":"#E67E22","MEDIA":"#27AE60","BAJA":"#2E6DA4"},
            title="Top 15 por usuarios")
        fig_p.update_layout(height=380,margin=dict(l=0,r=0,t=40,b=0),plot_bgcolor="white",paper_bgcolor="white",font=dict(family="Arial",size=11))
        st.plotly_chart(fig_p,use_container_width=True)
    with cr2:
        res=dp.groupby("Proyecto").agg(Municipios=("Municipio","count"),Usuarios=("Usuarios","sum")).reset_index().sort_values("Usuarios",ascending=False)
        st.dataframe(res,use_container_width=True,hide_index=True)

# ══════════════════════════════════════════════════════════════════════════
# NUEVA TAREA
# ══════════════════════════════════════════════════════════════════════════
elif pagina == "➕ Nueva Tarea":
    st.markdown("## ➕ Nueva Tarea")
    with st.form("form_nueva",clear_on_submit=True):
        r1,r2=st.columns(2)
        with r1:
            nt_t=st.text_area("Título / Descripción *",height=80)
            nt_a=st.selectbox("Área *",list(AREA_COLOR.keys()))
            nt_p=st.text_input("Proyecto / Municipio",placeholder="Todos / Cartagena del Chaira...")
        with r2:
            nt_r=st.text_input("Responsable *")
            nt_pr=st.selectbox("Prioridad *",PRIORIDADES)
            nt_c=st.text_input("Categoría",placeholder="Jornada / Cartera / BBDD...")
            nt_fl=st.date_input("Fecha límite",value=None)
            nt_av=st.slider("Avance inicial %",0,100,0)
        nt_obs=st.text_area("Observaciones",height=70)
        if st.form_submit_button("✅ Crear tarea",type="primary"):
            if not nt_t.strip(): st.error("El título es obligatorio.")
            elif not nt_r.strip(): st.error("El responsable es obligatorio.")
            else:
                df_new=st.session_state["df"].copy(); nid=next_id(df_new)
                nr={c:"" for c in COLS}
                nr.update({"ID":nid,"ÁREA":nt_a,"PROYECTO":nt_p or "Todos","TAREA":nt_t.strip(),
                    "CATEGORÍA":nt_c or "General","RESPONSABLE":nt_r.strip(),
                    "FECHA INICIO":datetime.now().strftime("%d/%m/%Y"),
                    "FECHA LÍMITE":nt_fl.strftime("%d/%m/%Y") if nt_fl else "",
                    "PRIORIDAD":nt_pr,"AVANCE":str(nt_av),"ESTADO":"Sin iniciar","OBSERVACIÓN":nt_obs.strip()})
                df_new=pd.concat([df_new,pd.DataFrame([nr])],ignore_index=True)
                commit(df_new,f"{nid} creada: {nt_t[:45]} [{nt_a}]")
                st.success(f"✅ Tarea {nid} creada."); st.rerun()

# ══════════════════════════════════════════════════════════════════════════
# HISTORIAL
# ══════════════════════════════════════════════════════════════════════════
elif pagina == "📜 Historial":
    st.markdown("## 📜 Historial")
    hist=st.session_state.get("hist",[])
    if not hist: st.info("Sin actividad registrada aún.")
    else:
        for h in hist:
            st.markdown(f'<div style="padding:7px 0;border-bottom:0.5px solid #eee;font-size:.86rem">'
                        f'<span style="color:#aaa;margin-right:10px">{h["ts"]}</span>{h["msg"]}</div>',
                        unsafe_allow_html=True)
    st.markdown("---")
    ec1,ec2=st.columns(2)
    with ec1:
        st.download_button("⬇️ CSV",df.to_csv(index=False).encode("utf-8-sig"),
            f"DISPOWER_{datetime.now().strftime('%Y%m%d')}.csv","text/csv")
    with ec2:
        buf2=io.BytesIO(); df.to_excel(buf2,index=False,engine="openpyxl")
        st.download_button("⬇️ Excel",buf2.getvalue(),
            f"DISPOWER_{datetime.now().strftime('%Y%m%d')}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ══════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════════════════
elif pagina == "⚙️ Configuración":
    st.markdown("## ⚙️ Configuración — Conexión OneDrive")
    cfg_now=get_cfg(); tok2=get_token()
    if tok2: st.markdown('<div class="bok">✅ Token Microsoft Graph OK.</div>',unsafe_allow_html=True)
    else:    st.markdown('<div class="berr">❌ No se pudo obtener el token.</div>',unsafe_allow_html=True)
    if st.session_state["drive_id"]:
        st.markdown('<div class="bok">✅ Archivo encontrado en OneDrive.</div>',unsafe_allow_html=True)
    elif tok2:
        st.markdown('<div class="bwarn">⚠️ Archivo no localizado. Ejecuta el diagnóstico.</div>',unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### Buscar el archivo Excel en OneDrive")
    st.markdown('<div class="binfo">Se usa <code>/users/{email}/drive</code> porque el token es de aplicación (client_credentials), no de usuario.</div>',unsafe_allow_html=True)
    upn_in  =st.text_input("Email del propietario del archivo",value=cfg_now.get("user_upn","javier_agudelo@dispower.co"))
    fname_in=st.text_input("Nombre del archivo",value=cfg_now.get("filename",EXCEL_LOCAL))

    if st.button("🔍 Buscar en OneDrive",type="primary",disabled=not tok2):
        found=[]
        with st.spinner(f"Buscando en el OneDrive de {upn_in}..."):
            for ep in [f"users/{upn_in}/drive/root/search(q='.xlsx')",f"users/{upn_in}/drive/recent"]:
                try:
                    res=g_get(ep,tok2)
                    if res and res.get("value"):
                        ex_ids={i["item_id"] for i in found}
                        for item in res["value"]:
                            nm=item.get("name",""); iid=item.get("id","")
                            if (".xlsx" in nm.lower() or ".xls" in nm.lower()) and iid not in ex_ids:
                                found.append({"Nombre":nm,
                                    "drive_id":item.get("parentReference",{}).get("driveId",""),
                                    "item_id":iid,
                                    "Modificado":(item.get("lastModifiedDateTime") or "")[:10]})
                except Exception: continue
        if found:
            st.success(f"✅ {len(found)} archivos encontrados:")
            st.dataframe(pd.DataFrame(found),use_container_width=True,hide_index=True)
            best=next((i for i in found if fname_in.lower() in i["Nombre"].lower()),found[0])
            st.markdown(f"**Archivo sugerido:** `{best['Nombre']}`")
            st.markdown("### Copia este secrets.toml y reinicia el app:")
            st.code(
                f"[graph]\n"
                f"tenant_id     = \"{cfg_now.get('tenant_id','')}\"\n"
                f"client_id     = \"{cfg_now.get('client_id','')}\"\n"
                f"client_secret = \"{cfg_now.get('client_secret','')}\"\n"
                f"user_upn      = \"{upn_in}\"\n"
                f"filename      = \"{best['Nombre']}\"\n"
                f"drive_id      = \"{best['drive_id']}\"\n"
                f"item_id       = \"{best['item_id']}\"\n",
                language="toml")
            st.markdown('<div class="bok">Copia el bloque en <b>.streamlit/secrets.toml</b> (o en Secrets de Streamlit Cloud) y reinicia.</div>',unsafe_allow_html=True)
        else:
            try:
                usr=g_get(f"users/{upn_in}",tok2)
                if usr and usr.get("id"):
                    st.error(f"Usuario {usr.get('displayName')} existe pero sin Excel accesible.")
                    st.markdown('<div class="bwarn">Verifica que <b>Files.ReadWrite.All</b> tenga consentimiento de administrador.</div>',unsafe_allow_html=True)
                else: st.error(f"Usuario {upn_in} no encontrado. Verifica el email.")
            except Exception: st.error("Error al conectar con Graph API.")

    st.markdown("---")
    st.markdown("### secrets.toml mínimo")
    st.code("""[graph]
tenant_id     = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
client_id     = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
client_secret = "tu-secreto"
user_upn      = "javier_agudelo@dispower.co"
filename      = "DISPOWER_Tareas_Streamlit.xlsx"
drive_id      = ""   # completar con el diagnóstico
item_id       = ""   # completar con el diagnóstico
""",language="toml")
    st.markdown("**Permisos Azure AD (Permisos de aplicación):** `Files.ReadWrite.All` + consentimiento de administrador")

    st.markdown("---")
    if st.button("🗑️ Resetear sesión"):
        st.session_state.clear(); get_token.clear(); resolve_file.clear(); download_excel.clear(); st.rerun()
