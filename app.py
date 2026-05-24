# ============================================================
#  AUTOCARES ALEGRE — Sistema de Gestión de Servicios Extras
#  Versión DEFINITIVA 6.0
# ============================================================

import streamlit as st
import pandas as pd
import json
from supabase import create_client
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date, datetime, time
import io

st.set_page_config(
    page_title="Autocares Alegre",
    page_icon="🚌",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# ============================================================
# CONSTANTES
# ============================================================
PASSWORD_ADMIN = "alegre2026"

# Orden y precios de servicios fijos
SERVICIOS_GRID = [
    "Pilotos", "Pescadores", "Logista-Saludes",
    "Colegio", "Transfer",
    "Boda Ida", "Boda Regreso 1", "Boda Regreso 2"
]
# Abreviaciones para cabecera de la cuadrícula
ABREV = {
    "Pilotos":          "PIL",
    "Pescadores":       "PES",
    "Logista-Saludes":  "LOG",
    "Colegio":          "COL",
    "Transfer":         "TRA",
    "Boda Ida":         "B.Ida",
    "Boda Regreso 1":   "BR.1",
    "Boda Regreso 2":   "BR.2",
}
PRECIOS_FIJOS = {
    "Pilotos":          15.00,
    "Pescadores":       15.00,
    "Logista-Saludes":  15.00,
    "Colegio":          25.00,
    "Transfer":         20.00,
    "Boda Ida":         40.00,
    "Boda Regreso 1":   40.00,
    "Boda Regreso 2":   40.00,
}
N_FILAS_GRID = 5

PRECIO_DIETA       = 15.00
PRECIO_HORA_SEMANA = 10.00
PRECIO_HORA_FINDE  = 12.00

FESTIVOS = {
    date(2025,1,1),date(2025,1,6),date(2025,3,19),date(2025,4,17),date(2025,4,18),
    date(2025,4,28),date(2025,5,1),date(2025,8,15),date(2025,10,9),date(2025,10,12),
    date(2025,11,1),date(2025,12,6),date(2025,12,8),date(2025,12,25),
    date(2026,1,1),date(2026,1,6),date(2026,3,19),date(2026,4,2),date(2026,4,3),
    date(2026,4,20),date(2026,5,1),date(2026,8,15),date(2026,10,9),date(2026,10,12),
    date(2026,11,1),date(2026,12,6),date(2026,12,8),date(2026,12,25),
}

COLUMNAS_BD = ["conductor","fecha","dieta","servicios_fijos","servicios_horas","observaciones"]

CSS = """
<style>
[data-testid="stSidebar"] { display: none; }
.block-container {
    padding-top: 0.8rem !important;
    padding-bottom: 0.8rem !important;
    max-width: 500px !important;
}
.stButton > button {
    height: 2.8rem; font-size: 0.95rem;
    font-weight: 700; border-radius: 10px; width: 100%;
}
.stSelectbox > label, .stDateInput > label,
.stTextInput > label, .stTimeInput > label,
.stCheckbox > label span, .stTextArea > label {
    font-size: 0.92rem !important; font-weight: 600 !important;
}
/* Cuadrícula compacta */
.grid-header {
    font-size: 0.72rem !important;
    font-weight: 700 !important;
    text-align: center !important;
    color: #1F4E79 !important;
    line-height: 1.1 !important;
}
.grid-label {
    font-size: 0.82rem !important;
    font-weight: 600 !important;
    color: #444 !important;
    padding-top: 0.5rem !important;
}
/* Ocultar etiquetas de checkboxes en la cuadrícula */
.grid-check label { display: none !important; }
/* Reducir espacio entre widgets */
div[data-testid="stVerticalBlock"] > div { margin-bottom: -0.6rem; }
.admin-link {
    text-align: center; color: #cccccc !important;
    font-size: 0.7rem; margin-top: 1.5rem; opacity: 0.4;
}
</style>
"""

CSS_ADMIN = """
<style>
[data-testid="stSidebar"] { display: block; }
.block-container { max-width: 1200px !important; padding-top: 1rem !important; }
</style>
"""

# ============================================================
# SUPABASE
# ============================================================
@st.cache_resource
def get_supabase():
    return create_client(st.secrets["supabase"]["url"], st.secrets["supabase"]["key"])

@st.cache_data(ttl=60)
def cargar_conductores():
    try:
        resp = get_supabase().table("conductores").select("nombre").order("nombre").execute()
        return [f["nombre"] for f in resp.data if f.get("nombre")]
    except Exception as e:
        st.error(f"❌ Error al cargar conductores: {e}"); return []

@st.cache_data(ttl=20)
def cargar_datos_brutos():
    try:
        resp = get_supabase().table("datos_brutos").select(",".join(COLUMNAS_BD)).order("id").execute()
        if not resp.data: return pd.DataFrame(columns=COLUMNAS_BD)
        df = pd.DataFrame(resp.data)
        for col in COLUMNAS_BD:
            if col not in df.columns: df[col] = None
        return df[COLUMNAS_BD]
    except Exception as e:
        st.error(f"❌ Error al cargar datos: {e}"); return pd.DataFrame(columns=COLUMNAS_BD)

def guardar_fila(fila_dict):
    try:
        get_supabase().table("datos_brutos").insert(fila_dict).execute()
        cargar_datos_brutos.clear(); return True
    except Exception as e:
        st.error(f"❌ Error al guardar: {e}"); return False

def guardar_tabla_completa(df):
    try:
        sb = get_supabase()
        ids_resp = sb.table("datos_brutos").select("id").execute()
        if ids_resp.data:
            ids = [f["id"] for f in ids_resp.data]
            for i in range(0, len(ids), 500):
                sb.table("datos_brutos").delete().in_("id", ids[i:i+500]).execute()
        if not df.empty:
            filas = []
            for _, row in df.iterrows():
                fila = {col: (str(row[col]) if row[col] is not None else None)
                        for col in COLUMNAS_BD if col in df.columns}
                if "dieta" in fila:
                    fila["dieta"] = str(fila["dieta"]).lower() in ["true","1","sí","si"]
                filas.append(fila)
            for i in range(0, len(filas), 500):
                sb.table("datos_brutos").insert(filas[i:i+500]).execute()
        cargar_datos_brutos.clear(); return True
    except Exception as e:
        st.error(f"❌ Error: {e}"); return False

# ============================================================
# CÁLCULO
# ============================================================
def es_dia_especial(fecha):
    if isinstance(fecha, datetime): fecha = fecha.date()
    elif isinstance(fecha, str):
        try: fecha = pd.to_datetime(fecha).date()
        except: return False
    return fecha.weekday() >= 5 or fecha in FESTIVOS

def calcular_horas(h1, h2):
    try:
        if isinstance(h1, str): h1 = datetime.strptime(h1.strip(), "%H:%M").time()
        if isinstance(h2, str): h2 = datetime.strptime(h2.strip(), "%H:%M").time()
        i = datetime.combine(date.today(), h1)
        f = datetime.combine(date.today(), h2)
        return round((f-i).total_seconds()/3600, 4) if f > i else 0.0
    except: return 0.0

def calcular_liquidacion(df):
    if df.empty: return pd.DataFrame()
    df = df.copy()
    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce").dt.date
    def a_bool(v):
        if isinstance(v, bool): return v
        return str(v).strip().lower() in ["true","1","sí","si"]
    df["dieta"] = df["dieta"].fillna(False).apply(a_bool)
    resultados = []
    for _, fila in df.iterrows():
        conductor = fila.get("conductor","")
        fecha = fila.get("fecha")
        if pd.isna(fecha): continue
        tarifa = PRECIO_HORA_FINDE if es_dia_especial(fecha) else PRECIO_HORA_SEMANA
        imp_dieta = PRECIO_DIETA if fila.get("dieta") else 0.0
        conceptos = []; horas_tot = 0.0; imp_horas = 0.0; imp_fijos = 0.0
        sf_raw = fila.get("servicios_fijos")
        if sf_raw and str(sf_raw) not in ("","nan","None","[]"):
            try:
                for item in json.loads(str(sf_raw)):
                    tipo = item.get("tipo",""); cant = int(item.get("cantidad",1))
                    if tipo in PRECIOS_FIJOS:
                        imp_fijos += PRECIOS_FIJOS[tipo] * cant
                        conceptos.append(f"{tipo}" if cant==1 else f"{tipo} x{cant}")
            except: pass
        sh_raw = fila.get("servicios_horas")
        if sh_raw and str(sh_raw) not in ("","nan","None","[]"):
            try:
                for item in json.loads(str(sh_raw)):
                    concepto = item.get("concepto","")
                    h1 = item.get("inicio",""); h2 = item.get("fin","")
                    if h1 and h2:
                        h = calcular_horas(h1, h2); horas_tot += h; imp_horas += h * tarifa
                    if concepto: conceptos.append(concepto)
            except: pass
        obs = str(fila.get("observaciones","")) if fila.get("observaciones") else ""
        resultados.append({
            "Conductor": conductor, "Fecha": fecha,
            "Conceptos del Día": " | ".join(conceptos) if conceptos else "—",
            "Observaciones": obs,
            "Horas Extras Totales": round(horas_tot,2),
            "IMPORTE HORAS (€)": round(imp_horas,2),
            "IMPORTE FIJOS (€)": round(imp_fijos,2),
            "IMPORTE DIETAS (€)": round(imp_dieta,2),
            "TOTAL GENERAL A PAGAR (€)": round(imp_horas+imp_fijos+imp_dieta,2),
        })
    return pd.DataFrame(resultados).sort_values(["Conductor","Fecha"]).reset_index(drop=True)

# ============================================================
# EXCEL
# ============================================================
def generar_excel_bytes(df_cierre):
    output = io.BytesIO()
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Cierre Mensual"
    C_AZUL="1F4E79"; C_AMAR="FFF2CC"; C_VERDE="C6EFCE"
    borde = Border(left=Side(style="thin",color="AAAAAA"),right=Side(style="thin",color="AAAAAA"),
                   top=Side(style="thin",color="AAAAAA"),bottom=Side(style="thin",color="AAAAAA"))
    al_c=Alignment(horizontal="center",vertical="center",wrap_text=True)
    al_d=Alignment(horizontal="right",vertical="center")
    al_i=Alignment(horizontal="left",vertical="center",wrap_text=True)
    cabeceras=["Conductor","Fecha","Conceptos del Día","Observaciones",
               "Horas Extras\nTotales","IMPORTE\nHORAS (€)","IMPORTE\nFIJOS (€)",
               "IMPORTE\nDIETAS (€)","TOTAL GENERAL\nA PAGAR (€)"]
    for cn,t in enumerate(cabeceras,1):
        c=ws.cell(row=1,column=cn,value=t)
        c.font=Font(bold=True,color="FFFFFF",size=11,name="Calibri")
        c.fill=PatternFill("solid",fgColor=C_AZUL); c.alignment=al_c; c.border=borde
    ws.row_dimensions[1].height=42
    for idx,fila in df_cierre.iterrows():
        fxls=idx+2; fecha=fila["Fecha"]
        fondo=es_dia_especial(fecha)
        fill=PatternFill("solid",fgColor=C_AMAR) if fondo else None
        vals=[fila["Conductor"],fecha,fila["Conceptos del Día"],fila["Observaciones"],
              fila["Horas Extras Totales"],fila["IMPORTE HORAS (€)"],fila["IMPORTE FIJOS (€)"],
              fila["IMPORTE DIETAS (€)"],fila["TOTAL GENERAL A PAGAR (€)"]]
        for cn,valor in enumerate(vals,1):
            c=ws.cell(row=fxls,column=cn,value=valor); c.border=borde
            if fill: c.fill=fill
            if cn==1: c.font=Font(bold=True,size=10,name="Calibri"); c.alignment=al_i
            elif cn==2: c.number_format="DD/MM/YYYY"; c.alignment=al_c; c.font=Font(size=10,name="Calibri")
            elif cn in(3,4): c.alignment=al_i; c.font=Font(size=10,name="Calibri")
            elif cn==5: c.number_format="0.00"; c.alignment=al_c; c.font=Font(size=10,name="Calibri")
            elif cn in(6,7,8): c.number_format='#,##0.00 "€"'; c.alignment=al_d; c.font=Font(size=10,name="Calibri")
            elif cn==9:
                c.number_format='#,##0.00 "€"'; c.alignment=al_d
                c.font=Font(bold=True,size=10,name="Calibri")
                if not fondo: c.fill=PatternFill("solid",fgColor=C_VERDE)
        ws.row_dimensions[fxls].height=20
    ft=len(df_cierre)+2
    ct=ws.cell(row=ft,column=1,value="✔  TOTAL GLOBAL")
    ct.font=Font(bold=True,size=12,color="FFFFFF",name="Calibri")
    ct.fill=PatternFill("solid",fgColor=C_AZUL); ct.alignment=al_c; ct.border=borde
    for cn in(2,3,4):
        c=ws.cell(row=ft,column=cn); c.fill=PatternFill("solid",fgColor=C_AZUL); c.border=borde
    l5=get_column_letter(5)
    c5=ws.cell(row=ft,column=5,value=f"=SUM({l5}2:{l5}{ft-1})")
    c5.number_format="0.00"; c5.font=Font(bold=True,size=11,name="Calibri")
    c5.fill=PatternFill("solid",fgColor=C_VERDE); c5.alignment=al_c; c5.border=borde
    for cn in(6,7,8,9):
        letra=get_column_letter(cn)
        c=ws.cell(row=ft,column=cn,value=f"=SUM({letra}2:{letra}{ft-1})")
        c.number_format='#,##0.00 "€"'; c.font=Font(bold=True,size=11,name="Calibri")
        c.fill=PatternFill("solid",fgColor=C_VERDE); c.alignment=al_d; c.border=borde
    ws.row_dimensions[ft].height=28
    anchos={1:28,2:13,3:45,4:30,5:12,6:16,7:16,8:16,9:22}
    for cn,ancho in anchos.items():
        ws.column_dimensions[get_column_letter(cn)].width=ancho
    ws.freeze_panes="A2"; ws.auto_filter.ref=f"A1:I{ft-1}"
    wb.save(output); output.seek(0); return output.getvalue()

# ============================================================
# VISTA CONDUCTOR
# ============================================================
def vista_conductor():
    st.markdown(CSS, unsafe_allow_html=True)

    # ── PANTALLA ÉXITO ───────────────────────────────────────
    if st.session_state.get("envio_ok"):
        info = st.session_state.get("envio_info", {})
        st.markdown("## 🚌 Autocares Alegre")
        st.divider()
        st.success(
            f"✅ **¡Registro enviado!**\n\n"
            f"👤 **{info.get('nombre','')}**  —  📅 {info.get('fecha','')}"
        )
        if st.button("➕ Registrar otro servicio", type="primary", use_container_width=True):
            for key in ["envio_ok","resumen","datos_resumen","n_horas"]:
                st.session_state.pop(key, None)
            st.rerun()
        st.markdown('<p class="admin-link">· · ·</p>', unsafe_allow_html=True)
        return

    # ── PANTALLA RESUMEN ─────────────────────────────────────
    if st.session_state.get("resumen"):
        datos = st.session_state.get("datos_resumen", {})
        st.markdown("## 🚌 Autocares Alegre")
        st.markdown("### 📋 Revisa tu registro antes de enviar")
        st.divider()
        st.markdown(f"👤 **Conductor:** {datos.get('nombre','')}")
        st.markdown(f"📅 **Fecha:** {datos.get('fecha_str','')}")
        st.markdown(f"🍽️ **Dieta:** {'✅ Sí' if datos.get('dieta') else '❌ No'}")
        fijos_ok = datos.get("fijos_ok", [])
        if fijos_ok:
            st.markdown("**📌 Servicios Fijos:**")
            for f in fijos_ok:
                c = f['cantidad']
                st.markdown(f"- {f['tipo']}" + (f" × {c}" if c > 1 else ""))
        horas_ok = datos.get("horas_ok", [])
        if horas_ok:
            st.markdown("**🕐 Servicios Por Horas:**")
            for h in horas_ok:
                st.markdown(f"- {h['concepto']}  ({h['inicio']} → {h['fin']})")
        obs = datos.get("observaciones","")
        if obs:
            st.markdown(f"**💬 Observaciones:** {obs}")
        st.divider()
        c1, c2 = st.columns(2)
        with c1:
            if st.button("✏️ Corregir", type="secondary", use_container_width=True):
                st.session_state.resumen = False
                st.rerun()
        with c2:
            if st.button("✅ Confirmar y Enviar", type="primary", use_container_width=True):
                fila = {
                    "conductor":       datos["nombre"],
                    "fecha":           datos["fecha_iso"],
                    "dieta":           datos["dieta"],
                    "servicios_fijos": json.dumps(datos["fijos_ok"], ensure_ascii=False),
                    "servicios_horas": json.dumps(datos["horas_ok"], ensure_ascii=False),
                    "observaciones":   datos.get("observaciones") or None,
                }
                with st.spinner("Enviando..."):
                    ok = guardar_fila(fila)
                if ok:
                    st.session_state.envio_ok   = True
                    st.session_state.envio_info = {"nombre": datos["nombre"], "fecha": datos["fecha_str"]}
                    st.session_state.resumen    = False
                    for key in ["n_horas"]: st.session_state.pop(key, None)
                    st.rerun()
        return

    # ── FORMULARIO PRINCIPAL ─────────────────────────────────
    st.markdown("## 🚌 Autocares Alegre")
    st.markdown("##### Registro de Servicios Extras")
    st.divider()

    # Buscador + selector de nombre
    conductores = cargar_conductores()
    busqueda = st.text_input("🔍 Buscar conductor", placeholder="Escribe las primeras letras...")
    conductores_filtrados = (
        [c for c in conductores if busqueda.strip().lower() in c.lower()]
        if busqueda.strip() else conductores
    )
    nombre_sel = st.selectbox(
        "👤 Selecciona tu Nombre",
        options=["— Selecciona —"] + conductores_filtrados
    )

    # Fecha
    fecha_sel = st.date_input("📅 Fecha del Servicio", value=date.today(), format="DD/MM/YYYY")

    st.markdown("---")

    # ════════════════════════════════════════════════════════
    # CUADRÍCULA DE SERVICIOS FIJOS
    # ════════════════════════════════════════════════════════
    st.markdown("**📌 Servicios Fijos**")
    st.caption("Marca los servicios realizados. Cada línea = un servicio.")

    # Cabecera de la cuadrícula
    # Columna 0 = etiqueta fila, columnas 1-8 = servicios
    pesos = [1.2, 0.9, 0.9, 1.0, 0.8, 0.9, 0.9, 0.9, 0.9]
    cab = st.columns(pesos)
    cab[0].markdown("")
    for i, s in enumerate(SERVICIOS_GRID):
        cab[i+1].markdown(
            f"<div class='grid-header'>{ABREV[s]}</div>",
            unsafe_allow_html=True
        )

    # 5 filas de checkboxes
    grid_checks = {}
    for fila in range(N_FILAS_GRID):
        cols = st.columns(pesos)
        cols[0].markdown(
            f"<div class='grid-label'>Serv.{fila+1}</div>",
            unsafe_allow_html=True
        )
        for i, s in enumerate(SERVICIOS_GRID):
            key = f"grid_{fila}_{i}"
            grid_checks[key] = cols[i+1].checkbox(
                "", key=key, label_visibility="collapsed"
            )

    # Leyenda abreviaturas
    leyenda = " · ".join([f"{v}={k}" for k,v in ABREV.items()])
    st.caption(f"📖 {leyenda}")

    st.markdown("---")

    # ════════════════════════════════════════════════════════
    # SERVICIOS POR HORAS
    # ════════════════════════════════════════════════════════
    st.markdown("**🕐 Servicios Por Horas**")
    st.caption("Solo si el servicio se paga por horas.")

    if "n_horas" not in st.session_state:
        st.session_state.n_horas = 1

    horas_capturadas = []
    for i in range(st.session_state.n_horas):
        if i > 0: st.markdown("---")
        concepto_inp = st.text_input(
            f"Concepto / Destino #{i+1}",
            placeholder="Excursión o salida / Viaje a...",
            key=f"h_concepto_{i}"
        )
        c1, c2 = st.columns(2)
        with c1:
            h_ini = st.time_input("🕐 Inicio", value=time(8,0), step=300, key=f"h_ini_{i}")
        with c2:
            h_fin = st.time_input("🕔 Fin",    value=time(16,0), step=300, key=f"h_fin_{i}")
        horas_capturadas.append({"concepto": concepto_inp, "hora_ini": h_ini, "hora_fin": h_fin})

    c3, c4 = st.columns(2)
    with c3:
        if st.button("➕ Añadir horas", type="secondary", use_container_width=True):
            st.session_state.n_horas += 1; st.rerun()
    with c4:
        if st.session_state.n_horas > 1:
            if st.button("➖ Quitar", type="secondary", use_container_width=True):
                st.session_state.n_horas -= 1; st.rerun()

    st.markdown("---")

    # ════════════════════════════════════════════════════════
    # DIETA + OBSERVACIONES
    # ════════════════════════════════════════════════════════
    dieta_sel = st.checkbox("🍽️ ¿Te corresponde Dieta hoy?", value=False)

    observaciones = st.text_area(
        "💬 Observaciones (opcional)",
        placeholder="Incidencias, dudas o comentarios del día...",
        height=65
    )

    st.divider()

    # ── BOTÓN VER RESUMEN ────────────────────────────────────
    if st.button("👁️ VER RESUMEN Y CONFIRMAR", type="primary", use_container_width=True):
        errores = []

        if nombre_sel == "— Selecciona —":
            errores.append("❌ Debes seleccionar tu nombre.")

        # Recoger servicios fijos marcados en la cuadrícula
        conteo = {}
        for fila in range(N_FILAS_GRID):
            for i, s in enumerate(SERVICIOS_GRID):
                key = f"grid_{fila}_{i}"
                if st.session_state.get(key, False):
                    conteo[s] = conteo.get(s, 0) + 1
        fijos_ok = [{"tipo": t, "cantidad": c} for t, c in conteo.items()]

        # Recoger servicios por horas
        horas_ok = []
        for i, h in enumerate(horas_capturadas):
            if h["concepto"].strip():
                if h["hora_fin"] <= h["hora_ini"]:
                    errores.append(f"❌ Horas #{i+1}: la hora fin debe ser posterior al inicio.")
                else:
                    horas_ok.append({
                        "concepto": h["concepto"].strip(),
                        "inicio":   h["hora_ini"].strftime("%H:%M"),
                        "fin":      h["hora_fin"].strftime("%H:%M"),
                    })

        if not fijos_ok and not horas_ok and not errores:
            errores.append("❌ Marca al menos un servicio fijo o añade un servicio por horas.")

        if errores:
            for msg in errores: st.error(msg)
        else:
            st.session_state.resumen = True
            st.session_state.datos_resumen = {
                "nombre":        nombre_sel,
                "fecha_str":     fecha_sel.strftime("%d/%m/%Y"),
                "fecha_iso":     fecha_sel.strftime("%Y-%m-%d"),
                "dieta":         dieta_sel,
                "fijos_ok":      fijos_ok,
                "horas_ok":      horas_ok,
                "observaciones": observaciones.strip(),
            }
            st.rerun()

    st.markdown(
        '<p class="admin-link"><a href="?admin=1" style="color:#cccccc;text-decoration:none;">· · ·</a></p>',
        unsafe_allow_html=True
    )


# ============================================================
# VISTA ADMINISTRACIÓN
# ============================================================
def vista_admin():
    st.markdown(CSS_ADMIN, unsafe_allow_html=True)

    if "admin_ok" not in st.session_state:
        st.session_state.admin_ok = False

    if not st.session_state.admin_ok:
        st.title("🔐 Administración")
        st.divider()
        pw = st.text_input("Contraseña", type="password")
        if st.button("▶ Entrar", type="primary", use_container_width=True):
            if pw == PASSWORD_ADMIN:
                st.session_state.admin_ok = True; st.rerun()
            else:
                st.error("❌ Contraseña incorrecta.")
        return

    c1, c2 = st.columns([5,1])
    with c1: st.title("📊 Administración — Autocares Alegre")
    with c2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🚪 Salir"):
            st.session_state.admin_ok = False; st.rerun()
    st.divider()

    # ── Registros en tiempo real ─────────────────────────────
    st.markdown("## 👁️ Registros en Tiempo Real")
    if st.button("🔄 Actualizar", type="secondary"):
        cargar_datos_brutos.clear(); st.rerun()

    df_brutos = cargar_datos_brutos()

    if df_brutos.empty:
        st.info("ℹ️ Todavía no hay registros.")
    else:
        filas_vista = []
        for _, row in df_brutos.iterrows():
            sf_texto = ""
            try:
                sf = json.loads(str(row.get("servicios_fijos") or "[]"))
                sf_texto = ", ".join(
                    f"{i['tipo']} x{i['cantidad']}" if i['cantidad']>1 else i['tipo']
                    for i in sf if i.get("tipo")
                )
            except: pass
            sh_texto = ""
            try:
                sh = json.loads(str(row.get("servicios_horas") or "[]"))
                sh_texto = ", ".join(
                    f"{i['concepto']} ({i['inicio']}-{i['fin']})"
                    for i in sh if i.get("concepto")
                )
            except: pass
            filas_vista.append({
                "Conductor":       row.get("conductor",""),
                "Fecha":           row.get("fecha",""),
                "Dieta":           "✓" if str(row.get("dieta","")).lower() in ["true","1"] else "",
                "Servicios Fijos": sf_texto,
                "Por Horas":       sh_texto,
                "Observaciones":   str(row.get("observaciones","")) if row.get("observaciones") else "",
            })
        st.dataframe(pd.DataFrame(filas_vista), use_container_width=True, hide_index=True)

    st.divider()

    # ── Auditoría ────────────────────────────────────────────
    st.markdown("## 📋 Tabla de Auditoría")
    if not df_brutos.empty:
        df_show = df_brutos.copy()
        df_show["dieta"] = df_show["dieta"].fillna(False).apply(
            lambda x: bool(x) if isinstance(x, bool) else str(x).lower() in ["true","1"]
        )
        for col in ["conductor","fecha","servicios_fijos","servicios_horas","observaciones"]:
            df_show[col] = df_show[col].fillna("").astype(str)

        df_editado = st.data_editor(
            df_show, use_container_width=True, num_rows="dynamic", hide_index=True,
            column_config={
                "conductor":       st.column_config.TextColumn("Conductor",    width="medium"),
                "fecha":           st.column_config.TextColumn("Fecha"),
                "dieta":           st.column_config.CheckboxColumn("🍽️ Dieta", default=False),
                "servicios_fijos": st.column_config.TextColumn("Fijos (JSON)", width="large"),
                "servicios_horas": st.column_config.TextColumn("Horas (JSON)", width="large"),
                "observaciones":   st.column_config.TextColumn("Observaciones",width="medium"),
            },
            key="editor_brutos",
        )
        if st.button("💾 Guardar Cambios", type="primary"):
            with st.spinner("Guardando..."):
                ok = guardar_tabla_completa(df_editado)
            if ok: st.success("✅ Cambios guardados.")

    st.divider()

    # ── Liquidación ──────────────────────────────────────────
    st.markdown("## 💰 Liquidación Mensual")
    if st.button("📊 CALCULAR Y GENERAR EXCEL", type="primary", use_container_width=True):
        df_calc = cargar_datos_brutos()
        if df_calc.empty:
            st.error("❌ No hay datos para calcular.")
        else:
            with st.spinner("⏳ Calculando..."):
                df_cierre = calcular_liquidacion(df_calc)
            if df_cierre.empty:
                st.error("❌ No se pudieron calcular los importes.")
            else:
                nc=df_cierre["Conductor"].nunique(); nd=len(df_cierre)
                st.success(f"✅ **{nd} días** de **{nc} conductor(es)**.")
                m1,m2,m3,m4=st.columns(4)
                m1.metric("💼 Horas",  f"{df_cierre['IMPORTE HORAS (€)'].sum():,.2f} €")
                m2.metric("📌 Fijos",  f"{df_cierre['IMPORTE FIJOS (€)'].sum():,.2f} €")
                m3.metric("🍽️ Dietas", f"{df_cierre['IMPORTE DIETAS (€)'].sum():,.2f} €")
                m4.metric("💰 TOTAL",   f"{df_cierre['TOTAL GENERAL A PAGAR (€)'].sum():,.2f} €")
                st.dataframe(df_cierre, use_container_width=True, hide_index=True)
                excel_bytes=generar_excel_bytes(df_cierre)
                nombre_arch=f"cierre_mensual_{date.today().strftime('%Y_%m_%d')}.xlsx"
                st.download_button(
                    label="⬇️ DESCARGAR EXCEL CIERRE MENSUAL",
                    data=excel_bytes, file_name=nombre_arch,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary", use_container_width=True,
                )


# ============================================================
# NAVEGACIÓN
# ============================================================
def main():
    if "vista_actual" not in st.session_state:
        st.session_state.vista_actual = "conductor"

    params = st.query_params
    if params.get("admin") == "1" and st.session_state.vista_actual == "conductor":
        st.session_state.vista_actual = "admin"

    if st.session_state.vista_actual == "admin":
        with st.sidebar:
            st.markdown("## 🚌 Autocares Alegre")
            st.markdown("---")
            if st.button("👨‍✈️ Vista Conductor", use_container_width=True):
                st.session_state.vista_actual = "conductor"
                st.query_params.clear()
                st.rerun()
            st.markdown("---")
            st.caption("v6.0 · Autocares Alegre")
        vista_admin()
    else:
        vista_conductor()

if __name__ == "__main__":
    main()
