# ============================================================
#  AUTOCARES ALEGRE — Sistema de Gestión de Servicios Extras
#  Versión DEFINITIVA 4.0
#  - 1 fila por conductor por día
#  - Servicios fijos con cantidad
#  - Múltiples servicios Por Horas
#  - Campo de Observaciones
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

# ============================================================
# 1. CONFIGURACIÓN GENERAL
# ============================================================
st.set_page_config(
    page_title="Autocares Alegre",
    page_icon="🚌",
    layout="centered",
    initial_sidebar_state="auto"
)

# ============================================================
# 2. CONSTANTES
# ============================================================
PASSWORD_ADMIN = "alegre2026"

PRECIOS_FIJOS = {
    "Pilotos":         15.00,
    "Pescadores":      15.00,
    "Logista":         15.00,
    "Saludes":         15.00,
    "Boda Ida":        40.00,
    "Boda Regreso 1":  40.00,
    "Boda Regreso 2":  40.00,
    "Transfer":        20.00,
}
LISTA_FIJOS = list(PRECIOS_FIJOS.keys())

PRECIO_DIETA       = 15.00
PRECIO_HORA_SEMANA = 10.00
PRECIO_HORA_FINDE  = 12.00

FESTIVOS = {
    date(2025,  1,  1), date(2025,  1,  6), date(2025,  3, 19),
    date(2025,  4, 17), date(2025,  4, 18), date(2025,  4, 28),
    date(2025,  5,  1), date(2025,  8, 15), date(2025, 10,  9),
    date(2025, 10, 12), date(2025, 11,  1), date(2025, 12,  6),
    date(2025, 12,  8), date(2025, 12, 25),
    date(2026,  1,  1), date(2026,  1,  6), date(2026,  3, 19),
    date(2026,  4,  2), date(2026,  4,  3), date(2026,  4, 20),
    date(2026,  5,  1), date(2026,  8, 15), date(2026, 10,  9),
    date(2026, 10, 12), date(2026, 11,  1), date(2026, 12,  6),
    date(2026, 12,  8), date(2026, 12, 25),
}

# Columnas de la tabla datos_brutos en Supabase
COLUMNAS_BD = [
    "conductor", "fecha", "dieta",
    "servicios_fijos",   # JSON: [{"tipo":"Pilotos","cantidad":2}, ...]
    "servicios_horas",   # JSON: [{"concepto":"...","inicio":"08:00","fin":"14:00"}, ...]
    "observaciones"
]

CSS = """
<style>
    .stButton > button {
        height: 3.2rem; font-size: 1rem;
        font-weight: 700; border-radius: 12px;
    }
    .stSelectbox > label, .stDateInput > label,
    .stTextInput > label, .stTimeInput > label,
    .stNumberInput > label, .stCheckbox > label span,
    .stTextArea > label {
        font-size: 1.02rem !important; font-weight: 600 !important;
    }
    .servicio-box {
        background: linear-gradient(135deg, #f0f4ff 0%, #e8eeff 100%);
        border-left: 5px solid #1F4E79;
        padding: 1.2rem 1rem 0.8rem 1rem;
        border-radius: 10px; margin-bottom: 1rem;
    }
    .horas-box {
        background: linear-gradient(135deg, #fff8f0 0%, #ffeedc 100%);
        border-left: 5px solid #E67E22;
        padding: 1.2rem 1rem 0.8rem 1rem;
        border-radius: 10px; margin-bottom: 1rem;
    }
</style>
"""

# ============================================================
# 3. CONEXIÓN A SUPABASE
# ============================================================

@st.cache_resource
def get_supabase():
    url = st.secrets["supabase"]["url"]
    key = st.secrets["supabase"]["key"]
    return create_client(url, key)


# ============================================================
# 4. LECTURA Y ESCRITURA EN SUPABASE
# ============================================================

@st.cache_data(ttl=60)
def cargar_conductores():
    try:
        sb = get_supabase()
        resp = sb.table("conductores").select("nombre").order("nombre").execute()
        return [f["nombre"] for f in resp.data if f.get("nombre")]
    except Exception as e:
        st.error(f"❌ Error al cargar conductores: {e}")
        return []


@st.cache_data(ttl=20)
def cargar_datos_brutos():
    try:
        sb = get_supabase()
        resp = sb.table("datos_brutos") \
                 .select(",".join(COLUMNAS_BD)) \
                 .order("id").execute()
        if not resp.data:
            return pd.DataFrame(columns=COLUMNAS_BD)
        df = pd.DataFrame(resp.data)
        for col in COLUMNAS_BD:
            if col not in df.columns:
                df[col] = None
        return df[COLUMNAS_BD]
    except Exception as e:
        st.error(f"❌ Error al cargar datos: {e}")
        return pd.DataFrame(columns=COLUMNAS_BD)


def guardar_fila(fila_dict):
    try:
        sb = get_supabase()
        sb.table("datos_brutos").insert(fila_dict).execute()
        cargar_datos_brutos.clear()
        return True
    except Exception as e:
        st.error(f"❌ Error al guardar: {e}")
        return False


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
                    v = fila["dieta"]
                    fila["dieta"] = str(v).lower() in ["true","1","sí","si"]
                filas.append(fila)
            for i in range(0, len(filas), 500):
                sb.table("datos_brutos").insert(filas[i:i+500]).execute()
        cargar_datos_brutos.clear()
        return True
    except Exception as e:
        st.error(f"❌ Error al guardar cambios: {e}")
        return False


# ============================================================
# 5. FUNCIONES DE CÁLCULO
# ============================================================

def es_dia_especial(fecha):
    if isinstance(fecha, datetime): fecha = fecha.date()
    elif isinstance(fecha, str):
        try: fecha = pd.to_datetime(fecha).date()
        except: return False
    return fecha.weekday() >= 5 or fecha in FESTIVOS


def calcular_horas(h_ini, h_fin):
    try:
        if isinstance(h_ini, str): h_ini = datetime.strptime(h_ini.strip(), "%H:%M").time()
        if isinstance(h_fin, str): h_fin = datetime.strptime(h_fin.strip(), "%H:%M").time()
        inicio = datetime.combine(date.today(), h_ini)
        fin    = datetime.combine(date.today(), h_fin)
        return round((fin - inicio).total_seconds() / 3600, 4) if fin > inicio else 0.0
    except:
        return 0.0


def calcular_liquidacion(df):
    if df.empty:
        return pd.DataFrame()

    df = df.copy()
    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce").dt.date

    def a_bool(v):
        if isinstance(v, bool): return v
        return str(v).strip().lower() in ["true","1","sí","si"]
    df["dieta"] = df["dieta"].fillna(False).apply(a_bool)

    resultados = []

    for _, fila in df.iterrows():
        conductor = fila.get("conductor", "")
        fecha     = fila.get("fecha")
        if pd.isna(fecha): continue

        es_esp    = es_dia_especial(fecha)
        tarifa    = PRECIO_HORA_FINDE if es_esp else PRECIO_HORA_SEMANA
        imp_dieta = PRECIO_DIETA if fila.get("dieta") else 0.0
        conceptos = []
        horas_tot = 0.0
        imp_horas = 0.0
        imp_fijos = 0.0

        # ── Servicios fijos ──────────────────────────────
        sf_raw = fila.get("servicios_fijos")
        if sf_raw and str(sf_raw) not in ("", "nan", "None", "[]"):
            try:
                sf_lista = json.loads(str(sf_raw))
                for item in sf_lista:
                    tipo     = item.get("tipo", "")
                    cantidad = int(item.get("cantidad", 1))
                    if tipo in PRECIOS_FIJOS:
                        imp_fijos += PRECIOS_FIJOS[tipo] * cantidad
                        label = f"{tipo}" if cantidad == 1 else f"{tipo} x{cantidad}"
                        conceptos.append(label)
            except:
                pass

        # ── Servicios por horas ──────────────────────────
        sh_raw = fila.get("servicios_horas")
        if sh_raw and str(sh_raw) not in ("", "nan", "None", "[]"):
            try:
                sh_lista = json.loads(str(sh_raw))
                for item in sh_lista:
                    concepto = item.get("concepto", "")
                    h_ini    = item.get("inicio", "")
                    h_fin    = item.get("fin", "")
                    if h_ini and h_fin:
                        horas      = calcular_horas(h_ini, h_fin)
                        horas_tot += horas
                        imp_horas += horas * tarifa
                    if concepto:
                        conceptos.append(concepto)
            except:
                pass

        obs = str(fila.get("observaciones","")) if fila.get("observaciones") else ""

        resultados.append({
            "Conductor":                 conductor,
            "Fecha":                     fecha,
            "Conceptos del Día":         " | ".join(conceptos) if conceptos else "—",
            "Observaciones":             obs,
            "Horas Extras Totales":      round(horas_tot, 2),
            "IMPORTE HORAS (€)":         round(imp_horas, 2),
            "IMPORTE FIJOS (€)":         round(imp_fijos, 2),
            "IMPORTE DIETAS (€)":        round(imp_dieta, 2),
            "TOTAL GENERAL A PAGAR (€)": round(imp_horas + imp_fijos + imp_dieta, 2),
        })

    df_res = pd.DataFrame(resultados)
    return df_res.sort_values(["Conductor","Fecha"]).reset_index(drop=True)


# ============================================================
# 6. GENERACIÓN EXCEL
# ============================================================

def generar_excel_bytes(df_cierre):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cierre Mensual"

    C_AZUL  = "1F4E79"
    C_AMAR  = "FFF2CC"
    C_VERDE = "C6EFCE"

    borde = Border(
        left=Side(style="thin", color="AAAAAA"),
        right=Side(style="thin", color="AAAAAA"),
        top=Side(style="thin", color="AAAAAA"),
        bottom=Side(style="thin", color="AAAAAA"),
    )
    al_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al_d = Alignment(horizontal="right",  vertical="center")
    al_i = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    cabeceras = [
        "Conductor", "Fecha", "Conceptos del Día", "Observaciones",
        "Horas Extras\nTotales", "IMPORTE\nHORAS (€)",
        "IMPORTE\nFIJOS (€)", "IMPORTE\nDIETAS (€)",
        "TOTAL GENERAL\nA PAGAR (€)",
    ]
    for cn, t in enumerate(cabeceras, 1):
        c = ws.cell(row=1, column=cn, value=t)
        c.font      = Font(bold=True, color="FFFFFF", size=11, name="Calibri")
        c.fill      = PatternFill("solid", fgColor=C_AZUL)
        c.alignment = al_c
        c.border    = borde
    ws.row_dimensions[1].height = 42

    for idx, fila in df_cierre.iterrows():
        fxls  = idx + 2
        fecha = fila["Fecha"]
        fondo = es_dia_especial(fecha)
        fill  = PatternFill("solid", fgColor=C_AMAR) if fondo else None

        vals = [
            fila["Conductor"], fecha,
            fila["Conceptos del Día"], fila["Observaciones"],
            fila["Horas Extras Totales"],
            fila["IMPORTE HORAS (€)"], fila["IMPORTE FIJOS (€)"],
            fila["IMPORTE DIETAS (€)"], fila["TOTAL GENERAL A PAGAR (€)"],
        ]
        for cn, valor in enumerate(vals, 1):
            c = ws.cell(row=fxls, column=cn, value=valor)
            c.border = borde
            if fill: c.fill = fill
            if cn == 1:
                c.font = Font(bold=True, size=10, name="Calibri"); c.alignment = al_i
            elif cn == 2:
                c.number_format = "DD/MM/YYYY"; c.alignment = al_c
                c.font = Font(size=10, name="Calibri")
            elif cn in (3, 4):
                c.alignment = al_i; c.font = Font(size=10, name="Calibri")
            elif cn == 5:
                c.number_format = "0.00"; c.alignment = al_c
                c.font = Font(size=10, name="Calibri")
            elif cn in (6, 7, 8):
                c.number_format = '#,##0.00 "€"'; c.alignment = al_d
                c.font = Font(size=10, name="Calibri")
            elif cn == 9:
                c.number_format = '#,##0.00 "€"'; c.alignment = al_d
                c.font = Font(bold=True, size=10, name="Calibri")
                if not fondo: c.fill = PatternFill("solid", fgColor=C_VERDE)
        ws.row_dimensions[fxls].height = 20

    ft = len(df_cierre) + 2
    ct = ws.cell(row=ft, column=1, value="✔  TOTAL GLOBAL")
    ct.font = Font(bold=True, size=12, color="FFFFFF", name="Calibri")
    ct.fill = PatternFill("solid", fgColor=C_AZUL)
    ct.alignment = al_c; ct.border = borde
    for cn in (2, 3, 4):
        c = ws.cell(row=ft, column=cn)
        c.fill = PatternFill("solid", fgColor=C_AZUL); c.border = borde
    l5 = get_column_letter(5)
    c5 = ws.cell(row=ft, column=5, value=f"=SUM({l5}2:{l5}{ft-1})")
    c5.number_format = "0.00"; c5.font = Font(bold=True, size=11, name="Calibri")
    c5.fill = PatternFill("solid", fgColor=C_VERDE)
    c5.alignment = al_c; c5.border = borde
    for cn in (6, 7, 8, 9):
        letra = get_column_letter(cn)
        c = ws.cell(row=ft, column=cn, value=f"=SUM({letra}2:{letra}{ft-1})")
        c.number_format = '#,##0.00 "€"'
        c.font = Font(bold=True, size=11, name="Calibri")
        c.fill = PatternFill("solid", fgColor=C_VERDE)
        c.alignment = al_d; c.border = borde
    ws.row_dimensions[ft].height = 28

    anchos = {1:28, 2:13, 3:45, 4:30, 5:12, 6:16, 7:16, 8:16, 9:22}
    for cn, ancho in anchos.items():
        ws.column_dimensions[get_column_letter(cn)].width = ancho
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:I{ft-1}"

    wb.save(output)
    output.seek(0)
    return output.getvalue()


# ============================================================
# 7. VISTA DEL CONDUCTOR
# ============================================================

def vista_conductor():
    st.markdown(CSS, unsafe_allow_html=True)
    st.title("🚌 Autocares Alegre")
    st.markdown("#### Registro de Servicios Extras")
    st.divider()

    # Pantalla de confirmación
    if st.session_state.get("envio_ok"):
        info = st.session_state.get("envio_info", {})
        st.success(
            f"✅ **¡Registro enviado correctamente!**\n\n"
            f"👤 **{info.get('nombre','')}** — 📅 {info.get('fecha','')}"
        )
        st.balloons()
        if st.button("➕  Hacer Otro Registro", type="primary", use_container_width=True):
            st.session_state.envio_ok = False
            st.session_state.n_horas  = 1
            st.rerun()
        return

    # ── Nombre ──────────────────────────────────────────────
    conductores = cargar_conductores()
    nombre_sel  = st.selectbox("👤  Selecciona tu Nombre",
                               ["— Selecciona tu nombre —"] + conductores)

    # ── Fecha ────────────────────────────────────────────────
    fecha_sel = st.date_input("📅  Fecha del Servicio",
                              value=date.today(), format="DD/MM/YYYY")

    # ── Dieta ────────────────────────────────────────────────
    st.markdown("---")
    dieta_sel = st.checkbox("🍽️  ¿Te corresponde Dieta hoy?", value=False)
    st.markdown("---")

    # ════════════════════════════════════════════════════════
    # SECCIÓN A: SERVICIOS FIJOS (con cantidad)
    # ════════════════════════════════════════════════════════
    st.markdown("### 📌 Servicios Fijos")
    st.caption("Elige el servicio y la cantidad. Puedes añadir varios tipos distintos.")

    if "n_fijos" not in st.session_state:
        st.session_state.n_fijos = 1

    fijos_capturados = []

    for i in range(st.session_state.n_fijos):
        st.markdown('<div class="servicio-box">', unsafe_allow_html=True)
        c1, c2 = st.columns([3, 1])
        with c1:
            tipo_sel = st.selectbox(
                f"Tipo de servicio #{i+1}",
                options=["— Elige —"] + LISTA_FIJOS,
                key=f"fijo_tipo_{i}"
            )
        with c2:
            cant_sel = st.number_input(
                "Cantidad",
                min_value=1, max_value=20, value=1,
                key=f"fijo_cant_{i}"
            )
        fijos_capturados.append({"tipo": tipo_sel, "cantidad": cant_sel})
        st.markdown("</div>", unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        if st.button("➕ Añadir otro fijo", type="secondary", use_container_width=True):
            st.session_state.n_fijos += 1
            st.rerun()
    with col2:
        if st.session_state.n_fijos > 1:
            if st.button("➖ Quitar último", type="secondary", use_container_width=True):
                st.session_state.n_fijos -= 1
                st.rerun()

    st.markdown("---")

    # ════════════════════════════════════════════════════════
    # SECCIÓN B: SERVICIOS POR HORAS
    # ════════════════════════════════════════════════════════
    st.markdown("### 🕐 Servicios Por Horas")
    st.caption("Añade uno o varios servicios con horario.")

    if "n_horas" not in st.session_state:
        st.session_state.n_horas = 1

    horas_capturadas = []

    for i in range(st.session_state.n_horas):
        st.markdown('<div class="horas-box">', unsafe_allow_html=True)
        st.markdown(f"**Servicio Por Horas #{i+1}**")
        concepto_inp = st.text_input(
            "🗺️ Concepto / Destino",
            placeholder="Ej: Excursión a Gandía / Vigilancia en Base",
            key=f"h_concepto_{i}"
        )
        c1, c2 = st.columns(2)
        with c1:
            h_ini = st.time_input("🕐 Hora Inicio", value=time(8,0),  step=300, key=f"h_ini_{i}")
        with c2:
            h_fin = st.time_input("🕔 Hora Fin",    value=time(16,0), step=300, key=f"h_fin_{i}")
        horas_capturadas.append({
            "concepto": concepto_inp,
            "hora_ini": h_ini,
            "hora_fin": h_fin
        })
        st.markdown("</div>", unsafe_allow_html=True)

    col3, col4 = st.columns(2)
    with col3:
        if st.button("➕ Añadir otra hora", type="secondary", use_container_width=True):
            st.session_state.n_horas += 1
            st.rerun()
    with col4:
        if st.session_state.n_horas > 1:
            if st.button("➖ Quitar última", type="secondary", use_container_width=True):
                st.session_state.n_horas -= 1
                st.rerun()

    st.markdown("---")

    # ════════════════════════════════════════════════════════
    # SECCIÓN C: OBSERVACIONES
    # ════════════════════════════════════════════════════════
    observaciones = st.text_area(
        "💬  Observaciones (opcional)",
        placeholder="Escribe aquí cualquier incidencia o comentario del día...",
        height=80
    )

    st.divider()

    # ── Botón Enviar ─────────────────────────────────────────
    if st.button("📤  ENVIAR REGISTRO DIARIO", type="primary", use_container_width=True):
        errores = []

        if nombre_sel == "— Selecciona tu nombre —":
            errores.append("❌ Debes seleccionar tu nombre.")

        # Validar fijos
        fijos_ok = []
        for i, f in enumerate(fijos_capturados):
            if f["tipo"] != "— Elige —":
                fijos_ok.append({"tipo": f["tipo"], "cantidad": int(f["cantidad"])})

        # Validar horas
        horas_ok = []
        for i, h in enumerate(horas_capturadas):
            if h["concepto"].strip():
                if h["hora_fin"] <= h["hora_ini"]:
                    errores.append(f"❌ Servicio Por Horas #{i+1}: la hora fin debe ser posterior al inicio.")
                else:
                    horas_ok.append({
                        "concepto": h["concepto"].strip(),
                        "inicio":   h["hora_ini"].strftime("%H:%M"),
                        "fin":      h["hora_fin"].strftime("%H:%M")
                    })

        if not fijos_ok and not horas_ok and not errores:
            errores.append("❌ Añade al menos un servicio fijo o por horas.")

        if errores:
            for msg in errores:
                st.error(msg)
        else:
            fila = {
                "conductor":       nombre_sel,
                "fecha":           fecha_sel.strftime("%Y-%m-%d"),
                "dieta":           dieta_sel,
                "servicios_fijos": json.dumps(fijos_ok,  ensure_ascii=False),
                "servicios_horas": json.dumps(horas_ok,  ensure_ascii=False),
                "observaciones":   observaciones.strip() if observaciones else None,
            }
            with st.spinner("Enviando..."):
                ok = guardar_fila(fila)
            if ok:
                st.session_state.envio_ok   = True
                st.session_state.envio_info = {
                    "nombre": nombre_sel,
                    "fecha":  fecha_sel.strftime("%d/%m/%Y"),
                }
                st.session_state.n_fijos = 1
                st.session_state.n_horas = 1
                st.rerun()


# ============================================================
# 8. VISTA DE ADMINISTRACIÓN
# ============================================================

def vista_admin():
    st.markdown(CSS, unsafe_allow_html=True)

    if "admin_ok" not in st.session_state:
        st.session_state.admin_ok = False

    if not st.session_state.admin_ok:
        st.title("🔐 Panel de Administración")
        st.divider()
        pw = st.text_input("Contraseña", type="password")
        if st.button("▶  Entrar", type="primary", use_container_width=True):
            if pw == PASSWORD_ADMIN:
                st.session_state.admin_ok = True
                st.rerun()
            else:
                st.error("❌ Contraseña incorrecta.")
        return

    c1, c2 = st.columns([5,1])
    with c1:
        st.title("📊 Administración — Autocares Alegre")
    with c2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🚪 Salir"):
            st.session_state.admin_ok = False
            st.rerun()
    st.divider()

    # ════════════════════════════════════════════════════════
    # SECCIÓN A: REGISTROS EN TIEMPO REAL
    # ════════════════════════════════════════════════════════
    st.markdown("## 👁️ Registros en Tiempo Real")
    st.caption("Todos los servicios enviados por los conductores. Se actualiza solo cada 20 segundos.")

    df_brutos = cargar_datos_brutos()

    if st.button("🔄 Actualizar ahora", type="secondary"):
        cargar_datos_brutos.clear()
        st.rerun()

    if df_brutos.empty:
        st.info("ℹ️ Todavía no hay registros.")
    else:
        # Vista legible: expandir el JSON para mostrar en tabla
        filas_vista = []
        for _, row in df_brutos.iterrows():
            # Parsear servicios fijos
            sf_texto = ""
            try:
                sf = json.loads(str(row.get("servicios_fijos") or "[]"))
                sf_texto = ", ".join(
                    f"{i['tipo']} x{i['cantidad']}" if i['cantidad'] > 1 else i['tipo']
                    for i in sf if i.get("tipo") and i["tipo"] != "— Elige —"
                )
            except: pass

            # Parsear servicios horas
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

        df_vista = pd.DataFrame(filas_vista)
        st.dataframe(df_vista, use_container_width=True, hide_index=True)

    st.divider()

    # ════════════════════════════════════════════════════════
    # SECCIÓN B: TABLA DE AUDITORÍA (editable)
    # ════════════════════════════════════════════════════════
    st.markdown("## 📋 Tabla de Auditoría")
    st.caption("Aquí puedes editar o eliminar registros incorrectos.")

    if not df_brutos.empty:
        df_show = df_brutos.copy()
        df_show["dieta"] = df_show["dieta"].fillna(False).apply(
            lambda x: bool(x) if not isinstance(x, bool)
            else x if isinstance(x, bool)
            else str(x).lower() in ["true","1"]
        )
        for col in ["conductor","fecha","servicios_fijos","servicios_horas","observaciones"]:
            df_show[col] = df_show[col].fillna("").astype(str)

        df_editado = st.data_editor(
            df_show,
            use_container_width=True,
            num_rows="dynamic",
            hide_index=True,
            column_config={
                "conductor":       st.column_config.TextColumn("Conductor",     width="medium"),
                "fecha":           st.column_config.TextColumn("Fecha\n(AAAA-MM-DD)"),
                "dieta":           st.column_config.CheckboxColumn("🍽️ Dieta",  default=False),
                "servicios_fijos": st.column_config.TextColumn("Servicios Fijos (JSON)", width="large"),
                "servicios_horas": st.column_config.TextColumn("Por Horas (JSON)",       width="large"),
                "observaciones":   st.column_config.TextColumn("Observaciones",           width="medium"),
            },
            key="editor_brutos",
        )

        if st.button("💾  Guardar Cambios", type="primary"):
            with st.spinner("Guardando..."):
                ok = guardar_tabla_completa(df_editado)
            if ok:
                st.success("✅ Cambios guardados.")

    st.divider()

    # ════════════════════════════════════════════════════════
    # SECCIÓN C: LIQUIDACIÓN
    # ════════════════════════════════════════════════════════
    st.markdown("## 💰 Liquidación Mensual")

    if st.button("📊  CALCULAR IMPORTES Y GENERAR EXCEL",
                 type="primary", use_container_width=True):
        df_calc = cargar_datos_brutos()
        if df_calc.empty:
            st.error("❌ No hay datos para calcular.")
        else:
            with st.spinner("⏳ Calculando..."):
                df_cierre = calcular_liquidacion(df_calc)
            if df_cierre.empty:
                st.error("❌ No se pudieron calcular los importes.")
            else:
                nc = df_cierre["Conductor"].nunique()
                nd = len(df_cierre)
                st.success(f"✅ **{nd} días** de **{nc} conductor(es)**.")

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("💼 Horas",   f"{df_cierre['IMPORTE HORAS (€)'].sum():,.2f} €")
                m2.metric("📌 Fijos",   f"{df_cierre['IMPORTE FIJOS (€)'].sum():,.2f} €")
                m3.metric("🍽️ Dietas",  f"{df_cierre['IMPORTE DIETAS (€)'].sum():,.2f} €")
                m4.metric("💰 TOTAL",    f"{df_cierre['TOTAL GENERAL A PAGAR (€)'].sum():,.2f} €")

                st.dataframe(df_cierre, use_container_width=True, hide_index=True)

                excel_bytes = generar_excel_bytes(df_cierre)
                nombre_arch = f"cierre_mensual_{date.today().strftime('%Y_%m_%d')}.xlsx"
                st.download_button(
                    label     = "⬇️  DESCARGAR EXCEL CIERRE MENSUAL",
                    data      = excel_bytes,
                    file_name = nombre_arch,
                    mime      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type      = "primary",
                    use_container_width=True,
                )


# ============================================================
# 9. NAVEGACIÓN PRINCIPAL
# ============================================================

def main():
    if "vista_actual" not in st.session_state:
        st.session_state.vista_actual = "conductor"

    with st.sidebar:
        st.markdown("## 🚌 Autocares Alegre")
        st.markdown("---")
        st.markdown("**¿Quién eres?**")
        if st.button("👨‍✈️  Soy Conductor", use_container_width=True,
                     type="primary" if st.session_state.vista_actual == "conductor" else "secondary"):
            st.session_state.vista_actual = "conductor"
            st.rerun()
        if st.button("🔐  Administración", use_container_width=True,
                     type="primary" if st.session_state.vista_actual == "admin" else "secondary"):
            st.session_state.vista_actual = "admin"
            st.rerun()
        st.markdown("---")
        st.caption("v4.0 · Autocares Alegre")

    if st.session_state.vista_actual == "conductor":
        vista_conductor()
    else:
        vista_admin()

if __name__ == "__main__":
    main()

