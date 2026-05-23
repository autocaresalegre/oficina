# ============================================================
#  AUTOCARES ALEGRE — Sistema de Gestión de Servicios Extras
#  Versión DEFINITIVA 3.0 | Backend: Supabase (nube gratuita)
# ============================================================

import streamlit as st
import pandas as pd
from supabase import create_client
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date, datetime, time
import io

# ============================================================
# 1. CONFIGURACIÓN GENERAL DE LA PÁGINA
# ============================================================
st.set_page_config(
    page_title="Autocares Alegre",
    page_icon="🚌",
    layout="centered",
    initial_sidebar_state="auto"
)

# ============================================================
# 2. CONSTANTES: CONTRASEÑA, TARIFAS Y FESTIVOS
# ============================================================

PASSWORD_ADMIN = "alegre2026"

# Tarifas (€)
PRECIOS_FIJOS = {
    "Pilotos": 15.00, "Pescadores": 15.00,
    "Logista": 15.00, "Saludes":    15.00,
    "Boda":    40.00, "Transfer":   20.00,
}
PRECIO_DIETA       = 15.00
PRECIO_HORA_SEMANA = 10.00
PRECIO_HORA_FINDE  = 12.00

# Festivos nacionales España + Comunitat Valenciana 2025-2026
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

# ============================================================
# 3. CSS PERSONALIZADO
# ============================================================
CSS = """
<style>
    .stButton > button {
        height: 3.2rem; font-size: 1rem;
        font-weight: 700; border-radius: 12px;
    }
    .stSelectbox > label, .stDateInput > label,
    .stTextInput > label, .stTimeInput > label,
    .stCheckbox > label span {
        font-size: 1.02rem !important; font-weight: 600 !important;
    }
    .servicio-box {
        background: linear-gradient(135deg, #f0f4ff 0%, #e8eeff 100%);
        border-left: 5px solid #1F4E79;
        padding: 1.2rem 1rem 0.8rem 1rem;
        border-radius: 10px; margin-bottom: 1rem;
    }
</style>
"""

# ============================================================
# 4. CONEXIÓN A SUPABASE
# ============================================================

@st.cache_resource
def get_supabase():
    """
    Crea y cachea la conexión a Supabase.
    Lee las credenciales del archivo secrets.toml.
    Solo se conecta una vez; la conexión dura toda la sesión.
    """
    url = st.secrets["supabase"]["url"]
    key = st.secrets["supabase"]["key"]
    return create_client(url, key)


# ============================================================
# 5. FUNCIONES DE LECTURA Y ESCRITURA EN SUPABASE
# ============================================================

@st.cache_data(ttl=60)
def cargar_conductores():
    """
    Lee la tabla 'conductores' de Supabase y devuelve la lista de nombres.
    Se refresca automáticamente cada 60 segundos.
    """
    try:
        sb = get_supabase()
        resp = sb.table("conductores").select("nombre").order("nombre").execute()
        return [fila["nombre"] for fila in resp.data if fila.get("nombre")]
    except Exception as e:
        st.error(f"❌ Error al cargar conductores: {e}")
        return []


@st.cache_data(ttl=20)
def cargar_datos_brutos():
    """
    Lee la tabla 'datos_brutos' de Supabase y devuelve un DataFrame.
    Se refresca automáticamente cada 20 segundos.
    No incluye la columna 'id' interna de Supabase.
    """
    columnas = [
        "conductor", "fecha", "dieta", "tipo_extra",
        "tipo_fijo", "concepto_destino", "hora_inicio", "hora_fin"
    ]
    try:
        sb = get_supabase()
        resp = sb.table("datos_brutos") \
                 .select("conductor,fecha,dieta,tipo_extra,tipo_fijo,concepto_destino,hora_inicio,hora_fin") \
                 .order("id") \
                 .execute()
        if not resp.data:
            return pd.DataFrame(columns=columnas)
        df = pd.DataFrame(resp.data)
        for col in columnas:
            if col not in df.columns:
                df[col] = None
        return df[columnas]
    except Exception as e:
        st.error(f"❌ Error al cargar datos: {e}")
        return pd.DataFrame(columns=columnas)


def guardar_nuevas_filas(lista_filas):
    """
    Añade nuevas filas a la tabla 'datos_brutos' en Supabase.
    Recibe una lista de diccionarios, uno por cada servicio.
    """
    try:
        sb = get_supabase()
        sb.table("datos_brutos").insert(lista_filas).execute()
        cargar_datos_brutos.clear()   # Invalidar caché para ver los datos nuevos
        return True
    except Exception as e:
        st.error(f"❌ Error al guardar en Supabase: {e}")
        return False


def guardar_datos_brutos_completo(df):
    """
    Reemplaza TODOS los datos de la tabla con el DataFrame editado por la admin.
    Proceso: borrar todo → reinsertar todo.
    Se usa cuando la admin guarda cambios en la tabla de auditoría.
    """
    try:
        sb = get_supabase()

        # Paso 1: Obtener todos los IDs actuales para borrarlos
        ids_resp = sb.table("datos_brutos").select("id").execute()
        if ids_resp.data:
            ids = [fila["id"] for fila in ids_resp.data]
            # Borrar en grupos de 500 para evitar límites de la API
            for i in range(0, len(ids), 500):
                lote = ids[i:i+500]
                sb.table("datos_brutos").delete().in_("id", lote).execute()

        # Paso 2: Reinsertar todo el DataFrame editado
        if not df.empty:
            columnas_bd = [
                "conductor", "fecha", "dieta", "tipo_extra",
                "tipo_fijo", "concepto_destino", "hora_inicio", "hora_fin"
            ]
            filas = []
            for _, row in df.iterrows():
                fila = {}
                for col in columnas_bd:
                    valor = row.get(col)
                    # Convertir NaN/None a None de Python (Supabase lo acepta)
                    if pd.isna(valor) if not isinstance(valor, (bool, str)) else False:
                        fila[col] = None
                    elif col == "dieta":
                        # Asegurar que dieta es booleano real
                        fila[col] = bool(valor) if not isinstance(valor, str) else valor.lower() in ["true", "1"]
                    elif col == "fecha" and hasattr(valor, "strftime"):
                        fila[col] = valor.strftime("%Y-%m-%d")
                    else:
                        fila[col] = str(valor) if valor is not None else None
                filas.append(fila)

            # Insertar en grupos de 500
            for i in range(0, len(filas), 500):
                lote = filas[i:i+500]
                sb.table("datos_brutos").insert(lote).execute()

        cargar_datos_brutos.clear()
        return True
    except Exception as e:
        st.error(f"❌ Error al guardar cambios: {e}")
        return False


# ============================================================
# 6. FUNCIONES AUXILIARES DE CÁLCULO
# ============================================================

def es_dia_especial(fecha):
    """True si la fecha es sábado, domingo o festivo → se cobra 12 €/hora."""
    if isinstance(fecha, datetime):
        fecha = fecha.date()
    elif isinstance(fecha, str):
        try:
            fecha = pd.to_datetime(fecha).date()
        except Exception:
            return False
    return fecha.weekday() >= 5 or fecha in FESTIVOS


def calcular_horas_decimal(hora_ini, hora_fin):
    """
    Calcula la diferencia exacta en horas decimales.
    Acepta strings 'HH:MM' o bien objetos time de Python.
    Ejemplo: 08:00 → 10:30 = 2.5 horas exactas.
    """
    try:
        if isinstance(hora_ini, str):
            hora_ini = datetime.strptime(hora_ini.strip(), "%H:%M").time()
        if isinstance(hora_fin, str):
            hora_fin = datetime.strptime(hora_fin.strip(), "%H:%M").time()
        inicio = datetime.combine(date.today(), hora_ini)
        fin    = datetime.combine(date.today(), hora_fin)
        if fin <= inicio:
            return 0.0
        return round((fin - inicio).total_seconds() / 3600, 4)
    except Exception:
        return 0.0


# ============================================================
# 7. LÓGICA DE LIQUIDACIÓN MENSUAL
# ============================================================

def calcular_liquidacion(df):
    """
    Genera el DataFrame del cierre mensual a partir de los datos brutos.

    Reglas:
    - Agrupa por conductor + fecha → una fila de resultado por día
    - Dieta: UNA SOLA de 15 € por día (aunque haya varios servicios ese día)
    - Horas exactas en decimal × tarifa según día de la semana / festivo
    - Fijos: suma de todos los servicios fijos del día
    - Resultado ordenado alfabéticamente por conductor y cronológicamente por fecha
    """
    if df.empty:
        return pd.DataFrame()

    df = df.copy()

    # Renombrar columnas de Supabase (minúsculas) a las del cálculo
    df.columns = [c.capitalize() if c != "tipo_extra" else "Tipo_Extra" for c in df.columns]
    df = df.rename(columns={
        "Conductor": "Conductor",
        "Fecha": "Fecha",
        "Dieta": "Dieta",
        "Tipo_extra": "Tipo_Extra",
        "Tipo_fijo": "Tipo_Fijo",
        "Concepto_destino": "Concepto_Destino",
        "Hora_inicio": "Hora_Inicio",
        "Hora_fin": "Hora_Fin",
    })

    # Normalizar tipos
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce").dt.date

    def a_bool(v):
        if isinstance(v, bool):
            return v
        return str(v).strip().lower() in ["true", "1", "sí", "si"]

    df["Dieta"] = df["Dieta"].fillna(False).apply(a_bool)

    resultados = []

    for (conductor, fecha), grupo in df.groupby(["Conductor", "Fecha"], sort=True):
        if pd.isna(fecha):
            continue

        es_especial = es_dia_especial(fecha)
        tarifa_hora = PRECIO_HORA_FINDE if es_especial else PRECIO_HORA_SEMANA
        tiene_dieta = bool(grupo["Dieta"].any())
        imp_dieta   = PRECIO_DIETA if tiene_dieta else 0.0
        conceptos   = []
        horas_total = 0.0
        imp_horas   = 0.0
        imp_fijos   = 0.0

        for _, fila in grupo.iterrows():
            tipo_extra = str(fila.get("Tipo_Extra", "")).strip()

            if tipo_extra == "Servicio Fijo":
                tipo_fijo = str(fila.get("Tipo_Fijo", "")).strip()
                if tipo_fijo in PRECIOS_FIJOS:
                    imp_fijos += PRECIOS_FIJOS[tipo_fijo]
                    conceptos.append(tipo_fijo)

            elif tipo_extra == "Por Horas":
                concepto = str(fila.get("Concepto_Destino", "")).strip()
                h_ini = fila.get("Hora_Inicio")
                h_fin = fila.get("Hora_Fin")
                if h_ini and h_fin and str(h_ini) not in ("", "nan", "None"):
                    horas       = calcular_horas_decimal(str(h_ini), str(h_fin))
                    horas_total += horas
                    imp_horas   += horas * tarifa_hora
                if concepto and concepto not in ("", "nan", "None"):
                    conceptos.append(concepto)

        total_dia = round(imp_horas + imp_fijos + imp_dieta, 2)
        resultados.append({
            "Conductor":                 conductor,
            "Fecha":                     fecha,
            "Conceptos del Día":         " | ".join(conceptos) if conceptos else "—",
            "Horas Extras Totales":      round(horas_total, 2),
            "IMPORTE HORAS (€)":         round(imp_horas,   2),
            "IMPORTE FIJOS (€)":         round(imp_fijos,   2),
            "IMPORTE DIETAS (€)":        round(imp_dieta,   2),
            "TOTAL GENERAL A PAGAR (€)": total_dia,
        })

    df_res = pd.DataFrame(resultados)
    return df_res.sort_values(["Conductor", "Fecha"]).reset_index(drop=True)


# ============================================================
# 8. GENERACIÓN DEL EXCEL DE CIERRE (CON FORMATO VISUAL)
# ============================================================

def generar_excel_bytes(df_cierre):
    """
    Genera el Excel del cierre mensual con formato visual profesional:
    - Cabecera azul oscuro con letras blancas
    - Filas de fin de semana / festivo en amarillo
    - Columna TOTAL en verde y negrita
    - Fila de TOTAL GLOBAL al final con sumas automáticas
    - Primera fila congelada y filtros automáticos activos
    """
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cierre Mensual"

    C_AZUL  = "1F4E79"
    C_AMAR  = "FFF2CC"
    C_VERDE = "C6EFCE"

    borde = Border(
        left=Side(style="thin",   color="AAAAAA"),
        right=Side(style="thin",  color="AAAAAA"),
        top=Side(style="thin",    color="AAAAAA"),
        bottom=Side(style="thin", color="AAAAAA"),
    )
    al_c = Alignment(horizontal="center",  vertical="center", wrap_text=True)
    al_d = Alignment(horizontal="right",   vertical="center")
    al_i = Alignment(horizontal="left",    vertical="center", wrap_text=True)

    # ── Fila 1: Cabeceras ───────────────────────────────────
    cabeceras = [
        "Conductor", "Fecha", "Conceptos del Día",
        "Horas Extras\nTotales", "IMPORTE\nHORAS (€)",
        "IMPORTE\nFIJOS (€)", "IMPORTE\nDIETAS (€)",
        "TOTAL GENERAL\nA PAGAR (€)",
    ]
    for cn, titulo in enumerate(cabeceras, 1):
        c = ws.cell(row=1, column=cn, value=titulo)
        c.font      = Font(bold=True, color="FFFFFF", size=11, name="Calibri")
        c.fill      = PatternFill("solid", fgColor=C_AZUL)
        c.alignment = al_c
        c.border    = borde
    ws.row_dimensions[1].height = 42

    # ── Filas de datos ──────────────────────────────────────
    for idx, fila in df_cierre.iterrows():
        fxls  = idx + 2
        fecha = fila["Fecha"]
        fondo = es_dia_especial(fecha)
        fill  = PatternFill("solid", fgColor=C_AMAR) if fondo else None

        valores = [
            fila["Conductor"], fecha,
            fila["Conceptos del Día"],
            fila["Horas Extras Totales"],
            fila["IMPORTE HORAS (€)"],
            fila["IMPORTE FIJOS (€)"],
            fila["IMPORTE DIETAS (€)"],
            fila["TOTAL GENERAL A PAGAR (€)"],
        ]
        for cn, valor in enumerate(valores, 1):
            c = ws.cell(row=fxls, column=cn, value=valor)
            c.border = borde
            if fill:
                c.fill = fill
            if cn == 1:
                c.font = Font(bold=True, size=10, name="Calibri")
                c.alignment = al_i
            elif cn == 2:
                c.number_format = "DD/MM/YYYY"
                c.alignment = al_c
                c.font = Font(size=10, name="Calibri")
            elif cn == 3:
                c.alignment = al_i
                c.font = Font(size=10, name="Calibri")
            elif cn == 4:
                c.number_format = "0.00"
                c.alignment = al_c
                c.font = Font(size=10, name="Calibri")
            elif cn in (5, 6, 7):
                c.number_format = '#,##0.00 "€"'
                c.alignment = al_d
                c.font = Font(size=10, name="Calibri")
            elif cn == 8:
                c.number_format = '#,##0.00 "€"'
                c.alignment = al_d
                c.font = Font(bold=True, size=10, name="Calibri")
                if not fondo:
                    c.fill = PatternFill("solid", fgColor=C_VERDE)
        ws.row_dimensions[fxls].height = 20

    # ── Fila TOTAL GLOBAL ───────────────────────────────────
    ft = len(df_cierre) + 2
    ct = ws.cell(row=ft, column=1, value="✔  TOTAL GLOBAL")
    ct.font      = Font(bold=True, size=12, color="FFFFFF", name="Calibri")
    ct.fill      = PatternFill("solid", fgColor=C_AZUL)
    ct.alignment = al_c
    ct.border    = borde
    for cn in (2, 3):
        c = ws.cell(row=ft, column=cn)
        c.fill   = PatternFill("solid", fgColor=C_AZUL)
        c.border = borde

    l4 = get_column_letter(4)
    c4 = ws.cell(row=ft, column=4, value=f"=SUM({l4}2:{l4}{ft-1})")
    c4.number_format = "0.00"
    c4.font  = Font(bold=True, size=11, name="Calibri")
    c4.fill  = PatternFill("solid", fgColor=C_VERDE)
    c4.alignment = al_c
    c4.border = borde

    for cn in (5, 6, 7, 8):
        letra = get_column_letter(cn)
        c = ws.cell(row=ft, column=cn, value=f"=SUM({letra}2:{letra}{ft-1})")
        c.number_format = '#,##0.00 "€"'
        c.font  = Font(bold=True, size=11, name="Calibri")
        c.fill  = PatternFill("solid", fgColor=C_VERDE)
        c.alignment = al_d
        c.border = borde
    ws.row_dimensions[ft].height = 28

    # ── Formato final ────────────────────────────────────────
    anchos = {1: 28, 2: 13, 3: 52, 4: 14, 5: 17, 6: 17, 7: 17, 8: 22}
    for cn, ancho in anchos.items():
        ws.column_dimensions[get_column_letter(cn)].width = ancho
    ws.freeze_panes  = "A2"
    ws.auto_filter.ref = f"A1:H{ft-1}"

    wb.save(output)
    output.seek(0)
    return output.getvalue()


# ============================================================
# 9. VISTA DEL CONDUCTOR (PANTALLA MÓVIL)
# ============================================================

def vista_conductor():
    st.markdown(CSS, unsafe_allow_html=True)
    st.title("🚌 Autocares Alegre")
    st.markdown("#### Registro de Servicios Extras")
    st.divider()

    # Pantalla de confirmación tras enviar
    if st.session_state.get("envio_ok"):
        info = st.session_state.get("envio_info", {})
        st.success(
            f"✅ **¡Registro enviado correctamente!**\n\n"
            f"👤 **{info.get('nombre','')}** — "
            f"📅 {info.get('fecha','')} — "
            f"📋 {info.get('n',0)} servicio(s) guardado(s)."
        )
        st.balloons()
        if st.button("➕  Hacer Otro Registro", type="primary", use_container_width=True):
            st.session_state.envio_ok    = False
            st.session_state.n_servicios = 1
            st.rerun()
        return

    # ── Campo 1: Nombre ─────────────────────────────────────
    conductores = cargar_conductores()
    nombre_sel  = st.selectbox(
        "👤  Selecciona tu Nombre",
        options=["— Selecciona tu nombre —"] + conductores
    )

    # ── Campo 2: Fecha ──────────────────────────────────────
    fecha_sel = st.date_input(
        "📅  Selecciona la Fecha del Servicio",
        value=date.today(), format="DD/MM/YYYY"
    )

    # ── Campo 3: Dieta ──────────────────────────────────────
    st.markdown("---")
    dieta_sel = st.checkbox("🍽️  ¿Te corresponde Dieta hoy?", value=False)
    st.markdown("---")

    # ── Servicios Extras ────────────────────────────────────
    st.markdown("### 📋 Servicios Extras Realizados Hoy")
    st.caption("Añade todos los servicios que hayas hecho hoy, uno a uno.")

    if "n_servicios" not in st.session_state:
        st.session_state.n_servicios = 1

    servicios_capturados = []

    for i in range(st.session_state.n_servicios):
        st.markdown('<div class="servicio-box">', unsafe_allow_html=True)
        st.markdown(f"**📌 Servicio Extra #{i+1}**")

        tipo_extra = st.selectbox(
            "Tipo de Extra",
            options=["— Elige el tipo —", "Servicio Fijo", "Por Horas"],
            key=f"tipo_{i}"
        )
        datos = {"tipo": tipo_extra, "tipo_fijo": None,
                 "concepto": None, "hora_ini": None, "hora_fin": None}

        if tipo_extra == "Servicio Fijo":
            tf = st.selectbox(
                "📁  ¿Qué tipo de servicio fijo?",
                options=["— Elige el servicio —", "Pilotos", "Pescadores",
                         "Logista", "Saludes", "Boda", "Transfer"],
                key=f"fijo_{i}"
            )
            datos["tipo_fijo"] = tf

        elif tipo_extra == "Por Horas":
            concepto_inp = st.text_input(
                "🗺️  Concepto / Destino  *(obligatorio)*",
                placeholder="Ej: Excursión a Gandía / Vigilancia en Base",
                key=f"concepto_{i}"
            )
            c1, c2 = st.columns(2)
            with c1:
                h_ini = st.time_input("🕐 Hora Inicio", value=time(8,0),  step=300, key=f"h_ini_{i}")
            with c2:
                h_fin = st.time_input("🕔 Hora Fin",    value=time(16,0), step=300, key=f"h_fin_{i}")
            datos["concepto"] = concepto_inp
            datos["hora_ini"] = h_ini
            datos["hora_fin"] = h_fin

        servicios_capturados.append(datos)
        st.markdown("</div>", unsafe_allow_html=True)
        st.markdown("")

    if st.button("➕  Añadir Otro Servicio Extra", type="secondary", use_container_width=True):
        st.session_state.n_servicios += 1
        st.rerun()

    st.divider()

    if st.button("📤  ENVIAR REGISTRO DIARIO", type="primary", use_container_width=True):
        errores      = []
        servicios_ok = []

        if nombre_sel == "— Selecciona tu nombre —":
            errores.append("❌ Debes seleccionar tu nombre.")

        for i, s in enumerate(servicios_capturados):
            n = i + 1
            if s["tipo"] == "— Elige el tipo —":
                errores.append(f"❌ Servicio #{n}: Elige el tipo de extra.")
            elif s["tipo"] == "Servicio Fijo":
                if not s["tipo_fijo"] or s["tipo_fijo"] == "— Elige el servicio —":
                    errores.append(f"❌ Servicio #{n}: Elige el tipo de servicio fijo.")
                else:
                    servicios_ok.append(s)
            elif s["tipo"] == "Por Horas":
                if not s["concepto"] or not s["concepto"].strip():
                    errores.append(f"❌ Servicio #{n}: El campo Concepto/Destino es obligatorio.")
                elif s["hora_fin"] <= s["hora_ini"]:
                    errores.append(f"❌ Servicio #{n}: La hora de fin debe ser posterior a la de inicio.")
                else:
                    servicios_ok.append(s)

        if not servicios_ok and not errores:
            errores.append("❌ Añade al menos un servicio extra válido.")

        if errores:
            for msg in errores:
                st.error(msg)
        else:
            nuevas_filas = []
            for s in servicios_ok:
                nuevas_filas.append({
                    "conductor":        nombre_sel,
                    "fecha":            fecha_sel.strftime("%Y-%m-%d"),
                    "dieta":            dieta_sel,
                    "tipo_extra":       s["tipo"],
                    "tipo_fijo":        s["tipo_fijo"] if s["tipo"] == "Servicio Fijo" else None,
                    "concepto_destino": s["concepto"]  if s["tipo"] == "Por Horas"    else None,
                    "hora_inicio":      s["hora_ini"].strftime("%H:%M") if s["hora_ini"] else None,
                    "hora_fin":         s["hora_fin"].strftime("%H:%M") if s["hora_fin"] else None,
                })

            with st.spinner("Enviando registro..."):
                ok = guardar_nuevas_filas(nuevas_filas)

            if ok:
                st.session_state.envio_ok   = True
                st.session_state.envio_info = {
                    "nombre": nombre_sel,
                    "fecha":  fecha_sel.strftime("%d/%m/%Y"),
                    "n":      len(servicios_ok),
                }
                st.session_state.n_servicios = 1
                st.rerun()


# ============================================================
# 10. VISTA DE ADMINISTRACIÓN
# ============================================================

def vista_admin():
    st.markdown(CSS, unsafe_allow_html=True)

    if "admin_ok" not in st.session_state:
        st.session_state.admin_ok = False

    if not st.session_state.admin_ok:
        st.title("🔐 Panel de Administración")
        st.markdown("#### Autocares Alegre")
        st.divider()
        pw = st.text_input("Contraseña", type="password",
                           placeholder="Introduce la contraseña...")
        if st.button("▶  Entrar", type="primary", use_container_width=True):
            if pw == PASSWORD_ADMIN:
                st.session_state.admin_ok = True
                st.rerun()
            else:
                st.error("❌ Contraseña incorrecta.")
        return

    # ── Cabecera ─────────────────────────────────────────────
    c1, c2 = st.columns([5, 1])
    with c1:
        st.title("📊 Administración — Autocares Alegre")
    with c2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🚪 Salir"):
            st.session_state.admin_ok = False
            st.rerun()
    st.divider()

    # ════════════════════════════════════════════════════════
    # SECCIÓN A: TABLA DE AUDITORÍA
    # ════════════════════════════════════════════════════════
    st.markdown("## 📋 Tabla de Auditoría")
    st.markdown("Haz clic en cualquier celda para editarla. "
                "Pulsa **Guardar Cambios** cuando termines.")

    df_brutos = cargar_datos_brutos()

    if df_brutos.empty:
        st.info("ℹ️ Todavía no hay registros. "
                "Los conductores deben enviar sus servicios primero.")
    else:
        df_editado = st.data_editor(
            df_brutos,
            use_container_width=True,
            num_rows="dynamic",
            hide_index=True,
            column_config={
                "conductor":        st.column_config.TextColumn("Conductor",       width="medium", required=True),
                "fecha":            st.column_config.DateColumn("Fecha",           format="DD/MM/YYYY"),
                "dieta":            st.column_config.CheckboxColumn("🍽️ Dieta",    default=False),
                "tipo_extra":       st.column_config.SelectboxColumn("Tipo Extra", options=["Servicio Fijo","Por Horas"]),
                "tipo_fijo":        st.column_config.SelectboxColumn("Tipo Fijo",  options=["Pilotos","Pescadores","Logista","Saludes","Boda","Transfer"]),
                "concepto_destino": st.column_config.TextColumn("Concepto / Destino", width="large"),
                "hora_inicio":      st.column_config.TextColumn("Hora Inicio\n(HH:MM)", width="small"),
                "hora_fin":         st.column_config.TextColumn("Hora Fin\n(HH:MM)",    width="small"),
            },
            key="editor_brutos",
        )

        if st.button("💾  Guardar Cambios Manuales", type="primary"):
            with st.spinner("Guardando en Supabase..."):
                ok = guardar_datos_brutos_completo(df_editado)
            if ok:
                st.success("✅ Cambios guardados correctamente.")

    st.divider()

    # ════════════════════════════════════════════════════════
    # SECCIÓN B: LIQUIDACIÓN MENSUAL
    # ════════════════════════════════════════════════════════
    st.markdown("## 💰 Liquidación Mensual")
    st.markdown("Cuando hayas revisado todos los datos, pulsa el botón "
                "para calcular los importes y descargar el Excel final.")

    if st.button("📊  CALCULAR IMPORTES Y GENERAR EXCEL RESUMEN",
                 type="primary", use_container_width=True):

        df_calcular = cargar_datos_brutos()

        if df_calcular.empty:
            st.error("❌ No hay datos para calcular. "
                     "Los conductores deben enviar sus servicios primero.")
        else:
            with st.spinner("⏳ Calculando importes..."):
                df_cierre = calcular_liquidacion(df_calcular)

            if df_cierre.empty:
                st.error("❌ No se pudieron calcular los importes. "
                         "Revisa que los registros están bien rellenados.")
            else:
                nc = df_cierre["Conductor"].nunique()
                nd = len(df_cierre)
                st.success(f"✅ ¡Listo! **{nd} días** de **{nc} conductor(es)**.")

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("💼 Horas",   f"{df_cierre['IMPORTE HORAS (€)'].sum():,.2f} €")
                m2.metric("📌 Fijos",   f"{df_cierre['IMPORTE FIJOS (€)'].sum():,.2f} €")
                m3.metric("🍽️ Dietas",  f"{df_cierre['IMPORTE DIETAS (€)'].sum():,.2f} €")
                m4.metric("💰 TOTAL",    f"{df_cierre['TOTAL GENERAL A PAGAR (€)'].sum():,.2f} €")

                st.markdown("### Vista previa")
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
                st.caption(f"📁 El archivo **{nombre_arch}** incluye colores, "
                           "filtros y totales automáticos.")


# ============================================================
# 11. NAVEGACIÓN PRINCIPAL
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
        st.caption("v3.0 · Autocares Alegre")

    if st.session_state.vista_actual == "conductor":
        vista_conductor()
    else:
        vista_admin()


if __name__ == "__main__":
    main()
