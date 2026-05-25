# ============================================================
#  AUTOCARES ALEGRE — Verificación Hojas de Ruta v1.0
#  Cruza PDFs de delsol con registros de conductores (App 1)
# ============================================================
import streamlit as st
import pandas as pd
import json
import io
import base64
import google.generativeai as genai
from supabase import create_client
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date, datetime

st.set_page_config(
    page_title="Verificación Hojas de Ruta — Autocares Alegre",
    page_icon="🚌",
    layout="wide",
    initial_sidebar_state="collapsed"
)

CSS = """
<style>
header{visibility:hidden;}
#MainMenu{visibility:hidden;}
footer{visibility:hidden;}
[data-testid="stToolbar"]{display:none;}
[data-testid="stDecoration"]{display:none;}
.block-container{padding-top:1rem !important;}
</style>
"""

# ── CONEXIONES ───────────────────────────────────────────────

@st.cache_resource
def get_supabase():
    return create_client(st.secrets["supabase"]["url"], st.secrets["supabase"]["key"])

@st.cache_resource
def get_gemini():
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    return genai.GenerativeModel("gemini-2.0-flash")

# ── EXTRACCIÓN PDF CON GEMINI ────────────────────────────────

PROMPT_EXTRACCION = """Analiza esta hoja de ruta de autocares y extrae los datos.

Devuelve SOLO el JSON sin texto adicional ni markdown.

Estructura exacta:
{
  "nº_hoja": "número de hoja (solo dígitos, ej: 3872)",
  "fecha": "DD/MM/YYYY",
  "cliente": "nombre completo del cliente",
  "descripcion": "descripción del servicio (título principal)",
  "grupos": [
    {
      "grupo": "1",
      "hora_salida": "HH:MM o vacío",
      "hora_llegada": "HH:MM o vacío",
      "destino": "destino si se menciona"
    }
  ]
}

Notas importantes:
- Si hay varios grupos (Grupo 1, Grupo 2...) pon uno por cada grupo.
- Si no hay grupos diferenciados, pon un solo elemento con grupo "1".
- Las horas pueden aparecer como: "salida: 10:00", "sale del cole: 11:00", "vuelta: 12:15", etc.
- El nº de hoja está junto a "HOJA DE RUTA" en el documento.
- Ignora las anotaciones manuscritas de conductor y matrícula.
"""

def extraer_pdf(pdf_bytes, filename):
    """Extrae datos estructurados de un PDF con Gemini."""
    try:
        model = get_gemini()
        pdf_part = {
            "mime_type": "application/pdf",
            "data": base64.b64encode(pdf_bytes).decode()
        }
        response = model.generate_content([PROMPT_EXTRACCION, pdf_part])
        text = response.text.strip().replace("```json", "").replace("```", "").strip()
        data = json.loads(text)
        data["archivo"] = filename
        return data
    except json.JSONDecodeError as e:
        return {"error": f"JSON inválido: {e}", "archivo": filename, "respuesta_raw": response.text[:300]}
    except Exception as e:
        return {"error": str(e), "archivo": filename}

# ── DATOS DE CONDUCTORES (Supabase App 1) ────────────────────

@st.cache_data(ttl=60)
def cargar_registros():
    """Carga todos los registros de conductores de Supabase."""
    try:
        r = get_supabase().table("datos_brutos").select(
            "id,conductor,fecha,hoja_servicio,servicios_horas,servicios_fijos,dieta,observaciones"
        ).execute()
        if not r.data:
            return pd.DataFrame()
        df = pd.DataFrame(r.data)
        df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce").dt.date
        df["hoja_servicio"] = df["hoja_servicio"].fillna("").astype(str).str.strip()
        return df
    except Exception as e:
        st.error(f"Error cargando conductores: {e}")
        return pd.DataFrame()

def horas_conductor(raw):
    """Extrae texto legible de los servicios por horas."""
    try:
        items = json.loads(str(raw or "[]"))
        partes = []
        for i in items:
            c = i.get("concepto","").strip()
            ini = i.get("inicio","")
            fin = i.get("fin","")
            if c and ini and fin and ini != "00:00":
                partes.append(f"{c} {ini}→{fin}")
            elif c:
                partes.append(c)
        return " | ".join(partes)
    except:
        return ""

def primera_hora(raw):
    """Devuelve (hora_inicio, hora_fin) del primer servicio de horas."""
    try:
        items = json.loads(str(raw or "[]"))
        if items:
            ini = items[0].get("inicio","").strip()
            fin = items[0].get("fin","").strip()
            return ini if ini != "00:00" else "", fin if fin != "00:00" else ""
    except:
        pass
    return "", ""

# ── CRUCE PDF vs CONDUCTOR ───────────────────────────────────

def estado_horario(h_pdf, h_cond):
    """Compara dos horas HH:MM. Devuelve 'ok', 'diferencia' o 'sin_dato'."""
    if not h_pdf or not h_cond:
        return "sin_dato"
    try:
        t1 = datetime.strptime(h_pdf.strip(), "%H:%M").time()
        t2 = datetime.strptime(h_cond.strip(), "%H:%M").time()
        return "ok" if t1 == t2 else "diferencia"
    except:
        return "sin_dato"

def cruzar(hojas_pdf, df_cond):
    """Genera tabla de verificación cruzando PDFs con registros."""
    filas = []

    for hoja in hojas_pdf:
        archivo  = hoja.get("archivo", "")

        # PDF con error de extracción
        if "error" in hoja:
            filas.append({
                "Archivo": archivo, "Nº Hoja": "—", "Fecha PDF": "—",
                "Cliente": "—", "Descripción": "—", "Grupo": "—",
                "Salida PDF": "—", "Llegada PDF": "—",
                "Conductor": "", "Salida Conductor": "", "Llegada Conductor": "",
                "Detalle conductor": "",
                "Estado": f"❌ Error al leer PDF: {hoja['error'][:80]}"
            })
            continue

        nro    = str(hoja.get("nº_hoja","")).lstrip("0") or ""
        nro_orig = str(hoja.get("nº_hoja","")).strip()
        fecha  = hoja.get("fecha","")
        cliente = hoja.get("cliente","")
        desc   = hoja.get("descripcion","")

        # Buscar registros del conductor por nº de hoja
        matches = pd.DataFrame()
        if not df_cond.empty and nro:
            # Comparar sin ceros a la izquierda para robustez
            mask = df_cond["hoja_servicio"].str.lstrip("0") == nro
            matches = df_cond[mask]

        for grupo in hoja.get("grupos", [{"grupo":"1","hora_salida":"","hora_llegada":"","destino":""}]):
            g_num  = grupo.get("grupo","1")
            h_sal  = grupo.get("hora_salida","").strip()
            h_ll   = grupo.get("hora_llegada","").strip()
            destino = grupo.get("destino","")

            if matches.empty:
                # Sin registro del conductor para este nº de hoja
                filas.append({
                    "Archivo": archivo, "Nº Hoja": nro_orig, "Fecha PDF": fecha,
                    "Cliente": cliente, "Descripción": desc, "Grupo": g_num,
                    "Salida PDF": h_sal, "Llegada PDF": h_ll,
                    "Conductor": "", "Salida Conductor": "", "Llegada Conductor": "",
                    "Detalle conductor": "",
                    "Estado": "⚠️ Sin registro del conductor"
                })
            else:
                for _, reg in matches.iterrows():
                    conductor = reg.get("conductor","")
                    detalle   = horas_conductor(reg.get("servicios_horas",""))
                    h_ini_c, h_fin_c = primera_hora(reg.get("servicios_horas",""))

                    e_sal = estado_horario(h_sal, h_ini_c)
                    e_ll  = estado_horario(h_ll, h_fin_c)

                    if e_sal == "ok" and e_ll == "ok":
                        estado = "✅ Coincide"
                    elif e_sal == "sin_dato" or e_ll == "sin_dato":
                        estado = "⚠️ Verificar horario"
                    else:
                        difs = []
                        if e_sal == "diferencia":
                            difs.append(f"Salida: PDF {h_sal} / Cond. {h_ini_c}")
                        if e_ll == "diferencia":
                            difs.append(f"Llegada: PDF {h_ll} / Cond. {h_fin_c}")
                        estado = "❌ " + " | ".join(difs)

                    filas.append({
                        "Archivo": archivo, "Nº Hoja": nro_orig, "Fecha PDF": fecha,
                        "Cliente": cliente, "Descripción": desc, "Grupo": g_num,
                        "Salida PDF": h_sal, "Llegada PDF": h_ll,
                        "Conductor": conductor,
                        "Salida Conductor": h_ini_c, "Llegada Conductor": h_fin_c,
                        "Detalle conductor": detalle,
                        "Estado": estado
                    })

    return pd.DataFrame(filas)

# ── EXCEL ────────────────────────────────────────────────────

def generar_excel(df):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Verificación"

    borde = Border(
        left=Side(style="thin",color="CCCCCC"), right=Side(style="thin",color="CCCCCC"),
        top=Side(style="thin",color="CCCCCC"),  bottom=Side(style="thin",color="CCCCCC")
    )
    al_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al_i = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    # Cabecera
    for cn, col in enumerate(df.columns, 1):
        c = ws.cell(row=1, column=cn, value=col)
        c.font      = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
        c.fill      = PatternFill("solid", fgColor="1A3A5C")
        c.alignment = al_c
        c.border    = borde
    ws.row_dimensions[1].height = 24

    # Filas
    for rn, (_, row) in enumerate(df.iterrows(), 2):
        estado = str(row.get("Estado",""))
        if   "✅" in estado: bg = "C6EFCE"
        elif "❌" in estado: bg = "FFC7CE"
        else:                bg = "FFEB9C"

        for cn, val in enumerate(row, 1):
            c = ws.cell(row=rn, column=cn, value=str(val) if val is not None else "")
            c.font      = Font(name="Calibri", size=9)
            c.alignment = al_i
            c.border    = borde
            if cn == len(df.columns):   # columna Estado → color
                c.fill = PatternFill("solid", fgColor=bg)
                c.font = Font(name="Calibri", size=9, bold=True)
        ws.row_dimensions[rn].height = 15

    # Anchos
    anchos = {
        "Archivo":18, "Nº Hoja":10, "Fecha PDF":12, "Cliente":28,
        "Descripción":32, "Grupo":8,
        "Salida PDF":12, "Llegada PDF":12,
        "Conductor":22, "Salida Conductor":14, "Llegada Conductor":14,
        "Detalle conductor":35, "Estado":38
    }
    for cn, col in enumerate(df.columns, 1):
        ws.column_dimensions[get_column_letter(cn)].width = anchos.get(col, 14)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(df.columns))}1"
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# ── INTERFAZ ─────────────────────────────────────────────────

def main():
    st.markdown(CSS, unsafe_allow_html=True)

    st.markdown(
        "<h2 style='color:#1a3a5c;margin-bottom:0;'>🚌 Verificación Hojas de Ruta</h2>"
        "<p style='color:#666;margin-top:0;'>Cruza los PDFs de delsol con los registros de los conductores</p>",
        unsafe_allow_html=True
    )
    st.divider()

    # ── Subida de PDFs ──────────────────────────────────────
    archivos = st.file_uploader(
        "📄 Sube los PDFs de hojas de ruta (puedes seleccionar varios a la vez)",
        type=["pdf"],
        accept_multiple_files=True
    )

    if not archivos:
        st.info("Sube uno o más PDFs del programa delsol para empezar.")

        st.markdown("#### ¿Cómo funciona?")
        col1, col2, col3 = st.columns(3)
        col1.markdown("**1. Sube los PDFs**\nHojas de ruta descargadas directamente del programa delsol")
        col2.markdown("**2. Cruce automático**\nSe compara con lo que registró cada conductor en la app")
        col3.markdown("**3. Resultado**\n✅ Coincide · ⚠️ Revisar · ❌ Diferencia detectada")
        return

    st.markdown(f"**{len(archivos)} PDF(s) cargados:** {', '.join(f.name for f in archivos[:5])}{'...' if len(archivos)>5 else ''}")
    st.divider()

    col_btn1, col_btn2 = st.columns([2,3])
    with col_btn1:
        procesar = st.button(
            f"🔍 Procesar y cruzar {len(archivos)} PDF(s)",
            type="primary", use_container_width=True
        )
    with col_btn2:
        if st.button("↺ Actualizar datos de conductores", type="secondary", use_container_width=True):
            cargar_registros.clear()
            st.success("Datos actualizados.")

    if procesar:
        hojas_pdf = []
        barra = st.progress(0, text="Iniciando...")

        for i, archivo in enumerate(archivos):
            barra.progress(i / len(archivos), text=f"Leyendo {archivo.name} ({i+1}/{len(archivos)})...")
            datos = extraer_pdf(archivo.read(), archivo.name)
            hojas_pdf.append(datos)

        barra.progress(0.95, text="Cruzando con registros de conductores...")
        df_cond = cargar_registros()
        df_resultado = cruzar(hojas_pdf, df_cond)
        barra.progress(1.0, text="Listo.")
        barra.empty()

        st.session_state["resultado"] = df_resultado
        st.rerun()

    # ── Resultados ──────────────────────────────────────────
    if "resultado" not in st.session_state:
        return

    df = st.session_state["resultado"]

    if df.empty:
        st.warning("No se pudieron extraer datos de los PDFs.")
        return

    # Métricas resumen
    total = len(df)
    n_ok   = df["Estado"].str.contains("✅").sum()
    n_warn = df["Estado"].str.contains("⚠️").sum()
    n_err  = df["Estado"].str.contains("❌").sum()

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total filas", total)
    m2.metric("✅ Coinciden",  int(n_ok))
    m3.metric("⚠️ Revisar",    int(n_warn))
    m4.metric("❌ Diferencias", int(n_err))
    st.divider()

    # Filtro rápido
    filtro = st.radio(
        "Mostrar:",
        ["Todos", "Solo ⚠️ y ❌ (pendientes)", "Solo ❌ diferencias"],
        horizontal=True
    )
    if filtro == "Solo ⚠️ y ❌ (pendientes)":
        df_vista = df[df["Estado"].str.contains("⚠️|❌")]
    elif filtro == "Solo ❌ diferencias":
        df_vista = df[df["Estado"].str.contains("❌")]
    else:
        df_vista = df

    # Tabla editable
    df_ed = st.data_editor(
        df_vista,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Archivo":           st.column_config.TextColumn(disabled=True,  width="small"),
            "Nº Hoja":           st.column_config.TextColumn(disabled=True,  width="small"),
            "Fecha PDF":         st.column_config.TextColumn(disabled=True,  width="small"),
            "Cliente":           st.column_config.TextColumn(disabled=True,  width="medium"),
            "Descripción":       st.column_config.TextColumn(disabled=True,  width="large"),
            "Grupo":             st.column_config.TextColumn(disabled=True,  width="small"),
            "Salida PDF":        st.column_config.TextColumn(disabled=True,  width="small"),
            "Llegada PDF":       st.column_config.TextColumn(disabled=True,  width="small"),
            "Conductor":         st.column_config.TextColumn(width="medium"),
            "Salida Conductor":  st.column_config.TextColumn(disabled=True,  width="small"),
            "Llegada Conductor": st.column_config.TextColumn(disabled=True,  width="small"),
            "Detalle conductor": st.column_config.TextColumn(disabled=True,  width="large"),
            "Estado":            st.column_config.TextColumn(width="medium"),
        },
        key="tabla_verificacion"
    )

    st.divider()
    col_d1, col_d2 = st.columns([2,2])
    with col_d1:
        nombre_arch = f"verificacion_{date.today().strftime('%Y_%m_%d')}.xlsx"
        st.download_button(
            "⬇️ Descargar Excel de verificación",
            data=generar_excel(df_ed),
            file_name=nombre_arch,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
    with col_d2:
        if st.button("🗑️ Limpiar resultados", use_container_width=True):
            del st.session_state["resultado"]
            st.rerun()

if __name__ == "__main__":
    main()
