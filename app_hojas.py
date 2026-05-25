# ============================================================
#  AUTOCARES ALEGRE — Verificación Hojas de Ruta v2.0
# ============================================================
import time
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
    page_title="Hojas de Ruta — Autocares Alegre",
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

# ── EXTRACCIÓN PDF ───────────────────────────────────────────

PROMPT = """Analiza esta hoja de ruta de autocares y extrae los datos.
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

Notas:
- Si hay varios grupos pon uno por cada grupo.
- Si no hay grupos diferenciados pon un solo elemento con grupo "1".
- Las horas pueden aparecer como: "salida: 10:00", "sale del cole: 11:00", "vuelta: 12:15".
- El nº de hoja está junto a "HOJA DE RUTA".
- Ignora las anotaciones manuscritas.
"""

def extraer_pdf(pdf_bytes, filename):
    try:
        model = get_gemini()
        pdf_part = {"mime_type": "application/pdf", "data": base64.b64encode(pdf_bytes).decode()}
        response = model.generate_content([PROMPT, pdf_part])
        text = response.text.strip().replace("```json","").replace("```","").strip()
        data = json.loads(text)
        data["archivo"] = filename
        return data
    except Exception as e:
        return {"error": str(e), "archivo": filename}

# ── EXCEL PASO 1 — solo datos PDF ───────────────────────────

def hojas_a_filas(hojas_pdf):
    """Convierte lista de hojas extraídas a filas planas."""
    filas = []
    for hoja in hojas_pdf:
        if "error" in hoja:
            filas.append({
                "Archivo": hoja.get("archivo",""),
                "Nº Hoja": "ERROR", "Fecha": "", "Cliente": "",
                "Descripción": "", "Grupo": "",
                "Hora Salida": "", "Hora Llegada": "", "Destino": "",
                "Notas": hoja["error"][:80]
            })
            continue
        for grupo in hoja.get("grupos", [{"grupo":"1","hora_salida":"","hora_llegada":"","destino":""}]):
            filas.append({
                "Archivo":      hoja.get("archivo",""),
                "Nº Hoja":      hoja.get("nº_hoja",""),
                "Fecha":        hoja.get("fecha",""),
                "Cliente":      hoja.get("cliente",""),
                "Descripción":  hoja.get("descripcion",""),
                "Grupo":        grupo.get("grupo","1"),
                "Hora Salida":  grupo.get("hora_salida",""),
                "Hora Llegada": grupo.get("hora_llegada",""),
                "Destino":      grupo.get("destino",""),
                "Notas":        ""
            })
    return pd.DataFrame(filas)

def generar_excel_pdf(df):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Hojas de Ruta"

    borde = Border(
        left=Side(style="thin",color="CCCCCC"), right=Side(style="thin",color="CCCCCC"),
        top=Side(style="thin",color="CCCCCC"),  bottom=Side(style="thin",color="CCCCCC")
    )
    al_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al_i = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    for cn, col in enumerate(df.columns, 1):
        c = ws.cell(row=1, column=cn, value=col)
        c.font      = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
        c.fill      = PatternFill("solid", fgColor="1A3A5C")
        c.alignment = al_c
        c.border    = borde
    ws.row_dimensions[1].height = 24

    for rn, (_, row) in enumerate(df.iterrows(), 2):
        bg = "FFF2CC" if str(row.get("Nº Hoja","")) == "ERROR" else "FFFFFF"
        fill = PatternFill("solid", fgColor=bg)
        for cn, val in enumerate(row, 1):
            c = ws.cell(row=rn, column=cn, value=str(val) if val is not None else "")
            c.font      = Font(name="Calibri", size=9)
            c.alignment = al_i
            c.border    = borde
            c.fill      = fill
        ws.row_dimensions[rn].height = 15

    anchos = {"Archivo":18,"Nº Hoja":10,"Fecha":12,"Cliente":28,
              "Descripción":35,"Grupo":7,"Hora Salida":12,
              "Hora Llegada":12,"Destino":25,"Notas":30}
    for cn, col in enumerate(df.columns, 1):
        ws.column_dimensions[get_column_letter(cn)].width = anchos.get(col, 14)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(df.columns))}1"
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# ── EXCEL PASO 2 — verificación ──────────────────────────────

@st.cache_data(ttl=60)
def cargar_registros():
    try:
        r = get_supabase().table("datos_brutos").select(
            "conductor,fecha,hoja_servicio,servicios_horas,tipo_registro"
        ).execute()
        if not r.data:
            return pd.DataFrame()
        df = pd.DataFrame(r.data)
        df["hoja_servicio"] = df["hoja_servicio"].fillna("").astype(str).str.strip()
        return df
    except Exception as e:
        st.error(f"Error cargando conductores: {e}")
        return pd.DataFrame()

def primera_hora(raw):
    try:
        items = json.loads(str(raw or "[]"))
        if items:
            ini = items[0].get("inicio","").strip()
            fin = items[0].get("fin","").strip()
            return (ini if ini != "00:00" else ""), (fin if fin != "00:00" else "")
    except:
        pass
    return "", ""

def comparar(h_pdf, h_cond):
    if not h_pdf or not h_cond: return "sin_dato"
    try:
        t1 = datetime.strptime(h_pdf.strip(), "%H:%M").time()
        t2 = datetime.strptime(h_cond.strip(), "%H:%M").time()
        return "ok" if t1 == t2 else "diferencia"
    except:
        return "sin_dato"

def verificar(df_hojas, df_cond):
    filas = []
    for _, row in df_hojas.iterrows():
        nro     = str(row.get("Nº Hoja","")).strip().lstrip("0")
        h_sal   = str(row.get("Hora Salida","")).strip()
        h_ll    = str(row.get("Hora Llegada","")).strip()

        matches = pd.DataFrame()
        if not df_cond.empty and nro:
            mask = df_cond["hoja_servicio"].str.lstrip("0") == nro
            matches = df_cond[mask]

        if matches.empty:
            filas.append({**row.to_dict(),
                "Conductor": "", "Tipo Registro": "",
                "Salida Conductor": "", "Llegada Conductor": "",
                "Estado": "⚠️ Sin registro del conductor"})
        else:
            for _, reg in matches.iterrows():
                conductor = reg.get("conductor","")
                tipo_r    = str(reg.get("tipo_registro","extra") or "extra")
                h_ini_c, h_fin_c = primera_hora(reg.get("servicios_horas",""))

                if tipo_r == "jornada":
                    estado = "📋 Jornada habitual"
                else:
                    e_sal = comparar(h_sal, h_ini_c)
                    e_ll  = comparar(h_ll, h_fin_c)
                    if e_sal == "ok" and e_ll == "ok":
                        estado = "✅ Coincide"
                    elif e_sal == "sin_dato" or e_ll == "sin_dato":
                        estado = "⚠️ Verificar horario"
                    else:
                        difs = []
                        if e_sal == "diferencia": difs.append(f"Salida: PDF {h_sal} / Cond. {h_ini_c}")
                        if e_ll  == "diferencia": difs.append(f"Llegada: PDF {h_ll} / Cond. {h_fin_c}")
                        estado = "❌ " + " | ".join(difs)

                filas.append({**row.to_dict(),
                    "Conductor": conductor,
                    "Tipo Registro": "Jornada" if tipo_r=="jornada" else "Extra",
                    "Salida Conductor": h_ini_c,
                    "Llegada Conductor": h_fin_c,
                    "Estado": estado})

    return pd.DataFrame(filas)

def generar_excel_verificacion(df):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Verificación"

    borde = Border(
        left=Side(style="thin",color="CCCCCC"), right=Side(style="thin",color="CCCCCC"),
        top=Side(style="thin",color="CCCCCC"),  bottom=Side(style="thin",color="CCCCCC")
    )

    for cn, col in enumerate(df.columns, 1):
        c = ws.cell(row=1, column=cn, value=col)
        c.font      = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
        c.fill      = PatternFill("solid", fgColor="1A3A5C")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = borde
    ws.row_dimensions[1].height = 24

    for rn, (_, row) in enumerate(df.iterrows(), 2):
        estado = str(row.get("Estado",""))
        if   "✅" in estado: bg = "C6EFCE"
        elif "❌" in estado: bg = "FFC7CE"
        elif "📋" in estado: bg = "EBF5FB"
        else:                bg = "FFEB9C"

        for cn, val in enumerate(row, 1):
            c = ws.cell(row=rn, column=cn, value=str(val) if val is not None else "")
            c.font      = Font(name="Calibri", size=9)
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            c.border    = borde
            if cn == len(df.columns):
                c.fill = PatternFill("solid", fgColor=bg)
                c.font = Font(name="Calibri", size=9, bold=True)
        ws.row_dimensions[rn].height = 15

    for cn in range(1, len(df.columns)+1):
        ws.column_dimensions[get_column_letter(cn)].width = 16
    ws.freeze_panes = "A2"
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# ── INTERFAZ ─────────────────────────────────────────────────

def main():
    st.markdown(CSS, unsafe_allow_html=True)
    st.markdown(
        "<h2 style='color:#1a3a5c;margin-bottom:0;'>🚌 Hojas de Ruta — Autocares Alegre</h2>",
        unsafe_allow_html=True
    )
    st.divider()

    # ── PASO 1: Subir y extraer PDFs ────────────────────────
    st.markdown("### Paso 1 — Extraer datos de los PDFs")
    archivos = st.file_uploader(
        "Sube los PDFs de hojas de ruta del programa delsol",
        type=["pdf"],
        accept_multiple_files=True
    )

    if archivos:
        st.markdown(f"**{len(archivos)} PDF(s) cargados**")
        if st.button(f"📄 Extraer datos de {len(archivos)} PDF(s)", type="primary", use_container_width=True):
            hojas_pdf = []
            barra = st.progress(0, text="Iniciando...")
            for i, archivo in enumerate(archivos):
                barra.progress(i / len(archivos), text=f"Leyendo {archivo.name} ({i+1}/{len(archivos)})...")
                datos = extraer_pdf(archivo.read(), archivo.name)
                hojas_pdf.append(datos)
                if i < len(archivos) - 1:
                    time.sleep(60)
            barra.empty()
            df_hojas = hojas_a_filas(hojas_pdf)
            st.session_state["hojas_pdf"]  = hojas_pdf
            st.session_state["df_hojas"]   = df_hojas
            st.session_state.pop("df_verif", None)
            st.rerun()

    # ── Mostrar tabla extraída ───────────────────────────────
    if "df_hojas" in st.session_state:
        df_hojas = st.session_state["df_hojas"]
        st.success(f"✅ {len(df_hojas)} servicio(s) extraídos de los PDFs")

        df_ed = st.data_editor(
            df_hojas,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Archivo":      st.column_config.TextColumn(disabled=True, width="small"),
                "Nº Hoja":      st.column_config.TextColumn(disabled=True, width="small"),
                "Fecha":        st.column_config.TextColumn(disabled=True, width="small"),
                "Cliente":      st.column_config.TextColumn(disabled=True, width="medium"),
                "Descripción":  st.column_config.TextColumn(disabled=True, width="large"),
                "Grupo":        st.column_config.TextColumn(disabled=True, width="small"),
                "Hora Salida":  st.column_config.TextColumn(disabled=True, width="small"),
                "Hora Llegada": st.column_config.TextColumn(disabled=True, width="small"),
                "Destino":      st.column_config.TextColumn(disabled=True, width="medium"),
                "Notas":        st.column_config.TextColumn(width="medium"),
            },
            key="tabla_hojas"
        )
        st.session_state["df_hojas"] = df_ed

        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                "⬇️ Descargar Excel con datos extraídos",
                data=generar_excel_pdf(df_ed),
                file_name=f"hojas_ruta_{date.today().strftime('%Y_%m_%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
        with col2:
            if st.button("🗑️ Limpiar y empezar de nuevo", use_container_width=True):
                for k in ["hojas_pdf","df_hojas","df_verif"]:
                    st.session_state.pop(k, None)
                st.rerun()

        st.divider()

        # ── PASO 2: Verificar con conductores ───────────────
        st.markdown("### Paso 2 — Verificar con registros de conductores")
        st.caption("Cruza los datos del PDF con lo que han registrado los conductores en la app.")

        if st.button("🔍 Verificar con conductores", type="primary", use_container_width=True):
            with st.spinner("Cargando registros de conductores..."):
                df_cond = cargar_registros()
                df_verif = verificar(df_ed, df_cond)
                st.session_state["df_verif"] = df_verif
            st.rerun()

    # ── Mostrar verificación ─────────────────────────────────
    if "df_verif" in st.session_state:
        df_verif = st.session_state["df_verif"]

        total  = len(df_verif)
        n_ok   = df_verif["Estado"].str.contains("✅").sum()
        n_jorn = df_verif["Estado"].str.contains("📋").sum()
        n_warn = df_verif["Estado"].str.contains("⚠️").sum()
        n_err  = df_verif["Estado"].str.contains("❌").sum()

        m1,m2,m3,m4,m5 = st.columns(5)
        m1.metric("Total",       total)
        m2.metric("✅ Coincide", int(n_ok))
        m3.metric("📋 Jornada",  int(n_jorn))
        m4.metric("⚠️ Revisar",  int(n_warn))
        m5.metric("❌ Diferencia",int(n_err))

        filtro = st.radio(
            "Mostrar:",
            ["Todos","Solo pendientes (⚠️ y ❌)","Solo diferencias (❌)"],
            horizontal=True
        )
        if filtro == "Solo pendientes (⚠️ y ❌)":
            df_vista = df_verif[df_verif["Estado"].str.contains("⚠️|❌")]
        elif filtro == "Solo diferencias (❌)":
            df_vista = df_verif[df_verif["Estado"].str.contains("❌")]
        else:
            df_vista = df_verif

        st.dataframe(df_vista, use_container_width=True, hide_index=True)

        st.download_button(
            "⬇️ Descargar Excel de verificación",
            data=generar_excel_verificacion(df_verif),
            file_name=f"verificacion_{date.today().strftime('%Y_%m_%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

if __name__ == "__main__":
    main()
