# ============================================================
#  AUTOCARES ALEGRE — Hojas de Ruta v2.0 (sin API externa)
# ============================================================
import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
import json
from supabase import create_client
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date, datetime

st.set_page_config(page_title="Hojas de Ruta — Autocares Alegre",
                   page_icon="🚌", layout="wide",
                   initial_sidebar_state="collapsed")

CSS = """<style>
header{visibility:hidden;}#MainMenu{visibility:hidden;}
footer{visibility:hidden;}[data-testid="stToolbar"]{display:none;}
[data-testid="stDecoration"]{display:none;}
.block-container{padding-top:1rem !important;}
</style>"""

# ── SUPABASE ─────────────────────────────────────────────────
@st.cache_resource
def get_supabase():
    return create_client(st.secrets["supabase"]["url"], st.secrets["supabase"]["key"])

@st.cache_data(ttl=60)
def cargar_registros():
    try:
        r = get_supabase().table("datos_brutos").select(
            "conductor,fecha,hoja_servicio,servicios_horas,tipo_registro"
        ).execute()
        if not r.data: return pd.DataFrame()
        df = pd.DataFrame(r.data)
        df["hoja_servicio"] = df["hoja_servicio"].fillna("").astype(str).str.strip()
        return df
    except Exception as e:
        st.error(f"Error cargando conductores: {e}"); return pd.DataFrame()

# ── PARSER PDF ───────────────────────────────────────────────
def parse_hora(texto):
    if not texto: return ""
    texto = re.sub(r'_+', '', str(texto)).strip().lower()
    texto = texto.replace('am','').replace('pm','').replace('h','').strip()
    texto = texto.replace('.', ':')
    try:
        p = texto.split(':')
        return f"{int(p[0]):02d}:{int(p[1]) if len(p)>1 else 0:02d}"
    except: return ""

def extraer_pdf(pdf_bytes, filename):
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            textos = []; nro_hoja = fecha = ""
            for page in pdf.pages:
                t = page.extract_text() or ""
                textos.append(t)
                if not nro_hoja:
                    m = re.search(r'HOJA DE RUTA\s+\d+\s+(\d+)\s+\d+\s+(\d{2}/\d{2}/\d{4})', t)
                    if m: nro_hoja, fecha = m.group(1), m.group(2)
            texto = "\n".join(textos)
            lines = texto.split('\n')

            # Cliente
            cliente = ""
            for i, l in enumerate(lines):
                if 'AUTOCARES ALEGRE' in l and i+1 < len(lines):
                    nx = lines[i+1].strip()
                    if nx and 'TOMAS SANZ' not in nx: cliente = nx; break

            # Descripción
            desc = ""
            m_salon = re.search(r'[Ss]al[oó]n.*?[Dd]estino[:\s•]+([^\n•]+)', texto)
            if m_salon:
                desc = m_salon.group(1).strip()
            else:
                m = re.search(r'DESCRIPCIÓN\s+CANTIDAD.*?\n(.+?)(?=TIPO\s+IMPORTE)', texto, re.DOTALL)
                if m:
                    dlines = []
                    for l in m.group(1).split('\n'):
                        l = l.strip()
                        if (l and not re.match(r'^[\d.,\s€]+$', l)
                                and '@' not in l and 'tel:' not in l.lower()
                                and 'contacto' not in l.lower()
                                and not re.match(r'^PT\d+', l)
                                and not re.match(r'^\d+\s+€', l)
                                and 'DNI' not in l and 'Teléfono' not in l):
                            dlines.append(l)
                    desc = " / ".join(dlines[:2])

            # Plazas globales (para excursiones y servicios simples)
            plazas_global = ""
            m_pax = re.search(r'(\d+)\s*PAX', texto, re.IGNORECASE)
            m_plz = re.search(r'(\d+)\s*PLAZAS', texto, re.IGNORECASE)
            m_npl = re.search(r'[Nn]º\s*plazas\s*totales[:\s_]*(\d+)', texto)
            if m_pax:   plazas_global = m_pax.group(1)
            elif m_plz: plazas_global = m_plz.group(1)
            elif m_npl: plazas_global = m_npl.group(1)

            # Grupos/buses
            grupos = []
            # Solo activa modo buses si hay "BUS 1:" o "•BUS 1:" etc. (no MICROBUS)
            buses_check = re.findall(r'(?:^|[•\n\s])BUS\s+(\d{1,2})\s*[:(]', texto, re.IGNORECASE|re.MULTILINE)
            if buses_check:
                ida_sec = re.search(r'(?:AUTOBUSES DE )?IDA.*?(?=(?:AUTOBUSES DE )?REGRESO|$)', texto, re.DOTALL|re.IGNORECASE)
                reg_sec = re.search(r'(?:AUTOBUSES DE )?REGRESO.*', texto, re.DOTALL|re.IGNORECASE)

                buses_ida = re.findall(
                    r'BUS\s+(\d+).*?\(Plazas[:\s]+_*(\d*)\D*\).*?Hora[:\s]+([0-9_.:\-hHaAmM]+).*?recogida[:\s]+([^\n•]+)',
                    ida_sec.group(0) if ida_sec else "", re.IGNORECASE)
                buses_reg = re.findall(
                    r'BUS\s+(\d+).*?\(Plazas[:\s]+_*(\d*)\D*\).*?Hora[:\s]+([0-9_.:\-hHaAmM]+).*?[Dd]estino[:\s]+([^\n•]+)',
                    reg_sec.group(0) if reg_sec else "", re.IGNORECASE)

                def limpiar(s): return re.sub(r'_+','',s).strip().strip('_ ').strip()

                bd = {}
                for num, plazas, hora, recogida in buses_ida:
                    n = int(num)
                    if n not in bd: bd[n] = {"grupo":str(n),"plazas":plazas,"hora_salida":"","hora_llegada":"","destino":"","destino_regreso":""}
                    bd[n]["hora_salida"] = parse_hora(hora)
                    bd[n]["plazas"] = plazas
                    bd[n]["destino"] = limpiar(recogida)[:60]
                for num, plazas, hora, destino in buses_reg:
                    n = int(num)
                    if n not in bd: bd[n] = {"grupo":str(n),"plazas":plazas,"hora_salida":"","hora_llegada":"","destino":"","destino_regreso":""}
                    bd[n]["hora_llegada"] = parse_hora(hora)
                    bd[n]["destino_regreso"] = limpiar(destino)[:60]
                    if not bd[n]["plazas"]: bd[n]["plazas"] = plazas
                grupos = [bd[k] for k in sorted(bd)]
            else:
                h_sal = h_reg = dest = ""
                m_s = re.search(r'SALIDA:\s*(\d{1,2}[:.]\d{2})\s*([^\n]+)?', texto)
                m_r = re.search(r'REGRESO:\s*[^\n]*?(\d{1,2}[:.]\d{2})', texto)
                if m_s:
                    h_sal = parse_hora(m_s.group(1))
                    dest = (m_s.group(2) or "").strip()[:60]
                if m_r:
                    h_reg = parse_hora(m_r.group(1))
                grupos = [{"grupo":"1","plazas":plazas_global,"hora_salida":h_sal,"hora_llegada":h_reg,"destino":dest,"destino_regreso":""}]

            if not grupos:
                grupos = [{"grupo":"1","plazas":plazas_global,"hora_salida":"","hora_llegada":"","destino":"","destino_regreso":""}]

            return {"archivo":filename,"nº_hoja":nro_hoja,"fecha":fecha,
                    "cliente":cliente,"descripcion":desc,"grupos":grupos}
    except Exception as e:
        return {"error":str(e),"archivo":filename}

# ── CONVERSIÓN A FILAS ───────────────────────────────────────
def hojas_a_filas(hojas):
    filas = []
    for h in hojas:
        if "error" in h:
            filas.append({"Archivo":h.get("archivo",""),"Nº Hoja":"ERROR","Fecha":"",
                          "Cliente":"","Descripción":"","Bus/Grupo":"","Plazas":"",
                          "Hora Salida":"","Punto Recogida":"",
                          "Hora Llegada":"","Destino Regreso":"",
                          "Notas":h["error"][:80]}); continue
        for g in h.get("grupos",[{"grupo":"1","plazas":"","hora_salida":"","hora_llegada":"","destino":"","destino_regreso":""}]):
            filas.append({
                "Archivo":         h.get("archivo",""),
                "Nº Hoja":         h.get("nº_hoja",""),
                "Fecha":           h.get("fecha",""),
                "Cliente":         h.get("cliente",""),
                "Descripción":     h.get("descripcion",""),
                "Bus/Grupo":       g.get("grupo","1"),
                "Plazas":          g.get("plazas",""),
                "Hora Salida":     g.get("hora_salida",""),
                "Punto Recogida":  g.get("destino",""),
                "Hora Llegada":    g.get("hora_llegada",""),
                "Destino Regreso": g.get("destino_regreso",""),
                "Notas":           ""
            })
    return pd.DataFrame(filas)

# ── VERIFICACIÓN ─────────────────────────────────────────────
def primera_hora(raw):
    try:
        items = json.loads(str(raw or "[]"))
        if items:
            ini = items[0].get("inicio","").strip()
            fin = items[0].get("fin","").strip()
            return (ini if ini != "00:00" else ""), (fin if fin != "00:00" else "")
    except: pass
    return "", ""

def comparar(h1, h2):
    if not h1 or not h2: return "sin_dato"
    try:
        t1 = datetime.strptime(h1.strip(),"%H:%M").time()
        t2 = datetime.strptime(h2.strip(),"%H:%M").time()
        return "ok" if t1==t2 else "diferencia"
    except: return "sin_dato"

def verificar(df_hojas, df_cond):
    filas = []
    for _, row in df_hojas.iterrows():
        nro = str(row.get("Nº Hoja","")).strip().lstrip("0")
        h_sal = str(row.get("Hora Salida","")).strip()
        h_ll  = str(row.get("Hora Llegada","")).strip()
        matches = pd.DataFrame()
        if not df_cond.empty and nro:
            mask = df_cond["hoja_servicio"].str.lstrip("0") == nro
            matches = df_cond[mask]
        if matches.empty:
            filas.append({**row.to_dict(),"Conductor":"","Tipo":"",
                          "Salida Conductor":"","Llegada Conductor":"",
                          "Estado":"⚠️ Sin registro del conductor"})
        else:
            for _, reg in matches.iterrows():
                conductor = reg.get("conductor","")
                tipo_r    = str(reg.get("tipo_registro","extra") or "extra")
                h_ini_c, h_fin_c = primera_hora(reg.get("servicios_horas",""))
                if tipo_r == "jornada":
                    estado = "📋 Jornada habitual"
                else:
                    e_s = comparar(h_sal, h_ini_c)
                    e_l = comparar(h_ll, h_fin_c)
                    if e_s=="ok" and e_l=="ok": estado = "✅ Coincide"
                    elif e_s=="sin_dato" or e_l=="sin_dato": estado = "⚠️ Verificar horario"
                    else:
                        difs = []
                        if e_s=="diferencia": difs.append(f"Salida: PDF {h_sal}/Cond. {h_ini_c}")
                        if e_l=="diferencia": difs.append(f"Llegada: PDF {h_ll}/Cond. {h_fin_c}")
                        estado = "❌ "+" | ".join(difs)
                filas.append({**row.to_dict(),"Conductor":conductor,
                              "Tipo":"Jornada" if tipo_r=="jornada" else "Extra",
                              "Salida Conductor":h_ini_c,"Llegada Conductor":h_fin_c,
                              "Estado":estado})
    return pd.DataFrame(filas)

# ── EXCEL ────────────────────────────────────────────────────
def borde_thin():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def estilo_cabecera(c):
    c.font = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
    c.fill = PatternFill("solid", fgColor="1A3A5C")
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = borde_thin()

def generar_excel(df, titulo="Datos"):
    output = io.BytesIO()
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = titulo
    for cn, col in enumerate(df.columns, 1):
        estilo_cabecera(ws.cell(row=1, column=cn, value=col))
    ws.row_dimensions[1].height = 24
    for rn, (_, row) in enumerate(df.iterrows(), 2):
        estado = str(row.get("Estado",""))
        if "✅" in estado: bg = "C6EFCE"
        elif "❌" in estado: bg = "FFC7CE"
        elif "📋" in estado: bg = "EBF5FB"
        elif estado: bg = "FFEB9C"
        else: bg = "FFFFFF"
        fill = PatternFill("solid", fgColor=bg)
        for cn, val in enumerate(row, 1):
            c = ws.cell(row=rn, column=cn, value=str(val) if val is not None else "")
            c.font = Font(name="Calibri", size=9)
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            c.border = borde_thin()
            if "Estado" in df.columns and cn == df.columns.get_loc("Estado")+1:
                c.fill = fill; c.font = Font(name="Calibri", size=9, bold=True)
        ws.row_dimensions[rn].height = 15
    anchos = {"Archivo":18,"Nº Hoja":10,"Fecha":12,"Cliente":28,"Descripción":35,
              "Bus/Grupo":8,"Plazas":8,"Hora Salida":12,"Punto Recogida":28,
              "Hora Llegada":12,"Destino Regreso":28,"Notas":30,
              "Conductor":22,"Tipo":10,"Salida Conductor":14,"Llegada Conductor":14,"Estado":38}
    for cn, col in enumerate(df.columns, 1):
        ws.column_dimensions[get_column_letter(cn)].width = anchos.get(col, 15)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(df.columns))}1"
    wb.save(output); output.seek(0); return output.getvalue()

# ── INTERFAZ ─────────────────────────────────────────────────
def main():
    st.markdown(CSS, unsafe_allow_html=True)
    st.markdown("<h2 style='color:#1a3a5c;margin-bottom:0;'>🚌 Hojas de Ruta — Autocares Alegre</h2>"
                "<p style='color:#666;margin-top:0;'>Extrae datos de los PDFs de delsol y cruza con registros de conductores</p>",
                unsafe_allow_html=True)
    st.divider()

    # ── Paso 1 ───────────────────────────────────────────────
    st.markdown("### Paso 1 — Subir PDFs de delsol")
    st.caption("Puedes seleccionar varios PDFs a la vez con Ctrl+clic en el explorador de archivos.")

    if "uploader_key" not in st.session_state:
        st.session_state["uploader_key"] = 0

    col_up1, col_up2 = st.columns([4,1])
    with col_up1:
        archivos = st.file_uploader("Selecciona uno o varios PDFs", type=["pdf"],
                                    accept_multiple_files=True,
                                    key=f"uploader_{st.session_state['uploader_key']}")
    with col_up2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🗑️ Limpiar ficheros", use_container_width=True):
            st.session_state["uploader_key"] += 1
            for k in ["df_hojas","df_verif"]: st.session_state.pop(k, None)
            st.rerun()

    if not archivos:
        st.info("Sube los PDFs del programa delsol para empezar.")
        col1,col2,col3 = st.columns(3)
        col1.markdown("**1. Sube los PDFs**\nHojas de ruta directamente de delsol")
        col2.markdown("**2. Revisa los datos**\nEdita si algo no cuadra")
        col3.markdown("**3. Verifica**\nCruza con los registros de conductores")
        return

    st.markdown(f"**{len(archivos)} PDF(s) cargados:** {', '.join(f.name for f in archivos[:5])}{'...' if len(archivos)>5 else ''}")

    col_b1, col_b2 = st.columns([2,2])
    with col_b1:
        if st.button(f"📄 Extraer datos de {len(archivos)} PDF(s)", type="primary", use_container_width=True):
            hojas = []
            barra = st.progress(0)
            for i, archivo in enumerate(archivos):
                barra.progress((i+1)/len(archivos), text=f"Leyendo {archivo.name}...")
                hojas.append(extraer_pdf(archivo.read(), archivo.name))
            barra.empty()
            st.session_state["df_hojas"] = hojas_a_filas(hojas)
            st.session_state.pop("df_verif", None)
            st.rerun()
    with col_b2:
        if st.button("↺ Actualizar datos conductores", type="secondary", use_container_width=True):
            cargar_registros.clear(); st.success("Actualizado.")

    # ── Tabla extraída ───────────────────────────────────────
    if "df_hojas" not in st.session_state: return
    df_hojas = st.session_state["df_hojas"]
    st.divider()
    errores = df_hojas[df_hojas["Nº Hoja"]=="ERROR"]
    if not errores.empty:
        st.warning(f"⚠️ {len(errores)} PDF(s) con error de lectura: {', '.join(errores['Archivo'].tolist())}")
    st.success(f"✅ {len(df_hojas)} servicio(s) extraídos")

    df_ed = st.data_editor(df_hojas, use_container_width=True, hide_index=True,
        column_config={
            "Archivo":          st.column_config.TextColumn(disabled=True, width="small"),
            "Nº Hoja":          st.column_config.TextColumn(disabled=True, width="small"),
            "Fecha":            st.column_config.TextColumn(disabled=True, width="small"),
            "Cliente":          st.column_config.TextColumn(disabled=True, width="medium"),
            "Descripción":      st.column_config.TextColumn(disabled=True, width="large"),
            "Bus/Grupo":        st.column_config.TextColumn(disabled=True, width="small"),
            "Plazas":           st.column_config.TextColumn(disabled=True, width="small"),
            "Hora Salida":      st.column_config.TextColumn(width="small"),
            "Punto Recogida":   st.column_config.TextColumn(disabled=True, width="medium"),
            "Hora Llegada":     st.column_config.TextColumn(width="small"),
            "Destino Regreso":  st.column_config.TextColumn(disabled=True, width="medium"),
            "Notas":            st.column_config.TextColumn(width="medium"),
        }, key="tabla_hojas")
    st.session_state["df_hojas"] = df_ed

    # ── Filtros Paso 1 ───────────────────────────────────────
    st.markdown("**Filtros para exportar:**")
    fcol1, fcol2 = st.columns(2)
    with fcol1:
        fechas_disp = sorted(df_ed["Fecha"].dropna().unique().tolist())
        fechas_sel = st.multiselect("📅 Filtrar por fecha:", options=fechas_disp,
                                    placeholder="Todas las fechas", key="f1_fechas")
    with fcol2:
        clientes_disp = sorted(df_ed["Cliente"].dropna().unique().tolist())
        clientes_sel = st.multiselect("🏢 Filtrar por cliente:", options=clientes_disp,
                                      placeholder="Todos los clientes", key="f1_clientes")
    df_export1 = df_ed.copy()
    if fechas_sel:   df_export1 = df_export1[df_export1["Fecha"].isin(fechas_sel)]
    if clientes_sel: df_export1 = df_export1[df_export1["Cliente"].isin(clientes_sel)]
    if fechas_sel or clientes_sel:
        st.caption(f"Mostrando {len(df_export1)} de {len(df_ed)} servicios tras filtrar.")

    col1, col2, col3 = st.columns([2,2,1])
    with col1:
        st.download_button("⬇️ Excel con datos extraídos",
            data=generar_excel(df_export1, "Hojas de Ruta"),
            file_name=f"hojas_ruta_{date.today().strftime('%Y_%m_%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", use_container_width=True)
    with col3:
        if st.button("🗑️ Limpiar", use_container_width=True):
            for k in ["df_hojas","df_verif"]: st.session_state.pop(k, None)
            st.rerun()

    st.divider()

    # ── Paso 2 ───────────────────────────────────────────────
    st.markdown("### Paso 2 — Verificar con registros de conductores")
    st.caption("Cruza los datos del PDF con lo que registraron los conductores en la app.")
    if st.button("🔍 Verificar con conductores", type="primary", use_container_width=True):
        with st.spinner("Cruzando datos..."):
            df_cond = cargar_registros()
            st.session_state["df_verif"] = verificar(df_ed, df_cond)
        st.rerun()

    if "df_verif" not in st.session_state: return
    df_verif = st.session_state["df_verif"]

    n_ok   = df_verif["Estado"].str.contains("✅").sum()
    n_jorn = df_verif["Estado"].str.contains("📋").sum()
    n_warn = df_verif["Estado"].str.contains("⚠️").sum()
    n_err  = df_verif["Estado"].str.contains("❌").sum()
    m1,m2,m3,m4,m5 = st.columns(5)
    m1.metric("Total",         len(df_verif))
    m2.metric("✅ Coincide",   int(n_ok))
    m3.metric("📋 Jornada",    int(n_jorn))
    m4.metric("⚠️ Revisar",    int(n_warn))
    m5.metric("❌ Diferencia", int(n_err))

    # ── Filtros Paso 2 ───────────────────────────────────────
    st.markdown("**Filtros:**")
    fc1, fc2, fc3 = st.columns(3)
    with fc1:
        conds_disp = sorted([c for c in df_verif["Conductor"].dropna().unique() if c])
        conds_sel = st.multiselect("👤 Conductor:", options=conds_disp,
                                   placeholder="Todos", key="f2_cond")
    with fc2:
        fechas2_disp = sorted(df_verif["Fecha"].dropna().unique().tolist())
        fechas2_sel = st.multiselect("📅 Fecha:", options=fechas2_disp,
                                     placeholder="Todas", key="f2_fecha")
    with fc3:
        filtro = st.radio("Estado:", ["Todos","⚠️ y ❌","Solo ❌"], horizontal=False, key="f2_estado")

    df_v = df_verif.copy()
    if conds_sel:   df_v = df_v[df_v["Conductor"].isin(conds_sel)]
    if fechas2_sel: df_v = df_v[df_v["Fecha"].isin(fechas2_sel)]
    if   filtro == "⚠️ y ❌": df_v = df_v[df_v["Estado"].str.contains("⚠️|❌")]
    elif filtro == "Solo ❌":  df_v = df_v[df_v["Estado"].str.contains("❌")]

    if len(df_v) < len(df_verif):
        st.caption(f"Mostrando {len(df_v)} de {len(df_verif)} registros tras filtrar.")
    st.dataframe(df_v, use_container_width=True, hide_index=True)

    col_d1, col_d2 = st.columns(2)
    with col_d1:
        st.download_button("⬇️ Excel verificación (filtrado)",
            data=generar_excel(df_v, "Verificación"),
            file_name=f"verificacion_{date.today().strftime('%Y_%m_%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", use_container_width=True)
    with col_d2:
        st.download_button("⬇️ Excel verificación (completo)",
            data=generar_excel(df_verif, "Verificación"),
            file_name=f"verificacion_completo_{date.today().strftime('%Y_%m_%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="secondary", use_container_width=True)

if __name__ == "__main__":
    main()
