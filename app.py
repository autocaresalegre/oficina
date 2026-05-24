# ============================================================
#  AUTOCARES ALEGRE v8.0
#  - Excel 2 filas por día (fijos + horas/obs)
#  - Admin panel limpio sin JSON
#  - Importes visibles antes de descargar
#  - PDF de resumen
#  - Formulario manual en admin
# ============================================================

import streamlit as st
import pandas as pd
import json
from supabase import create_client
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from fpdf import FPDF
from datetime import date, datetime, time
import io

st.set_page_config(page_title="Autocares Alegre", page_icon="🚌",
                   layout="centered", initial_sidebar_state="collapsed")

LOGO_B64 = "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"
PASSWORD_ADMIN = "alegre2026"
SERVICIOS_GRID = ["Pilotos","Pescadores","Logista-Saludes","Colegio","Transfer",
                  "Boda Ida","Boda Regreso 1","Boda Regreso 2"]
PRECIOS_FIJOS  = {"Pilotos":15,"Pescadores":15,"Logista-Saludes":15,"Colegio":25,
                  "Transfer":20,"Boda Ida":40,"Boda Regreso 1":40,"Boda Regreso 2":40}
PRECIO_DIETA=15; PRECIO_HORA_SEMANA=10; PRECIO_HORA_FINDE=12
FESTIVOS = {
    date(2025,1,1),date(2025,1,6),date(2025,3,19),date(2025,4,17),date(2025,4,18),
    date(2025,4,28),date(2025,5,1),date(2025,8,15),date(2025,10,9),date(2025,10,12),
    date(2025,11,1),date(2025,12,6),date(2025,12,8),date(2025,12,25),
    date(2026,1,1),date(2026,1,6),date(2026,3,19),date(2026,4,2),date(2026,4,3),
    date(2026,4,20),date(2026,5,1),date(2026,8,15),date(2026,10,9),date(2026,10,12),
    date(2026,11,1),date(2026,12,6),date(2026,12,8),date(2026,12,25),
}
COLUMNAS_BD=["conductor","fecha","dieta","servicios_fijos","servicios_horas","observaciones"]

# ── CSS ─────────────────────────────────────────────────────
CSS="""
<style>
header{visibility:hidden;}#MainMenu{visibility:hidden;}footer{visibility:hidden;}
[data-testid="stToolbar"]{display:none;}[data-testid="stDecoration"]{display:none;}
[data-testid="stSidebar"]{display:none !important;}
.block-container{padding-top:0.5rem !important;padding-bottom:0.5rem !important;max-width:480px !important;}
.stButton>button{height:2.8rem;font-size:0.95rem;font-weight:700;border-radius:10px;width:100%;}
.stSelectbox>label,.stDateInput>label,.stTextInput>label,.stTimeInput>label,
.stNumberInput>label,.stCheckbox>label span,.stTextArea>label{font-size:0.92rem !important;font-weight:600 !important;}
.admin-link{text-align:center;font-size:0.7rem;margin-top:1.5rem;opacity:0.3;}
</style>"""
CSS_ADMIN="""
<style>
[data-testid="stSidebar"]{display:block !important;}
.block-container{max-width:1200px !important;padding-top:1rem !important;}
</style>"""

def logo_html(w="170px"):
    return f'''<div style="text-align:center;margin-bottom:0.3rem;">
    <img src="data:image/jpeg;base64,{LOGO_B64}" style="max-width:{w};width:55%;border-radius:6px;">
    </div>'''

# ── HELPERS JSON ────────────────────────────────────────────
def fijos_a_texto(raw):
    try:
        items = json.loads(str(raw or "[]"))
        partes = []
        for i in items:
            tipo=i.get("tipo",""); cant=int(i.get("cantidad",1))
            if tipo: partes.append(f"{tipo} x{cant}" if cant>1 else tipo)
        return ", ".join(partes) if partes else ""
    except: return str(raw) if raw else ""

def horas_a_texto(raw):
    try:
        items = json.loads(str(raw or "[]"))
        partes = []
        for i in items:
            c=i.get("concepto",""); ini=i.get("inicio",""); fin=i.get("fin","")
            if c: partes.append(f"{c} {ini} a {fin}" if ini and fin else c)
        return ", ".join(partes) if partes else ""
    except: return str(raw) if raw else ""

def necesita_revision(row):
    sf = fijos_a_texto(row.get("servicios_fijos",""))
    sh = horas_a_texto(row.get("servicios_horas",""))
    if not sf and not sh: return "Sin servicios registrados"
    try: json.loads(str(row.get("servicios_fijos") or "[]"))
    except: return "JSON servicios fijos incorrecto"
    try: json.loads(str(row.get("servicios_horas") or "[]"))
    except: return "JSON servicios horas incorrecto"
    return ""

# ── SUPABASE ────────────────────────────────────────────────
@st.cache_resource
def get_supabase():
    return create_client(st.secrets["supabase"]["url"],st.secrets["supabase"]["key"])

@st.cache_data(ttl=60)
def cargar_conductores():
    try:
        r=get_supabase().table("conductores").select("nombre").order("nombre").execute()
        return [f["nombre"] for f in r.data if f.get("nombre")]
    except Exception as e: st.error(f"Error: {e}"); return []

@st.cache_data(ttl=20)
def cargar_datos_brutos():
    try:
        r=get_supabase().table("datos_brutos").select(",".join(COLUMNAS_BD)).order("id").execute()
        if not r.data: return pd.DataFrame(columns=COLUMNAS_BD)
        df=pd.DataFrame(r.data)
        for col in COLUMNAS_BD:
            if col not in df.columns: df[col]=None
        return df[COLUMNAS_BD]
    except Exception as e: st.error(f"Error: {e}"); return pd.DataFrame(columns=COLUMNAS_BD)

def guardar_fila(d):
    try: get_supabase().table("datos_brutos").insert(d).execute(); cargar_datos_brutos.clear(); return True
    except Exception as e: st.error(f"Error: {e}"); return False

def guardar_tabla_completa(df):
    try:
        sb=get_supabase()
        ids=sb.table("datos_brutos").select("id").execute()
        if ids.data:
            all_ids=[f["id"] for f in ids.data]
            for i in range(0,len(all_ids),500):
                sb.table("datos_brutos").delete().in_("id",all_ids[i:i+500]).execute()
        if not df.empty:
            filas=[]
            for _,row in df.iterrows():
                f={col:(str(row[col]) if row[col] is not None else None) for col in COLUMNAS_BD if col in df.columns}
                if "dieta" in f: f["dieta"]=str(f["dieta"]).lower() in ["true","1","si"]
                filas.append(f)
            for i in range(0,len(filas),500):
                sb.table("datos_brutos").insert(filas[i:i+500]).execute()
        cargar_datos_brutos.clear(); return True
    except Exception as e: st.error(f"Error: {e}"); return False

# ── CÁLCULO ──────────────────────────────────────────────────
def es_dia_especial(fecha):
    if isinstance(fecha,datetime): fecha=fecha.date()
    elif isinstance(fecha,str):
        try: fecha=pd.to_datetime(fecha).date()
        except: return False
    return fecha.weekday()>=5 or fecha in FESTIVOS

def calcular_horas(h1,h2):
    try:
        if isinstance(h1,str): h1=datetime.strptime(h1.strip(),"%H:%M").time()
        if isinstance(h2,str): h2=datetime.strptime(h2.strip(),"%H:%M").time()
        i=datetime.combine(date.today(),h1); f=datetime.combine(date.today(),h2)
        return round((f-i).total_seconds()/3600,4) if f>i else 0.0
    except: return 0.0

def calcular_liquidacion(df):
    if df.empty: return pd.DataFrame()
    df=df.copy()
    df["fecha"]=pd.to_datetime(df["fecha"],errors="coerce").dt.date
    def a_bool(v):
        if isinstance(v,bool): return v
        return str(v).strip().lower() in ["true","1","si"]
    df["dieta"]=df["dieta"].fillna(False).apply(a_bool)
    resultados=[]
    for _,fila in df.iterrows():
        conductor=fila.get("conductor",""); fecha=fila.get("fecha")
        if pd.isna(fecha): continue
        tarifa=PRECIO_HORA_FINDE if es_dia_especial(fecha) else PRECIO_HORA_SEMANA
        imp_dieta=PRECIO_DIETA if fila.get("dieta") else 0.0
        fijos_txt=[]; horas_txt=[]; horas_tot=0.0; imp_horas=0.0; imp_fijos=0.0
        sf_raw=fila.get("servicios_fijos")
        if sf_raw and str(sf_raw) not in ("","nan","None","[]"):
            try:
                for item in json.loads(str(sf_raw)):
                    tipo=item.get("tipo",""); cant=int(item.get("cantidad",1))
                    if tipo in PRECIOS_FIJOS:
                        imp_fijos+=PRECIOS_FIJOS[tipo]*cant
                        fijos_txt.append(f"{tipo}" if cant==1 else f"{tipo} x{cant}")
            except: pass
        sh_raw=fila.get("servicios_horas")
        if sh_raw and str(sh_raw) not in ("","nan","None","[]"):
            try:
                for item in json.loads(str(sh_raw)):
                    c=item.get("concepto",""); h1=item.get("inicio",""); h2=item.get("fin","")
                    if h1 and h2:
                        h=calcular_horas(h1,h2); horas_tot+=h; imp_horas+=h*tarifa
                    if c: horas_txt.append(f"{c} {h1} a {h2}" if h1 and h2 else c)
            except: pass
        obs=str(fila.get("observaciones","")) if fila.get("observaciones") else ""
        revision=necesita_revision(fila)
        resultados.append({
            "Conductor":conductor,"Fecha":fecha,
            "Servicios Fijos":(" | ".join(fijos_txt)) if fijos_txt else "—",
            "Por Horas":(" | ".join(horas_txt)) if horas_txt else "—",
            "Observaciones":obs,
            "Horas Tot.":round(horas_tot,2),
            "IMP. HORAS":round(imp_horas,2),"IMP. FIJOS":round(imp_fijos,2),
            "IMP. DIETAS":round(imp_dieta,2),
            "TOTAL":round(imp_horas+imp_fijos+imp_dieta,2),
            "REVISAR":"⚠️ "+revision if revision else "",
        })
    return pd.DataFrame(resultados).sort_values(["Conductor","Fecha"]).reset_index(drop=True)

# ── EXCEL 2 FILAS POR DÍA ───────────────────────────────────
def generar_excel_bytes(df_cierre):
    output=io.BytesIO()
    wb=openpyxl.Workbook(); ws=wb.active; ws.title="Liquidacion"
    C_AZUL="1F4E79"; C_AMAR="FFF2CC"; C_VERDE="C6EFCE"
    C_FIJO="EBF5FB"; C_HORA="FEF9E7"; C_SUB="D5F5E3"; C_REV="FADBD8"
    borde=Border(left=Side(style="thin",color="CCCCCC"),right=Side(style="thin",color="CCCCCC"),
                 top=Side(style="thin",color="CCCCCC"),bottom=Side(style="thin",color="CCCCCC"))
    al_c=Alignment(horizontal="center",vertical="center",wrap_text=True)
    al_d=Alignment(horizontal="right",vertical="center")
    al_i=Alignment(horizontal="left",vertical="center",wrap_text=True)
    # Cabeceras
    cabs=["Conductor","Fecha","Servicios / Horario","Obs./Revisar",
          "Horas","IMP. HORAS","IMP. FIJOS","IMP. DIETAS","TOTAL DÍA"]
    for cn,t in enumerate(cabs,1):
        c=ws.cell(row=1,column=cn,value=t)
        c.font=Font(bold=True,color="FFFFFF",size=10,name="Calibri")
        c.fill=PatternFill("solid",fgColor=C_AZUL); c.alignment=al_c; c.border=borde
    ws.row_dimensions[1].height=28
    fila_xls=2
    conductor_actual=None
    sub_fijos=0; sub_horas=0; sub_dietas=0; sub_total=0
    sub_primera_fila=2
    def escribir_subtotal(fxls,nombre,s_fijos,s_horas,s_dietas,s_total):
        c=ws.cell(row=fxls,column=1,value=f"SUBTOTAL — {nombre}")
        c.font=Font(bold=True,size=10,color="FFFFFF",name="Calibri")
        c.fill=PatternFill("solid",fgColor="1A5276"); c.alignment=al_i; c.border=borde
        for cn in(2,3,4,5): 
            cx=ws.cell(row=fxls,column=cn)
            cx.fill=PatternFill("solid",fgColor="1A5276"); cx.border=borde
        for cn,val in zip((6,7,8,9),(s_horas,s_fijos,s_dietas,s_total)):
            cx=ws.cell(row=fxls,column=cn,value=round(val,2))
            cx.number_format='#,##0.00 "€"'; cx.font=Font(bold=True,size=10,color="FFFFFF",name="Calibri")
            cx.fill=PatternFill("solid",fgColor="1A5276"); cx.alignment=al_d; cx.border=borde
        ws.row_dimensions[fxls].height=20
    for _,fila in df_cierre.iterrows():
        conductor=fila["Conductor"]; fecha=fila["Fecha"]
        fondo=es_dia_especial(fecha)
        fill_f=PatternFill("solid",fgColor=C_AMAR if fondo else C_FIJO)
        fill_h=PatternFill("solid",fgColor=C_AMAR if fondo else C_HORA)
        revisar=fila.get("REVISAR","")
        fill_rev=PatternFill("solid",fgColor=C_REV) if revisar else None
        # Subtotal del conductor anterior
        if conductor_actual is not None and conductor!=conductor_actual:
            escribir_subtotal(fila_xls,conductor_actual,sub_fijos,sub_horas,sub_dietas,sub_total)
            fila_xls+=1
            sub_fijos=0; sub_horas=0; sub_dietas=0; sub_total=0
        conductor_actual=conductor
        sub_fijos+=fila["IMP. FIJOS"]; sub_horas+=fila["IMP. HORAS"]
        sub_dietas+=fila["IMP. DIETAS"]; sub_total+=fila["TOTAL"]
        # ── FILA 1: servicios fijos ──
        vals1=[conductor,fecha,fila["Servicios Fijos"],
               revisar if revisar else ("🍽️ Dieta" if fila["IMP. DIETAS"]>0 else ""),
               "",None,fila["IMP. FIJOS"],fila["IMP. DIETAS"],fila["TOTAL"]]
        for cn,val in enumerate(vals1,1):
            c=ws.cell(row=fila_xls,column=cn,value=val); c.border=borde
            c.fill=fill_rev if (revisar and cn in(3,4)) else fill_f
            if cn==1: c.font=Font(bold=True,size=10,name="Calibri"); c.alignment=al_i
            elif cn==2: c.number_format="DD/MM/YYYY"; c.alignment=al_c; c.font=Font(size=10,name="Calibri")
            elif cn in(3,4): c.alignment=al_i; c.font=Font(size=revisar and cn==4 and 9 or 10,name="Calibri",
                                                              bold=bool(revisar and cn==4),
                                                              color="C0392B" if (revisar and cn==4) else "000000")
            elif cn==5: c.alignment=al_c
            elif cn in(6,7,8):
                c.number_format='#,##0.00 "€"'; c.alignment=al_d; c.font=Font(size=10,name="Calibri")
            elif cn==9:
                c.number_format='#,##0.00 "€"'; c.alignment=al_d
                c.font=Font(bold=True,size=10,name="Calibri")
                if not (revisar or fondo): c.fill=PatternFill("solid",fgColor=C_VERDE)
        ws.row_dimensions[fila_xls].height=18; fila_xls+=1
        # ── FILA 2: horario + observaciones (solo si hay algo) ──
        por_horas=fila["Por Horas"]; obs=fila["Observaciones"]
        horas_tot=fila["Horas Tot."]; imp_horas=fila["IMP. HORAS"]
        if por_horas!="—" or obs:
            contenido_h=por_horas if por_horas!="—" else ""
            vals2=["","",contenido_h,obs,
                   horas_tot if horas_tot>0 else "",
                   imp_horas if imp_horas>0 else None,"","",""]
            for cn,val in enumerate(vals2,1):
                c=ws.cell(row=fila_xls,column=cn,value=val); c.border=borde
                c.fill=fill_h
                if cn in(3,4): c.alignment=al_i; c.font=Font(size=9,name="Calibri",italic=True)
                elif cn==5: c.alignment=al_c; c.font=Font(size=9,name="Calibri")
                elif cn==6:
                    c.number_format='#,##0.00 "€"'; c.alignment=al_d
                    c.font=Font(size=9,name="Calibri",italic=True)
            ws.row_dimensions[fila_xls].height=16; fila_xls+=1
    # Último subtotal
    if conductor_actual:
        escribir_subtotal(fila_xls,conductor_actual,sub_fijos,sub_horas,sub_dietas,sub_total)
        fila_xls+=1
    anchos={1:24,2:12,3:42,4:28,5:8,6:14,7:14,8:13,9:14}
    for cn,ancho in anchos.items():
        ws.column_dimensions[get_column_letter(cn)].width=ancho
    ws.freeze_panes="A2"; ws.auto_filter.ref=f"A1:I1"
    wb.save(output); output.seek(0); return output.getvalue()

# ── PDF ──────────────────────────────────────────────────────
def generar_pdf_bytes(df_cierre):
    pdf=FPDF(orientation="P",unit="mm",format="A4")
    pdf.set_auto_page_break(auto=True,margin=15)
    pdf.add_page()
    # Cabecera
    pdf.set_font("Helvetica","B",16)
    pdf.set_text_color(31,78,121)
    pdf.cell(0,10,"AUTOCARES ALEGRE",new_x="LMARGIN",new_y="NEXT",align="C")
    pdf.set_font("Helvetica","",11)
    pdf.set_text_color(80,80,80)
    pdf.cell(0,6,f"Resumen de Liquidacion  -  Generado: {date.today().strftime('%d/%m/%Y')}",
             new_x="LMARGIN",new_y="NEXT",align="C")
    pdf.ln(4)
    # Resumen global
    total_g=df_cierre["TOTAL"].sum()
    pdf.set_font("Helvetica","B",10)
    pdf.set_fill_color(235,245,251)
    pdf.set_text_color(0,0,0)
    pdf.cell(0,7,f"  TOTAL GENERAL A PAGAR: {total_g:,.2f} EUR   |   {df_cierre['Conductor'].nunique()} conductores   |   {len(df_cierre)} dias",
             fill=True,new_x="LMARGIN",new_y="NEXT")
    pdf.ln(5)
    # Por conductor
    for conductor, grupo in df_cierre.groupby("Conductor"):
        pdf.set_font("Helvetica","B",11)
        pdf.set_fill_color(31,78,121)
        pdf.set_text_color(255,255,255)
        pdf.cell(0,8,f"  {conductor}",fill=True,new_x="LMARGIN",new_y="NEXT")
        pdf.set_font("Helvetica","",9)
        pdf.set_text_color(0,0,0)
        # Cabecera tabla
        pdf.set_fill_color(215,230,242)
        pdf.set_font("Helvetica","B",8)
        pdf.cell(22,6,"Fecha",fill=True,border=1)
        pdf.cell(65,6,"Servicios Fijos",fill=True,border=1)
        pdf.cell(55,6,"Por Horas",fill=True,border=1)
        pdf.cell(16,6,"Horas",fill=True,border=1,align="C")
        pdf.cell(22,6,"Total dia",fill=True,border=1,align="R",new_x="LMARGIN",new_y="NEXT")
        pdf.set_font("Helvetica","",8)
        for _,row in grupo.iterrows():
            fecha_str=row["Fecha"].strftime("%d/%m/%Y") if hasattr(row["Fecha"],"strftime") else str(row["Fecha"])
            fijos=row["Servicios Fijos"][:60] if row["Servicios Fijos"]!="—" else ""
            horas=row["Por Horas"][:50] if row["Por Horas"]!="—" else ""
            revisar=row.get("REVISAR","")
            bg=(252,215,215) if revisar else (255,255,255)
            pdf.set_fill_color(*bg)
            pdf.cell(22,5,fecha_str,fill=True,border=1)
            pdf.cell(65,5,fijos,fill=True,border=1)
            pdf.cell(55,5,horas,fill=True,border=1)
            pdf.cell(16,5,str(row["Horas Tot."]) if row["Horas Tot."]>0 else "",fill=True,border=1,align="C")
            pdf.cell(22,5,f"{row['TOTAL']:,.2f} EUR",fill=True,border=1,align="R",new_x="LMARGIN",new_y="NEXT")
            if revisar:
                pdf.set_font("Helvetica","I",7)
                pdf.set_text_color(192,57,43)
                pdf.cell(0,4,f"  {revisar}",new_x="LMARGIN",new_y="NEXT")
                pdf.set_font("Helvetica","",8); pdf.set_text_color(0,0,0)
            if row["Observaciones"]:
                pdf.set_font("Helvetica","I",7)
                pdf.set_text_color(100,100,100)
                pdf.cell(0,4,f"  Obs: {row['Observaciones'][:80]}",new_x="LMARGIN",new_y="NEXT")
                pdf.set_font("Helvetica","",8); pdf.set_text_color(0,0,0)
        # Subtotal conductor
        pdf.set_font("Helvetica","B",9)
        pdf.set_fill_color(196,230,197)
        sub=grupo["TOTAL"].sum()
        h_sub=grupo["IMP. HORAS"].sum(); f_sub=grupo["IMP. FIJOS"].sum(); d_sub=grupo["IMP. DIETAS"].sum()
        pdf.cell(0,6,
            f"  Subtotal {conductor}:  Fijos {f_sub:,.2f} EUR  +  Horas {h_sub:,.2f} EUR  +  Dietas {d_sub:,.2f} EUR  =  {sub:,.2f} EUR",
            fill=True,new_x="LMARGIN",new_y="NEXT")
        pdf.ln(4)
    buf=io.BytesIO()
    pdf.output(buf)
    buf.seek(0); return buf.getvalue()

# ── VISTA CONDUCTOR ──────────────────────────────────────────
def vista_conductor():
    st.markdown(CSS,unsafe_allow_html=True)
    if st.session_state.get("envio_ok"):
        info=st.session_state.get("envio_info",{})
        st.markdown(logo_html(),unsafe_allow_html=True); st.divider()
        st.success(f"✅ **Registro enviado**\n\n👤 **{info.get('nombre','')}** — 📅 {info.get('fecha','')}")
        if st.button("➕ Registrar otro servicio",type="primary",use_container_width=True):
            for k in ["envio_ok","resumen","datos_resumen","n_horas"]: st.session_state.pop(k,None)
            st.rerun()
        st.markdown('<p class="admin-link">· · ·</p>',unsafe_allow_html=True)
        return
    if st.session_state.get("resumen"):
        datos=st.session_state.get("datos_resumen",{})
        st.markdown(logo_html(),unsafe_allow_html=True)
        st.markdown("<h4 style='text-align:center;color:#1F4E79;'>📋 Revisa tu registro</h4>",unsafe_allow_html=True)
        st.divider()
        st.markdown(f"👤 **{datos.get('nombre','')}**  —  📅 {datos.get('fecha_str','')}")
        st.markdown(f"🍽️ Dieta: {'✅ Sí' if datos.get('dieta') else '❌ No'}")
        if datos.get("fijos_ok"):
            st.markdown("**📌 Fijos:**  " + ",  ".join(
                f"{f['tipo']} x{f['cantidad']}" if f['cantidad']>1 else f['tipo']
                for f in datos["fijos_ok"]))
        if datos.get("horas_ok"):
            st.markdown("**🕐 Horas:**  " + ",  ".join(
                f"{h['concepto']} ({h['inicio']}→{h['fin']})" for h in datos["horas_ok"]))
        if datos.get("observaciones"):
            st.markdown(f"**💬** {datos['observaciones']}")
        st.divider()
        c1,c2=st.columns(2)
        with c1:
            if st.button("✏️ Corregir",type="secondary",use_container_width=True):
                st.session_state.resumen=False; st.rerun()
        with c2:
            if st.button("✅ Confirmar y Enviar",type="primary",use_container_width=True):
                fila={"conductor":datos["nombre"],"fecha":datos["fecha_iso"],"dieta":datos["dieta"],
                      "servicios_fijos":json.dumps(datos["fijos_ok"],ensure_ascii=False),
                      "servicios_horas":json.dumps(datos["horas_ok"],ensure_ascii=False),
                      "observaciones":datos.get("observaciones") or None}
                with st.spinner("Enviando..."): ok=guardar_fila(fila)
                if ok:
                    st.session_state.envio_ok=True
                    st.session_state.envio_info={"nombre":datos["nombre"],"fecha":datos["fecha_str"]}
                    st.session_state.resumen=False; st.session_state.pop("n_horas",None); st.rerun()
        return
    # FORMULARIO
    st.markdown(logo_html(),unsafe_allow_html=True)
    st.markdown("<h4 style='text-align:center;color:#1F4E79;margin-bottom:0;'>Registro de Servicios Extras</h4>",unsafe_allow_html=True)
    st.divider()
    conductores=cargar_conductores()
    busqueda=st.text_input("🔍 Buscar conductor",placeholder="Escribe las primeras letras...")
    cond_f=([c for c in conductores if busqueda.strip().lower() in c.lower()] if busqueda.strip() else conductores)
    nombre_sel=st.selectbox("👤 Selecciona tu Nombre",options=["— Selecciona —"]+cond_f)
    fecha_sel=st.date_input("📅 Fecha del Servicio",value=date.today(),format="DD/MM/YYYY")
    st.markdown("---")
    st.markdown("**📌 Servicios Fijos** — ¿Cuántos de cada uno?")
    st.caption("Pon 0 si no has hecho ese servicio.")
    c1,c2=st.columns(2); conteo_fijos={}
    for idx,s in enumerate(SERVICIOS_GRID):
        col=c1 if idx%2==0 else c2
        conteo_fijos[s]=col.number_input(s,min_value=0,max_value=5,value=0,key=f"fijo_{idx}",step=1)
    st.markdown("---")
    st.markdown("**🕐 Servicios Por Horas**")
    if "n_horas" not in st.session_state: st.session_state.n_horas=1
    horas_cap=[]
    for i in range(st.session_state.n_horas):
        if i>0: st.markdown("---")
        conc=st.text_input(f"Concepto / Destino #{i+1}",placeholder="Excursión o salida / Viaje a...",key=f"h_c_{i}")
        co1,co2=st.columns(2)
        with co1: hi=st.time_input("🕐 Inicio",value=time(8,0),step=300,key=f"h_i_{i}")
        with co2: hf=st.time_input("🕔 Fin",value=time(16,0),step=300,key=f"h_f_{i}")
        horas_cap.append({"concepto":conc,"hora_ini":hi,"hora_fin":hf})
    ca,cb=st.columns(2)
    with ca:
        if st.button("➕ Añadir horas",type="secondary",use_container_width=True):
            st.session_state.n_horas+=1; st.rerun()
    with cb:
        if st.session_state.n_horas>1:
            if st.button("➖ Quitar",type="secondary",use_container_width=True):
                st.session_state.n_horas-=1; st.rerun()
    st.markdown("---")
    dieta_sel=st.checkbox("🍽️ ¿Te corresponde Dieta hoy?",value=False)
    obs=st.text_area("💬 Observaciones (opcional)",placeholder="Incidencias, dudas o comentarios...",height=65)
    st.divider()
    if st.button("👁️ VER RESUMEN Y CONFIRMAR",type="primary",use_container_width=True):
        errores=[]
        if nombre_sel=="— Selecciona —": errores.append("❌ Debes seleccionar tu nombre.")
        fijos_ok=[{"tipo":s,"cantidad":int(conteo_fijos[s])} for s in SERVICIOS_GRID if conteo_fijos.get(s,0)>0]
        horas_ok=[]
        for i,h in enumerate(horas_cap):
            if h["concepto"].strip():
                if h["hora_fin"]<=h["hora_ini"]: errores.append(f"❌ Horas #{i+1}: fin debe ser posterior al inicio.")
                else: horas_ok.append({"concepto":h["concepto"].strip(),"inicio":h["hora_ini"].strftime("%H:%M"),"fin":h["hora_fin"].strftime("%H:%M")})
        if not fijos_ok and not horas_ok and not errores: errores.append("❌ Añade al menos un servicio.")
        if errores:
            for msg in errores: st.error(msg)
        else:
            st.session_state.resumen=True
            st.session_state.datos_resumen={"nombre":nombre_sel,"fecha_str":fecha_sel.strftime("%d/%m/%Y"),
                "fecha_iso":fecha_sel.strftime("%Y-%m-%d"),"dieta":dieta_sel,
                "fijos_ok":fijos_ok,"horas_ok":horas_ok,"observaciones":obs.strip()}
            st.rerun()
    st.markdown('<p class="admin-link"><a href="?admin=1" style="color:#aaa;text-decoration:none;">· · ·</a></p>',unsafe_allow_html=True)

# ── VISTA ADMIN ──────────────────────────────────────────────
def vista_admin():
    st.markdown(CSS_ADMIN,unsafe_allow_html=True)
    if "admin_ok" not in st.session_state: st.session_state.admin_ok=False
    if not st.session_state.admin_ok:
        st.title("🔐 Administración — Autocares Alegre"); st.divider()
        pw=st.text_input("Contraseña",type="password")
        if st.button("▶ Entrar",type="primary",use_container_width=True):
            if pw==PASSWORD_ADMIN: st.session_state.admin_ok=True; st.rerun()
            else: st.error("❌ Contraseña incorrecta.")
        return
    c1,c2=st.columns([5,1])
    with c1: st.title("📊 Administración — Autocares Alegre")
    with c2:
        st.markdown("<br>",unsafe_allow_html=True)
        if st.button("🚪 Salir"): st.session_state.admin_ok=False; st.rerun()
    st.divider()

    # ══ 1. REGISTROS LEGIBLES ══════════════════════════════════
    st.markdown("## 👁️ Registros enviados por los conductores")
    c_ref,c_esp=st.columns([1,3])
    with c_ref:
        if st.button("🔄 Actualizar",type="secondary"):
            cargar_datos_brutos.clear(); st.rerun()
    df_brutos=cargar_datos_brutos()
    if df_brutos.empty:
        st.info("ℹ️ Todavía no hay registros.")
    else:
        filas_v=[]
        for _,row in df_brutos.iterrows():
            revision=necesita_revision(row)
            filas_v.append({
                "Conductor":row.get("conductor",""),
                "Fecha":row.get("fecha",""),
                "Dieta":"✓" if str(row.get("dieta","")).lower() in ["true","1"] else "",
                "Servicios Fijos":fijos_a_texto(row.get("servicios_fijos","")),
                "Por Horas":horas_a_texto(row.get("servicios_horas","")),
                "Observaciones":str(row.get("observaciones","")) if row.get("observaciones") else "",
                "⚠️":"⚠️ REVISAR" if revision else "",
            })
        st.dataframe(pd.DataFrame(filas_v),use_container_width=True,hide_index=True)
    st.divider()

    # ══ 2. IMPORTES CALCULADOS (SIEMPRE VISIBLES) ═════════════
    st.markdown("## 💰 Importes Calculados")
    st.caption("Se calculan automáticamente con los datos actuales. Revisa antes de descargar.")
    if not df_brutos.empty:
        df_cierre=calcular_liquidacion(df_brutos)
        if not df_cierre.empty:
            # Tabla limpia con columnas clave
            cols_mostrar=["Conductor","Fecha","Servicios Fijos","Por Horas",
                          "IMP. FIJOS","IMP. HORAS","IMP. DIETAS","TOTAL","REVISAR"]
            df_mostrar=df_cierre[[c for c in cols_mostrar if c in df_cierre.columns]].copy()
            # Colorear filas con REVISAR
            st.dataframe(df_mostrar,use_container_width=True,hide_index=True,
                column_config={
                    "IMP. FIJOS":st.column_config.NumberColumn(format="%.2f €"),
                    "IMP. HORAS":st.column_config.NumberColumn(format="%.2f €"),
                    "IMP. DIETAS":st.column_config.NumberColumn(format="%.2f €"),
                    "TOTAL":st.column_config.NumberColumn(format="%.2f €"),
                })
            # Totales rápidos
            m1,m2,m3,m4=st.columns(4)
            m1.metric("Fijos",f"{df_cierre['IMP. FIJOS'].sum():,.2f} €")
            m2.metric("Horas",f"{df_cierre['IMP. HORAS'].sum():,.2f} €")
            m3.metric("Dietas",f"{df_cierre['IMP. DIETAS'].sum():,.2f} €")
            m4.metric("TOTAL",f"{df_cierre['TOTAL'].sum():,.2f} €")
            st.markdown("---")
            nombre_arch=f"cierre_{date.today().strftime('%Y_%m_%d')}"
            c_xl,c_pdf=st.columns(2)
            with c_xl:
                st.download_button("⬇️ EXCEL",data=generar_excel_bytes(df_cierre),
                    file_name=f"{nombre_arch}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",use_container_width=True)
            with c_pdf:
                st.download_button("📄 PDF RESUMEN",data=generar_pdf_bytes(df_cierre),
                    file_name=f"{nombre_arch}.pdf",mime="application/pdf",
                    type="primary",use_container_width=True)
    st.divider()

    # ══ 3. AÑADIR REGISTRO MANUAL ═════════════════════════════
    with st.expander("➕ Añadir Registro Manual"):
        conductores=cargar_conductores()
        m_cond=st.selectbox("Conductor",["— Selecciona —"]+conductores,key="m_cond")
        m_fecha=st.date_input("Fecha",value=date.today(),format="DD/MM/YYYY",key="m_fecha")
        m_dieta=st.checkbox("🍽️ Dieta",key="m_dieta")
        st.markdown("**Servicios Fijos:**")
        mc1,mc2=st.columns(2); m_fijos={}
        for idx,s in enumerate(SERVICIOS_GRID):
            col=mc1 if idx%2==0 else mc2
            m_fijos[s]=col.number_input(s,min_value=0,max_value=5,value=0,key=f"m_fijo_{idx}",step=1)
        st.markdown("**Por Horas:**")
        m_conc=st.text_input("Concepto",placeholder="Destino / Servicio",key="m_conc")
        mh1,mh2=st.columns(2)
        with mh1: m_ini=st.time_input("Inicio",value=time(8,0),step=300,key="m_ini")
        with mh2: m_fin=st.time_input("Fin",value=time(16,0),step=300,key="m_fin")
        m_obs=st.text_area("Observaciones",height=60,key="m_obs")
        if st.button("💾 Guardar Registro Manual",type="primary",use_container_width=True):
            if m_cond=="— Selecciona —":
                st.error("❌ Selecciona un conductor.")
            else:
                m_fijos_ok=[{"tipo":s,"cantidad":int(m_fijos[s])} for s in SERVICIOS_GRID if m_fijos.get(s,0)>0]
                m_horas_ok=[]
                if m_conc.strip():
                    m_horas_ok=[{"concepto":m_conc.strip(),"inicio":m_ini.strftime("%H:%M"),"fin":m_fin.strftime("%H:%M")}]
                fila={"conductor":m_cond,"fecha":m_fecha.strftime("%Y-%m-%d"),"dieta":m_dieta,
                      "servicios_fijos":json.dumps(m_fijos_ok,ensure_ascii=False),
                      "servicios_horas":json.dumps(m_horas_ok,ensure_ascii=False),
                      "observaciones":m_obs.strip() or None}
                if guardar_fila(fila): st.success("✅ Registro guardado."); st.rerun()
    st.divider()

    # ══ 4. TABLA DE AUDITORÍA (edición avanzada) ═════════════
    with st.expander("✏️ Edición avanzada (Tabla de Auditoría)"):
        st.caption("Para correcciones directas. Los servicios se muestran como texto legible.")
        if not df_brutos.empty:
            df_show=df_brutos.copy()
            df_show["dieta"]=df_show["dieta"].fillna(False).apply(
                lambda x: bool(x) if isinstance(x,bool) else str(x).lower() in ["true","1"])
            df_show["servicios_fijos"]=df_show["servicios_fijos"].apply(fijos_a_texto)
            df_show["servicios_horas"]=df_show["servicios_horas"].apply(horas_a_texto)
            for col in ["conductor","fecha","observaciones"]:
                df_show[col]=df_show[col].fillna("").astype(str)
            df_show=df_show.rename(columns={
                "conductor":"Conductor","fecha":"Fecha","dieta":"Dieta",
                "servicios_fijos":"Servicios Fijos","servicios_horas":"Por Horas","observaciones":"Observaciones"})
            st.info("ℹ️ Aquí puedes borrar filas incorrectas. Para editar servicios usa el formulario de arriba.")
            df_edit=st.data_editor(df_show,use_container_width=True,num_rows="dynamic",hide_index=True,
                column_config={
                    "Conductor":st.column_config.TextColumn(width="medium"),
                    "Fecha":st.column_config.TextColumn(),
                    "Dieta":st.column_config.CheckboxColumn(default=False),
                    "Servicios Fijos":st.column_config.TextColumn(width="large",disabled=True),
                    "Por Horas":st.column_config.TextColumn(width="large",disabled=True),
                    "Observaciones":st.column_config.TextColumn(width="medium"),
                },key="edit_audit")
            if st.button("💾 Guardar cambios de auditoría",type="secondary"):
                # Reconvertir texto legible de vuelta a JSON no es trivial;
                # solo guardamos cambios en Dieta, Conductor, Fecha y Observaciones
                df_orig=df_brutos.copy()
                if len(df_edit)==len(df_orig):
                    df_orig["conductor"]=df_edit["Conductor"].values
                    df_orig["fecha"]=df_edit["Fecha"].values
                    df_orig["dieta"]=df_edit["Dieta"].values
                    df_orig["observaciones"]=df_edit["Observaciones"].values
                    if guardar_tabla_completa(df_orig): st.success("✅ Cambios guardados.")
                else:
                    # Filas eliminadas
                    if guardar_tabla_completa(df_orig.iloc[:len(df_edit)]): st.success("✅ Cambios guardados.")

def main():
    if "vista_actual" not in st.session_state: st.session_state.vista_actual="conductor"
    params=st.query_params
    if params.get("admin")=="1" and st.session_state.vista_actual=="conductor":
        st.session_state.vista_actual="admin"
    if st.session_state.vista_actual=="admin":
        with st.sidebar:
            st.markdown("## Autocares Alegre"); st.markdown("---")
            if st.button("Vista Conductor",use_container_width=True):
                st.session_state.vista_actual="conductor"; st.query_params.clear(); st.rerun()
            st.caption("v8.0")
        vista_admin()
    else:
        vista_conductor()

if __name__=="__main__":
    main()
