import streamlit as st
import pandas as pd
import os, subprocess, io, base64, smtplib, json, uuid
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from email.message import EmailMessage
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from streamlit_drawable_canvas import st_canvas
from PIL import Image
import gspread
from google.oauth2.service_account import Credentials

# =============================================================================
# 0. CONFIGURACIÓN INICIAL Y ESTILOS (CORREGIDO PARA EDGE)
# =============================================================================
st.set_page_config(page_title="Atlas Spence | Reportes", layout="wide", page_icon="⚙️", initial_sidebar_state="expanded")

def aplicar_estilos_premium():
    st.markdown("""
        <style>
        :root { --ac-blue: #007CA6; --ac-dark: #005675; --ac-light: #e6f2f7; --bhp-orange: #FF6600; }
        #MainMenu {visibility: hidden;} footer {visibility: hidden;}
        
        div.stButton > button:first-child {
            background-color: var(--ac-blue); color: white; border-radius: 6px; border: none;
            font-weight: 600; padding: 0.5rem 1rem; transition: all 0.3s ease; box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        div.stButton > button:first-child:hover { background-color: var(--ac-dark); transform: translateY(-2px); box-shadow: 0 4px 8px rgba(0,0,0,0.2); }
        .stTextInput>div>div>input:focus, .stNumberInput>div>div>input:focus, .stSelectbox>div>div>select:focus { border-color: var(--ac-blue) !important; box-shadow: 0 0 0 1px var(--ac-blue) !important; }
        h1 { border-bottom: 3px solid var(--ac-blue); padding-bottom: 10px; font-weight: 900; }
        h2, h3 { font-family: 'Segoe UI', sans-serif; font-weight: 700; }
        .stTabs [data-baseweb="tab-list"] { gap: 24px; }
        .stTabs [data-baseweb="tab"] { height: 50px; border-radius: 4px 4px 0 0; padding: 10px; }
        .stTabs [aria-selected="true"] { background-color: var(--ac-light); border-bottom: 3px solid var(--ac-blue); color: var(--ac-dark); font-weight: 600; }
        </style>
    """, unsafe_allow_html=True)

aplicar_estilos_premium()

# =============================================================================
# 1. CONSTANTES Y MAESTROS DE DATOS
# =============================================================================
RUTA_ONEDRIVE = "Reportes_Temporales" 
MI_CORREO_CORPORATIVO = "ignacio.a.morales@atlascopco.com"  
CORREO_REMITENTE = "informeatlas.spence@gmail.com"  
PASSWORD_APLICACION = "jbumdljbdpyomnna"  

USUARIOS = {"ignacio morales": "spence2026", "emian": "spence2026", "ignacio veas": "spence2026", "admin": "admin123"}

DEFAULT_SPECS = {
    "GA 18": {"Litros de Aceite": "14.1 L", "Cant. Filtros Aceite": "1", "N° Parte Filtro Aceite": "1625 4800 00", "Tipo de Aceite": "Roto Inject Fluid"},
    "GA 30": {"Litros de Aceite": "14.6 L", "Cant. Filtros Aceite": "1", "N° Parte Filtro Aceite": "1613 6105 00", "Tipo de Aceite": "Indurance - Xtend Duty"},
    "GA 37": {"Litros de Aceite": "14.6 L", "N° Parte Filtro Aceite": "1613 6105 00", "N° Parte Filtro Aire": "1613 7407 00", "Tipo de Aceite": "Indurance - Xtend Duty"},
    "GA 45": {"Litros de Aceite": "17.9 L", "Cant. Filtros Aceite": "1", "N° Parte Filtro Aceite": "1613 6105 00", "Tipo de Aceite": "Indurance - Xtend Duty"},
    "GA 75": {"Litros de Aceite": "35.2 L"},
    "GA 90": {"Litros de Aceite": "69 L", "Cant. Filtros Aceite": "3", "N° Parte Filtro Aceite": "1613 6105 00"},
    "GA 132": {"Litros de Aceite": "93 L", "Cant. Filtros Aceite": "3", "N° Parte Filtro Aceite": "1613 6105 90", "Tipo de Aceite": "Indurance - Xtend Duty"},
    "GA 250": {"Litros de Aceite": "130 L", "Cant. Filtros Aceite": "3", "Cant. Filtros Aire": "2", "Tipo de Aceite": "Indurance"},
    "ZT 37": {"Litros de Aceite": "23 L", "Cant. Filtros Aceite": "1", "N° Parte Filtro Aceite": "1614 8747 00", "Tipo de Aceite": "Roto Z fluid"},
    "CD 80+": {"Filtro de Gases": "DD/PD 80", "Desecante": "Alúmina", "Kit Válvulas": "2901 1622 00"},
    "CD 630": {"Filtro de Gases": "DD/PD 630", "Desecante": "Alúmina", "Kit Válvulas": "2901 1625 00"}
}

inventario_equipos = {
    "20-GC-001": ["GA 75", "AII482673", "truck shop", "mina"], "20-GC-002": ["GA 75", "AII482674", "truck shop", "mina"], "20-GC-003": ["GA 90", "AIF095178", "truck shop", "mina"], "20-GC-004": ["GA 37", "AII390776", "truck shop", "mina"],
    "35-GC-006": ["GA 250", "AIF095420", "chancado secundario", "área seca"], "35-GC-007": ["GA 250", "AIF095421", "chancado secundario", "área seca"], "35-GC-008": ["GA 250", "AIF095302", "chancado secundario", "área seca"],
    "50-GC-001": ["GA 45", "API542705", "planta SX", "área húmeda"], "50-GC-002": ["GA 45", "API542706", "planta SX", "área húmeda"], "50-GC-003": ["ZT 37", "API791692", "planta SX", "área húmeda"], "50-GC-004": ["ZT 37", "API791693", "planta SX", "área húmeda"], "50-CD-001": ["CD 80+", "API095825", "planta SX", "área húmeda"], "50-CD-002": ["CD 80+", "API095826", "planta SX", "área húmeda"],
    "55-GC-015": ["GA 30", "API501440", "planta borra", "área húmeda"],
    "65-GC-009": ["GA 250", "APF253608", "patio de estanques", "área húmeda"], "65-GC-011": ["GA 250", "APF253581", "patio de estanques", "área húmeda"], "65-CD-011": ["CD 630", "WXF300015", "patio de estanques", "área húmeda"], "65-CD-012": ["CD 630", "WXF300016", "patio de estanques", "área húmeda"],
    "70-GC-013": ["GA 132", "AIF095296", "descarga de acido", "área húmeda"], "70-GC-014": ["GA 132", "AIF095297", "descarga de acido", "área húmeda"],
    "80-GC-001": ["GA 18", "API335343", "laboratorio", "taller mecánico"]
}

# =============================================================================
# 2. GOOGLE SHEETS Y CACHÉ OPTIMIZADO (A PRUEBA DE FALLOS)
# =============================================================================
@st.cache_resource(show_spinner=False)
def get_gspread_client():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = Credentials.from_service_account_info(json.loads(st.secrets["gcp_json"]), scopes=scope)
    return gspread.authorize(creds)

def get_sheet(sheet_name):
    try:
        doc = get_gspread_client().open("Base_Datos_InforGem")
        if sheet_name in [hoja.title for hoja in doc.worksheets()]: return doc.worksheet(sheet_name)
        return doc.add_worksheet(title=sheet_name, rows="1000", cols="20")
    except Exception as e:
        if "200" in str(e): st.error("🚨 El archivo debe ser Google Sheets nativo, no Excel (.xlsx).")
        return None

# --- Helpers de Lectura/Escritura ---
def guardar_dato(sheet_name, row_data):
    try:
        hoja = get_sheet(sheet_name)
        if hoja: 
            hoja.append_row(row_data)
            st.cache_data.clear()
    except: pass

@st.cache_data(ttl=60, show_spinner=False)
def get_contactos():
    try:
        hoja = get_sheet("contactos")
        if not hoja: return []
        return sorted(list(set([r[0] for r in hoja.get_all_values() if len(r)>1 and r[1]=="ACTIVO"])))
    except: return []

@st.cache_data(ttl=60, show_spinner=False)
def get_estados_equipos():
    try:
        hoja = get_sheet("intervenciones")
        if not hoja: return {}
        return {r[0]: r[17] for r in hoja.get_all_values() if len(r) >= 18}
    except: return {}

@st.cache_data(ttl=60, show_spinner=False)
def get_historial(tag):
    try:
        hoja = get_sheet("intervenciones")
        if not hoja: return pd.DataFrame()
        return pd.DataFrame([{"fecha": r[5], "tipo": r[15], "estado": r[17], "user": r[19], "h_marcha": r[12]} for r in hoja.get_all_values() if len(r) >= 20 and r[0] == tag][::-1])
    except: return pd.DataFrame()

@st.cache_data(ttl=60, show_spinner=False)
def get_ultimo_registro(tag):
    try:
        hoja = get_sheet("intervenciones")
        if not hoja: return None
        for r in reversed(hoja.get_all_values()):
            if len(r) >= 20 and r[0] == tag: return r
    except: pass
    return None

@st.cache_data(ttl=60, show_spinner=False)
def get_especificaciones():
    specs = {k: dict(v) for k, v in DEFAULT_SPECS.items()}
    try:
        hoja = get_sheet("especificaciones")
        if not hoja: return specs
        for r in hoja.get_all_values():
            if len(r) >= 3:
                if r[0] not in specs: specs[r[0]] = {}
                specs[r[0]][r[1]] = r[2]
    except: pass
    return specs

@st.cache_data(ttl=60, show_spinner=False)
def get_datos_area(tag):
    try:
        hoja = get_sheet("datos_equipo")
        if not hoja: return {}
        return {r[1]: r[2] for r in hoja.get_all_values() if len(r) >= 3 and r[0] == tag}
    except: return {}

@st.cache_data(ttl=60, show_spinner=False)
def get_observaciones(tag):
    cols = ["id", "fecha", "user", "texto"]
    try:
        hoja = get_sheet("observaciones")
        if not hoja: return pd.DataFrame(columns=cols)
        data = hoja.get_all_values()
        obs = [{"id": r[0], "fecha": r[2], "user": r[3], "texto": r[4]} for r in data if len(r)>=6 and r[1]==tag and r[5]=="ACTIVO"]
        df = pd.DataFrame(obs)
        if not df.empty: return df.iloc[::-1]
        return pd.DataFrame(columns=cols)
    except: 
        return pd.DataFrame(columns=cols)

# =============================================================================
# 3. UTILIDADES (PDF Y CORREO)
# =============================================================================
def convertir_a_pdf(ruta_docx):
    ruta_pdf = ruta_docx.replace(".docx", ".pdf")
    try:
        subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', os.path.abspath(ruta_docx), '--outdir', os.path.dirname(os.path.abspath(ruta_docx))], capture_output=True)
        if os.path.exists(ruta_pdf): return ruta_pdf
    except: pass
    return None

def enviar_informes_correo(destinatario, informes):
    msg = MIMEMultipart()
    msg['From'], msg['To'], msg['Subject'] = CORREO_REMITENTE, destinatario, f"Reportes Firmados - {pd.Timestamp.now().strftime('%d/%m/%Y')}"
    cuerpo = f"Se adjuntan {len(informes)} reportes firmados.\n\n" + "\n".join([f"- {i['tag']} ({i['tipo']})" for i in informes])
    msg.attach(MIMEText(cuerpo, 'plain'))

    for i in informes:
        if os.path.exists(i['ruta']):
            with open(i['ruta'], "rb") as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{i["nombre"].replace("ó","o").replace("í","i")}"')
            msg.attach(part)
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(CORREO_REMITENTE, PASSWORD_APLICACION)
        server.send_message(msg)
        server.quit()
        return True, "Enviado con éxito."
    except Exception as e: return False, str(e)

# =============================================================================
# 4. VARIABLES DE SESIÓN Y NAVEGACIÓN
# =============================================================================
default_states = {'logged_in': False, 'usuario_actual': "", 'equipo_seleccionado': None, 'informes_pendientes': [], 'vista_firmas': False, 'input_cliente': "Lorena Rojas", 'input_tec1': "", 'input_tec2': "", 'input_h_marcha': 0, 'input_h_carga': 0, 'input_temp': "70.0", 'input_p_carga': "7.0", 'input_p_descarga': "7.5", 'input_estado': "", 'input_reco': "", 'input_estado_eq': "Operativo"}
for k, v in default_states.items():
    if k not in st.session_state: st.session_state[k] = v

def navegar(tag=None):
    st.session_state.equipo_seleccionado = tag
    st.session_state.vista_firmas = False
    if tag:
        reg = get_ultimo_registro(tag)
        if reg:
            st.session_state.update({'input_cliente': reg[1], 'input_tec1': reg[5], 'input_tec2': reg[6], 'input_estado': reg[3], 'input_reco': reg[11] or "", 'input_estado_eq': reg[12] or "Operativo", 'input_h_marcha': int(reg[9] or 0), 'input_h_carga': int(reg[10] or 0), 'input_temp': str(reg[2]).replace(',', '.') if reg[2] else "70.0"})
            try: st.session_state.input_p_carga = str(reg[7]).split()[0].replace(',', '.')
            except: st.session_state.input_p_carga = "7.0"
            try: st.session_state.input_p_descarga = str(reg[8]).split()[0].replace(',', '.')
            except: st.session_state.input_p_descarga = "7.5"
        else:
            st.session_state.update({'input_estado_eq': "Operativo", 'input_estado': "", 'input_reco': "", 'input_h_marcha': 0, 'input_h_carga': 0})
            # =============================================================================
# 5. PANTALLA DE LOGIN
# =============================================================================
if not st.session_state.logged_in:
    st.markdown("<br><br>", unsafe_allow_html=True)
    _, col_centro, _ = st.columns([1, 1.5, 1])
    with col_centro.container(border=True):
        st.markdown("<h1 style='text-align: center; border:none;'>⚙️ <span style='color:#007CA6;'>Atlas</span> <span style='color:#FF6600;'>Spence</span></h1>", unsafe_allow_html=True)
        with st.form("login"):
            usr = st.text_input("Usuario Corporativo").lower()
            pwd = st.text_input("Contraseña", type="password")
            if st.form_submit_button("Ingresar", type="primary", use_container_width=True):
                if usr in USUARIOS and USUARIOS[usr] == pwd:
                    st.session_state.update({'logged_in': True, 'usuario_actual': usr})
                    st.rerun()
                else: st.error("❌ Credenciales inválidas.")

# =============================================================================
# 6. APLICACIÓN PRINCIPAL
# =============================================================================
else:
    # --- BARRA LATERAL LATERAL ---
    with st.sidebar:
        st.markdown("<h2><span style='color:#007CA6;'>Atlas</span> <span style='color:#FF6600;'>Spence</span></h2>", unsafe_allow_html=True)
        st.caption(f"👤 {st.session_state.usuario_actual.title()}")
        if st.session_state.informes_pendientes:
            st.warning(f"📝 {len(st.session_state.informes_pendientes)} por firmar")
            if st.button("✍️ Pizarra de Firmas", type="primary", use_container_width=True):
                st.session_state.vista_firmas = True; st.session_state.equipo_seleccionado = None; st.rerun()
        st.markdown("---")
        if st.button("🚪 Salir", use_container_width=True): st.session_state.logged_in = False; st.rerun()

    # --- 6.1 VISTA DE FIRMAS ---
    if st.session_state.vista_firmas:
        st.button("⬅️ Volver al Panel", on_click=navegar)
        st.title("✍️ Pizarra de Firmas")
        for i, inf in enumerate(st.session_state.informes_pendientes):
            with st.expander(f"📄 Previsualizar: {inf['tag']}"):
                if inf.get('ruta_prev_pdf') and os.path.exists(inf['ruta_prev_pdf']):
                    with open(inf['ruta_prev_pdf'], "rb") as f:
                        st.markdown(f'<iframe src="data:application/pdf;base64,{base64.b64encode(f.read()).decode()}" width="100%" height="500"></iframe>', unsafe_allow_html=True)
                        f.seek(0)
                        st.download_button("📥 Descargar PDF", f, f"Borrador_{inf['tag']}.pdf", "application/pdf", key=f"dl_{i}")
        
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### 🧑‍🔧 Técnico")
            can_tec = st_canvas(stroke_width=3, height=150, width=350, key="tec")
        with c2:
            st.markdown("#### 👷 Cliente")
            can_cli = st_canvas(stroke_width=3, height=150, width=350, key="cli")

        if st.button("🚀 Firmar y Subir a Nube", type="primary", use_container_width=True):
            if can_tec.image_data is not None and can_cli.image_data is not None:
                with st.spinner("Procesando documentos oficiales..."):
                    informes_finales = []
                    for inf in st.session_state.informes_pendientes:
                        # Procesar imágenes
                        img_t, img_c = io.BytesIO(), io.BytesIO()
                        Image.fromarray(can_tec.image_data.astype('uint8'), 'RGBA').save(img_t, 'PNG')
                        Image.fromarray(can_cli.image_data.astype('uint8'), 'RGBA').save(img_c, 'PNG')
                        
                        doc = DocxTemplate(inf['plantilla'])
                        ctx = inf['contexto']
                        ctx.update({'firma_tecnico': InlineImage(doc, img_t, width=Mm(40)), 'firma_cliente': InlineImage(doc, img_c, width=Mm(40))})
                        doc.render(ctx)
                        doc.save(inf['ruta_docx'])
                        
                        ruta_pdf = convertir_a_pdf(inf['ruta_docx']) or inf['ruta_docx']
                        nombre_final = inf['nombre'].replace(".docx", ".pdf")
                        
                        tupla_db = list(inf['tupla_db'])
                        tupla_db[18] = ruta_pdf
                        guardar_dato("intervenciones", tupla_db)
                        
                        informes_finales.append({"tag": inf['tag'], "tipo": inf['tipo'], "ruta": ruta_pdf, "nombre": f"{inf['area']}@@{inf['tag']}@@{nombre_final}"})
                    
                    exito, msg = enviar_informes_correo(MI_CORREO_CORPORATIVO, informes_finales)
                    if exito: st.success("✅ Guardado y enviado!"); st.session_state.informes_pendientes = []; st.balloons()
                    else: st.error(f"Error correo: {msg}")
            else: st.warning("Faltan firmas.")

    # --- 6.2 PANEL PRINCIPAL (DASHBOARD) ---
    elif not st.session_state.equipo_seleccionado:
        st.title("🏭 Panel de Equipos")
        estados = get_estados_equipos()
        
        c1, c2 = st.columns([1, 2])
        filtro = c1.radio("Filtro:", ["Todos", "Compresores", "Secadores"], horizontal=True)
        busqueda = c2.text_input("🔍 Buscar TAG o Área...").lower()
        
        st.markdown("<br>", unsafe_allow_html=True)
        cols = st.columns(4)
        idx = 0
        for tag, (mod, ser, area, ubi) in inventario_equipos.items():
            es_cd = "CD" in mod.upper()
            if (filtro == "Compresores" and es_cd) or (filtro == "Secadores" and not es_cd): continue
            if busqueda in tag.lower() or busqueda in area.lower() or busqueda in mod.lower():
                est = estados.get(tag, "Operativo")
                bg, txt, ico = ("#eaffea", "#004d00", "🟢") if est == "Operativo" else ("#ffeaea", "#800000", "🔴")
                with cols[idx % 4].container(border=True):
                    st.markdown(f"<span style='background:{bg}; color:{txt}; padding:3px 6px; border-radius:4px; font-size:12px;'><b>{ico} {est.upper()}</b></span>", unsafe_allow_html=True)
                    st.markdown(f"<h3 style='margin:10px 0 0 0;'>{tag}</h3><p style='color:gray; font-size:14px; margin:0 0 10px 0;'>{mod} | {area.title()}</p>", unsafe_allow_html=True)
                    st.button("Ingresar", key=f"btn_{tag}", on_click=navegar, args=(tag,), use_container_width=True)
                idx += 1

    # --- 6.3 FICHA DEL EQUIPO (4 PESTAÑAS) ---
    else:
        tag = st.session_state.equipo_seleccionado
        mod, ser, area, ubi = inventario_equipos[tag]
        
        c_btn, c_tit = st.columns([1, 4])
        c_btn.button("⬅️ Volver", on_click=navegar, use_container_width=True)
        c_tit.markdown(f"<h1 style='margin-top:-15px;'>Ficha: <span style='color:#007CA6;'>{tag}</span></h1>", unsafe_allow_html=True)

        t1, t2, t3, t4 = st.tabs(["📋 1. Reporte", "📚 2. Especificaciones", "🔍 3. Bitácora", "👤 4. Área"])
        
        # PESTAÑA 1: REPORTE
        with t1:
            plan = st.selectbox("Orden:", ["PM03", "Inspección"] if "CD" in tag else ["PM03", "Inspección", "P1", "P2", "P3"])
            c1, c2, c3, c4 = st.columns(4)
            c1.text_input("Modelo", mod, disabled=True); c2.text_input("Serie", ser, disabled=True); c3.text_input("Área", area, disabled=True); c4.text_input("Ubicación", ubi, disabled=True)
            
            c5, c6, c7, c8 = st.columns([1,1,1,1.5])
            fecha = c5.text_input("Fecha", "25 de febrero de 2026")
            tec1 = c6.text_input("Técnico 1", st.session_state.input_tec1)
            tec2 = c7.text_input("Técnico 2", st.session_state.input_tec2)
            
            contactos = get_contactos()
            ops_cli = ["➕ Nuevo..."] + (contactos if contactos else ["Lorena Rojas"])
            cli_sel = c8.selectbox("Cliente", ops_cli, index=ops_cli.index(st.session_state.input_cliente) if st.session_state.input_cliente in ops_cli else 1)
            if cli_sel == "➕ Nuevo...":
                cli_final = c8.text_input("Escribe el nombre:")
                if c8.button("Guardar Contacto") and cli_final: guardar_dato("contactos", [cli_final.title(), "ACTIVO"]); st.rerun()
            else: cli_final = cli_sel

            st.markdown("---")
            m1, m2, m3, m4, m5, m6 = st.columns(6)
            hm = m1.number_input("H. Marcha", value=st.session_state.input_h_marcha, step=1)
            hc = m2.number_input("H. Carga", value=st.session_state.input_h_carga, step=1)
            up = m3.selectbox("Unidad", ["Bar", "psi"])
            pc = m4.text_input("P. Carga", st.session_state.input_p_carga)
            pd = m5.text_input("P. Descarga", st.session_state.input_p_descarga)
            ts = m6.text_input("Temp Salida", st.session_state.input_temp)

            st.markdown("---")
            est_eq = st.radio("Estado de Entrega:", ["Operativo", "Fuera de servicio"], horizontal=True)
            est_txt = st.text_area("Condición Final:", st.session_state.input_estado)
            reco = st.text_area("Recomendaciones:", st.session_state.input_reco)

            if st.button("📥 Añadir a Bandeja", type="primary", use_container_width=True):
                if "CD" in tag: tpl = "plantilla/secadorfueradeservicio.docx" if est_eq == "Fuera de servicio" else "plantilla/inspeccionsecador.docx"
                else: tpl = f"plantilla/{plan.lower()}.docx" if est_eq == "Operativo" and plan in ["P1","P2","P3"] else ("plantilla/fueradeservicio.docx" if est_eq == "Fuera de servicio" else "plantilla/inspeccion.docx")
                
                ctx = {"tipo_intervencion": plan, "modelo": mod, "tag": tag, "area": area, "ubicacion": ubi, "cliente_contacto": cli_final, "p_carga": f"{pc} {up}", "p_descarga": f"{pd} {up}", "temp_salida": ts, "horas_marcha": hm, "horas_carga": hc, "tecnico_1": tec1, "tecnico_2": tec2, "estado_equipo": est_eq, "estado_entrega": est_txt, "recomendaciones": reco, "serie": ser, "fecha": fecha, "firma_tecnico": "", "firma_cliente": ""}
                
                fname = f"Informe_{plan}_{tag}.docx"
                ruta_doc = os.path.join(RUTA_ONEDRIVE, fname)
                os.makedirs(RUTA_ONEDRIVE, exist_ok=True)
                
                with st.spinner("Generando Borrador..."):
                    d_prev = DocxTemplate(tpl); d_prev.render(ctx); d_prev.save(ruta_doc)
                    ruta_pdf = convertir_a_pdf(ruta_doc)
                
                t_db = (tag, mod, ser, area, ubi, fecha, cli_final, tec1, tec2, float(ts.replace(',','.')) if ts else 0.0, f"{pc} {up}", f"{pd} {up}", hm, hc, est_txt, plan, reco, est_eq, "", st.session_state.usuario_actual)
                
                st.session_state.informes_pendientes.append({"tag": tag, "area": area, "tipo": plan, "tec1": tec1, "cli": cli_final, "plantilla": tpl, "contexto": ctx, "tupla_db": t_db, "ruta_docx": ruta_doc, "nombre": fname, "ruta_prev_pdf": ruta_pdf})
                st.success("Guardado en bandeja."); navegar()

        # PESTAÑA 2: ESPECIFICACIONES
        with t2:
            with st.expander("✏️ Agregar Dato Técnico"):
                c1, c2 = st.columns(2)
                k = c1.selectbox("Dato:", ["N° Parte Filtro Aceite", "N° Parte Kit", "Litros de Aceite", "Otro..."])
                if k == "Otro...": k = c1.text_input("Nombre del dato:")
                v = c2.text_input("Valor:")
                if st.button("Guardar Dato"): guardar_dato("especificaciones", [mod, k, v]); st.rerun()
            
            specs = get_especificaciones().get(mod, {})
            cols = st.columns(3)
            for i, (key, val) in enumerate(specs.items()):
                cols[i%3].info(f"**{key}**\n\n{val}")

        # PESTAÑA 3: BITÁCORA
        with t3:
            obs_txt = st.text_area("Nueva observación:")
            if st.button("➕ Guardar Nota"): guardar_dato("observaciones", [str(uuid.uuid4())[:8], tag, pd.Timestamp.now().strftime("%d/%m/%Y %H:%M"), st.session_state.usuario_actual, obs_txt, "ACTIVO"]); st.rerun()
            
            for _, r in get_observaciones(tag).iterrows():
                st.markdown(f"> **{r['fecha']} | {r['user']}**: {r['texto']}")

        # PESTAÑA 4: ÁREA
        with t4:
            with st.expander("✏️ Actualizar Área"):
                c1, c2 = st.columns(2)
                k_area = c1.selectbox("Campo:", ["Dueño de Área", "PEA", "Frecuencia Radial", "Otro..."])
                if k_area == "Otro...": k_area = c1.text_input("Nombre:")
                v_area = c2.text_input("Info:")
                if st.button("Guardar Info"): guardar_dato("datos_equipo", [tag, k_area, v_area]); st.rerun()
                
            for k, v in get_datos_area(tag).items(): st.success(f"**{k}**: {v}")