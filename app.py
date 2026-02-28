import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import os, subprocess
import pandas as pd
import smtplib
from email.message import EmailMessage
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from streamlit_drawable_canvas import st_canvas
from PIL import Image
import io
import base64
import gspread
from google.oauth2.service_account import Credentials
import json
import uuid

# =============================================================================
# 0.1 CONFIGURACIÓN DE NUBE Y CORREO
# =============================================================================
RUTA_ONEDRIVE = "Reportes_Temporales" 
MI_CORREO_CORPORATIVO = "ignacio.a.morales@atlascopco.com"  
CORREO_REMITENTE = "informeatlas.spence@gmail.com"  
PASSWORD_APLICACION = "jbumdljbdpyomnna"  

def enviar_carrito_por_correo(destinatario, lista_informes):
    msg = MIMEMultipart()
    msg['From'] = CORREO_REMITENTE
    msg['To'] = destinatario
    msg['Subject'] = f"REVISIÓN PREVIA: Reportes Atlas Copco - Firmados - {pd.Timestamp.now().strftime('%d/%m/%Y')}"

    cuerpo = f"Estimado/a,\n\nSe adjuntan {len(lista_informes)} reportes de servicio técnico (Firmados) generados en la presente jornada para su revisión previa.\n\nEquipos intervenidos:\n"
    for item in lista_informes:
        cuerpo += f"- TAG: {item['tag']} | Orden: {item['tipo']}\n"
    cuerpo += "\nSaludos cordiales,\nSistema Integrado InforGem"

    msg.attach(MIMEText(cuerpo, 'plain'))

    for item in lista_informes:
        ruta = item['ruta']
        nombre_seguro = item["nombre_archivo"].replace("ó","o").replace("í","i").replace("á","a").replace("é","e").replace("ú","u")
        
        if os.path.exists(ruta):
            with open(ruta, "rb") as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Type', 'application/octet-stream', name=nombre_seguro)
            part.add_header('Content-Disposition', f'attachment; filename="{nombre_seguro}"')
            msg.attach(part)

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(CORREO_REMITENTE, PASSWORD_APLICACION)
        server.send_message(msg)
        server.quit()
        return True, "✅ Todos los informes fueron enviados a tu correo corporativo."
    except Exception as e:
        return False, f"❌ Error al enviar el correo: {e}"

# =============================================================================
# 0.2 ESTILOS PREMIUM
# =============================================================================
st.set_page_config(page_title="Atlas Spence | Gestión de Reportes", layout="wide", page_icon="⚙️")

def aplicar_estilos_premium():
    st.markdown("""
        <meta name="google" content="notranslate">
        <style>
        :root { --ac-blue: #007CA6; --ac-dark: #005675; --ac-light: #e6f2f7; --bhp-orange: #FF6600; }
        #MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}
        div.stButton > button:first-child {
            background-color: var(--ac-blue); color: white; border-radius: 6px; border: none;
            font-weight: 600; padding: 0.5rem 1rem; transition: all 0.3s ease; box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        div.stButton > button:first-child:hover { background-color: var(--ac-dark); transform: translateY(-2px); box-shadow: 0 4px 8px rgba(0,0,0,0.2); }
        .stTextInput>div>div>input:focus, .stNumberInput>div>div>input:focus, .stSelectbox>div>div>select:focus { border-color: var(--ac-blue) !important; box-shadow: 0 0 0 1px var(--ac-blue) !important; }
        h1, h2, h3 { font-family: 'Segoe UI', sans-serif; font-weight: 700; }
        h1 { border-bottom: 3px solid var(--ac-blue); padding-bottom: 10px; }
        .stTabs [data-baseweb="tab-list"] { gap: 24px; }
        .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; border-radius: 4px 4px 0 0; padding-top: 10px; padding-bottom: 10px; }
        .stTabs [aria-selected="true"] { background-color: var(--ac-light); border-bottom: 3px solid var(--ac-blue); color: var(--ac-dark); font-weight: 600; }
        </style>
    """, unsafe_allow_html=True)

aplicar_estilos_premium()

# =============================================================================
# 1. DATOS MAESTROS
# =============================================================================
USUARIOS = {"ignacio morales": "spence2026", "emian": "spence2026", "ignacio veas": "spence2026", "admin": "admin123"}

DEFAULT_SPECS = {
    "GA 18": {"Litros de Aceite": "14.1 L", "Cant. Filtros Aceite": "1", "N° Parte Filtro Aceite": "1625 4800 00 / 1625 7525 01", "Cant. Filtros Aire": "1", "N° Parte Filtro Aire": "1630 2201 36 / 1625 2204 36", "Tipo de Aceite": "Roto Inject Fluid", "Manual": "manuales/manual_ga18.pdf"},
    "GA 30": {"Litros de Aceite": "14.6 L", "Cant. Filtros Aceite": "1", "N° Parte Filtro Aceite": "1613 6105 00", "Cant. Filtros Aire": "1", "N° Parte Filtro Aire": "1613 7407 00", "N° Parte Kit": "2901-0326-00 / 2901 0325 00", "Tipo de Aceite": "Indurance - Xtend Duty", "Manual": "manuales/manual_ga30.pdf"},
    "GA 37": {"Litros de Aceite": "14.6 L", "N° Parte Filtro Aceite": "1613 6105 00", "N° Parte Filtro Aire": "1613 7407 00", "N° Parte Separador": "1613 7408 00", "N° Parte Kit": "2901 1626 00 / 10-1613 8397 02", "Tipo de Aceite": "Indurance - Xtend Duty", "Manual": "manuales/manual_ga37.pdf"},
    "GA 45": {"Litros de Aceite": "17.9 L", "Cant. Filtros Aceite": "1", "N° Parte Filtro Aceite": "1613 6105 00", "Cant. Filtros Aire": "1", "N° Parte Kit": "2901-0326-00 / 2901 0325 00", "Tipo de Aceite": "Indurance - Xtend Duty", "Manual": "manuales/manual_ga45.pdf"},
    "GA 75": {"Litros de Aceite": "35.2 L", "Manual": "manuales/manual_ga75.pdf"},
    "GA 90": {"Litros de Aceite": "69 L", "Cant. Filtros Aceite": "3", "N° Parte Filtro Aceite": "1613 6105 00", "N° Parte Filtro Aire": "2914 5077 00", "N° Parte Kit": "2901-0776-00", "Manual": "manuales/manual_ga90.pdf"},
    "GA 132": {"Litros de Aceite": "93 L", "Cant. Filtros Aceite": "3", "N° Parte Filtro Aceite": "1613 6105 90", "Cant. Filtros Aire": "1", "N° Parte Filtro Aire": "2914 5077 00", "N° Parte Kit": "2906 0604 00", "Tipo de Aceite": "Indurance / Indurance - Xtend Duty", "Manual": "manuales/manual_ga132.pdf"},
    "GA 250": {"Litros de Aceite": "130 L", "Cant. Filtros Aceite": "3", "Cant. Filtros Aire": "2", "Tipo de Aceite": "Indurance", "Manual": "manuales/manual_ga250.pdf"},
    "ZT 37": {"Litros de Aceite": "23 L", "Cant. Filtros Aceite": "1", "N° Parte Filtro Aceite": "1614 8747 00", "Cant. Filtros Aire": "1", "N° Parte Filtro Aire": "1613 7407 00", "N° Parte Kit": "2901-1122-00", "Tipo de Aceite": "Roto Z fluid", "Manual": "manuales/manual_zt37.pdf"},
    "CD 80+": {"Filtro de Gases": "DD/PD 80", "Desecante": "Alúmina", "Kit Válvulas": "2901 1622 00", "Silenciador": "1621 1234 00", "Manual": "manuales/manual_cd80.pdf"},
    "CD 630": {"Filtro de Gases": "DD/PD 630", "Desecante": "Alúmina", "Kit Válvulas": "2901 1625 00", "Silenciador": "1621 1235 00", "Manual": "manuales/manual_cd630.pdf"}
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
# 2. CONEXIÓN INMORTAL Y OPTIMIZADA A GOOGLE SHEETS (ANTI-BLOQUEOS)
# =============================================================================
@st.cache_resource
def get_gspread_client():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds_dict = json.loads(st.secrets["gcp_json"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return gspread.authorize(creds)

def get_sheet(sheet_name):
    try:
        client = get_gspread_client()
        doc = client.open("BaseDatos")
        pestañas = [hoja.title for hoja in doc.worksheets()]
        if sheet_name in pestañas: return doc.worksheet(sheet_name)
        else: return doc.add_worksheet(title=sheet_name, rows="1000", cols="20")
    except Exception as e:
        if "200" in str(e): st.error("🚨 ERROR DE FORMATO: Tu archivo en Google Drive es un Excel tradicional (.xlsx). Debes crear una 'Hoja de cálculo de Google' nativa.")
        else: st.error(f"🚨 ERROR DE CONEXIÓN CON GOOGLE: {e}")
        return None

# --- Funciones de Gestión de Área ---
def guardar_dato_equipo(tag, clave, valor):
    try:
        sheet = get_sheet("datos_equipo")
        sheet.append_row([tag, clave, valor])
        st.cache_data.clear() # Limpiamos memoria para ver los cambios de inmediato
    except: pass

@st.cache_data(ttl=30, show_spinner=False)
def obtener_datos_equipo(tag):
    datos = {}
    try:
        sheet = get_sheet("datos_equipo")
        data = sheet.get_all_values()
        for row in data:
            if len(row) >= 3 and row[0] == tag: datos[row[1]] = row[2]
    except: pass
    return datos

# --- Funciones de Bitácora ---
def agregar_observacion(tag, usuario, texto):
    if not texto.strip(): return
    fecha_actual = pd.Timestamp.now().strftime("%d/%m/%Y %H:%M")
    id_obs = str(uuid.uuid4())[:8]
    try:
        sheet = get_sheet("observaciones")
        sheet.append_row([id_obs, tag, fecha_actual, usuario.title(), texto.strip(), "ACTIVO"])
        st.cache_data.clear()
    except: pass

@st.cache_data(ttl=30, show_spinner=False)
def obtener_observaciones(tag):
    try:
        sheet = get_sheet("observaciones")
        data = sheet.get_all_values()
        obs = [{"id": r[0], "fecha": r[2], "usuario": r[3], "texto": r[4]} for r in data if len(r) >= 6 and r[1] == tag and r[5] == "ACTIVO"]
        df = pd.DataFrame(obs)
        if not df.empty: return df.iloc[::-1]
        return pd.DataFrame(columns=["id", "fecha", "usuario", "texto"])
    except: return pd.DataFrame(columns=["id", "fecha", "usuario", "texto"])

def eliminar_observacion(id_obs):
    try:
        sheet = get_sheet("observaciones")
        cell = sheet.find(id_obs)
        if cell: sheet.update_cell(cell.row, 6, "ELIMINADO")
        st.cache_data.clear()
    except: pass

# --- Funciones de Especificaciones ---
def guardar_especificacion_db(modelo, clave, valor):
    try:
        sheet = get_sheet("especificaciones")
        sheet.append_row([modelo, clave, valor])
        st.cache_data.clear()
    except: pass

@st.cache_data(ttl=60, show_spinner=False)
def obtener_especificaciones(defaults):
    specs = {k: dict(v) for k, v in defaults.items()}
    try:
        sheet = get_sheet("especificaciones")
        for row in sheet.get_all_values():
            if len(row) >= 3:
                mod, clave, valor = row[0], row[1], row[2]
                if mod not in specs: specs[mod] = {}
                specs[mod][clave] = valor
    except: pass
    return specs

# --- Funciones de Contactos ---
@st.cache_data(ttl=30, show_spinner=False)
def obtener_contactos():
    try:
        sheet = get_sheet("contactos")
        contactos = [row[0] for row in sheet.get_all_values() if len(row) > 1 and row[1] == "ACTIVO"]
        return sorted(list(set(contactos))) if contactos else ["Lorena Rojas"]
    except: return ["Lorena Rojas"]

def agregar_contacto(nombre):
    if not nombre.strip(): return
    try:
        sheet = get_sheet("contactos")
        sheet.append_row([nombre.strip().title(), "ACTIVO"])
        st.cache_data.clear()
    except: pass

def eliminar_contacto(nombre):
    try:
        sheet = get_sheet("contactos")
        for cell in sheet.findall(nombre): sheet.update_cell(cell.row, 2, "ELIMINADO")
        st.cache_data.clear()
    except: pass

# --- Funciones de Historial de Intervenciones ---
def guardar_registro(data_tuple):
    try:
        sheet = get_sheet("intervenciones")
        sheet.append_row([str(x) for x in data_tuple])
        st.cache_data.clear()
    except: pass

@st.cache_data(ttl=30, show_spinner=False)
def buscar_ultimo_registro(tag):
    try:
        sheet = get_sheet("intervenciones")
        for row in reversed(sheet.get_all_values()):
            if len(row) >= 20 and row[0] == tag: return (row[5], row[6], row[9], row[14], row[15], row[7], row[8], row[10], row[11], row[12], row[13], row[16], row[17])
    except: pass
    return None

@st.cache_data(ttl=30, show_spinner=False)
def obtener_todo_el_historial(tag):
    try:
        sheet = get_sheet("intervenciones")
        hist = [{"fecha": r[5], "tipo_intervencion": r[15], "estado_equipo": r[17], "Cuenta Usuario": r[19], "horas_marcha": r[12], "horas_carga": r[13], "p_carga": r[10], "p_descarga": r[11], "temp_salida": r[9]} for r in sheet.get_all_values() if len(r) >= 20 and r[0] == tag]
        df = pd.DataFrame(hist)
        if not df.empty: return df.iloc[::-1]
        return pd.DataFrame()
    except: return pd.DataFrame()

@st.cache_data(ttl=30, show_spinner=False)
def obtener_estados_actuales():
    estados = {}
    try:
        sheet = get_sheet("intervenciones")
        for row in sheet.get_all_values():
            if len(row) >= 18: estados[row[0]] = row[17]
    except: pass
    return estados

# =============================================================================
# 3. CONVERSIÓN A PDF HÍBRIDA
# =============================================================================
def convertir_a_pdf(ruta_docx):
    ruta_pdf = ruta_docx.replace(".docx", ".pdf")
    ruta_absoluta = os.path.abspath(ruta_docx)
    carpeta_salida = os.path.dirname(ruta_absoluta)
    
    try:
        comando = ['libreoffice', '--headless', '--convert-to', 'pdf', ruta_absoluta, '--outdir', carpeta_salida]
        subprocess.run(comando, capture_output=True, text=True)
        if os.path.exists(ruta_pdf): return ruta_pdf
    except: pass

    try:
        import pythoncom
        from docx2pdf import convert
        pythoncom.CoInitialize()
        convert(ruta_absoluta, ruta_pdf)
        if os.path.exists(ruta_pdf): return ruta_pdf
    except: pass
        
    return None

# =============================================================================
# 4. INICIALIZACIÓN DE LA APLICACIÓN Y VARIABLES DE SESIÓN
# =============================================================================
ESPECIFICACIONES = obtener_especificaciones(DEFAULT_SPECS)

default_states = {
    'logged_in': False, 'usuario_actual': "", 'equipo_seleccionado': None,
    'input_cliente': "Lorena Rojas", 'input_tec1': "Ignacio Morales", 'input_tec2': "emian Sanchez",
    'input_h_marcha': 0, 'input_h_carga': 0, 'input_temp': "70.0",
    'input_p_carga': "7.0", 'input_p_descarga': "7.5", 'input_estado': "",
    'input_reco': "", 'input_estado_eq': "Operativo",
    'informes_pendientes': [], 'vista_firmas': False
}
for key, value in default_states.items():
    if key not in st.session_state: st.session_state[key] = value

def seleccionar_equipo(tag):
    st.session_state.equipo_seleccionado = tag
    st.session_state.vista_firmas = False
    reg = buscar_ultimo_registro(tag)
    if reg:
        st.session_state.input_cliente = reg[1]
        st.session_state.input_tec1 = reg[5]
        st.session_state.input_tec2 = reg[6]
        st.session_state.input_estado = reg[3]
        st.session_state.input_reco = reg[11] if reg[11] else ""
        st.session_state.input_estado_eq = reg[12] if reg[12] else "Operativo"
        st.session_state.input_h_marcha = int(reg[9]) if reg[9] else 0
        st.session_state.input_h_carga = int(reg[10]) if reg[10] else 0
        st.session_state.input_temp = str(reg[2]).replace(',', '.') if reg[2] is not None else "70.0"
        try: st.session_state.input_p_carga = str(reg[7]).split()[0].replace(',', '.')
        except: st.session_state.input_p_carga = "7.0"
        try: st.session_state.input_p_descarga = str(reg[8]).split()[0].replace(',', '.')
        except: st.session_state.input_p_descarga = "7.5"
    else:
        st.session_state.update({'input_estado_eq': "Operativo", 'input_estado': "", 'input_reco': ""})

def volver_catalogo(): 
    st.session_state.equipo_seleccionado = None
    st.session_state.vista_firmas = False
    # =============================================================================
# 5. PANTALLA 1: SISTEMA DE LOGIN PREMIUM
# =============================================================================
if not st.session_state.logged_in:
    st.markdown("<br><br><br>", unsafe_allow_html=True)
    _, col_centro, _ = st.columns([1, 1.5, 1])
    with col_centro:
        with st.container(border=True):
            st.markdown("<h1 style='text-align: center; border-bottom:none;'>⚙️ <span style='color:#007CA6;'>Atlas</span> <span style='color:#FF6600;'>Spence</span></h1>", unsafe_allow_html=True)
            st.markdown("<p style='text-align: center; color: gray;'>Sistema de Gestión de Reportes Técnicos - Atlas Copco</p>", unsafe_allow_html=True)
            st.markdown("---")
            with st.form("form_login"):
                u_in = st.text_input("Usuario Corporativo").lower()
                p_in = st.text_input("Contraseña", type="password")
                st.markdown("<br>", unsafe_allow_html=True)
                if st.form_submit_button("Acceder de forma segura", type="primary", use_container_width=True):
                    if u_in in USUARIOS and USUARIOS[u_in] == p_in:
                        st.session_state.update({'logged_in': True, 'usuario_actual': u_in})
                        st.rerun()
                    else: st.error("❌ Credenciales inválidas.")

# =============================================================================
# 6. PANTALLA PRINCIPAL: APLICACIÓN AUTENTICADA
# =============================================================================
else:
    with st.sidebar:
        st.markdown("<h2 style='text-align: center; border-bottom:none; margin-top: -20px;'><span style='color:#007CA6;'>Atlas</span> <span style='color:#FF6600;'>Spence</span></h2>", unsafe_allow_html=True)
        st.markdown(f"**Usuario Activo:**<br>{st.session_state.usuario_actual.title()}", unsafe_allow_html=True)

        if len(st.session_state.informes_pendientes) > 0:
            st.markdown("---")
            st.warning(f"📝 Tienes {len(st.session_state.informes_pendientes)} reportes esperando firmas.")
            if st.button("✍️ Ir a Pizarra de Firmas", use_container_width=True, type="primary"):
                st.session_state.vista_firmas = True
                st.session_state.equipo_seleccionado = None
                st.rerun()

        st.markdown("---")
        if st.button("🚪 Cerrar Sesión", use_container_width=True):
            st.session_state.logged_in = False
            st.rerun()

    # --- 6.1 VISTA DE FIRMAS Y ENVÍO MÚLTIPLE ---
    if st.session_state.vista_firmas:
        c_v1, c_v2 = st.columns([1,4])
        with c_v1: 
            if st.button("⬅️ Volver", use_container_width=True): volver_catalogo(); st.rerun()
        with c_v2: 
            st.markdown("<h1 style='margin-top:-15px;'>✍️ Pizarra de Firmas Digital</h1>", unsafe_allow_html=True)

        st.markdown("---")
        st.markdown(f"### 📑 Revisión de Informes ({len(st.session_state.informes_pendientes)})")
        st.info("👀 **Para el Cliente:** Por favor, revise el documento oficial antes de firmar.")

        for i, inf in enumerate(st.session_state.informes_pendientes):
            with st.expander(f"📄 Ver documento preliminar: {inf['tag']} ({inf['tipo_plan']})"):
                if inf.get('ruta_prev_pdf') and os.path.exists(inf['ruta_prev_pdf']):
                    with open(inf['ruta_prev_pdf'], "rb") as f:
                        base64_pdf = base64.b64encode(f.read()).decode('utf-8')
                    
                    # Cambiamos iframe por embed, que es mucho más compatible con Edge y Chrome
                    pdf_display = f'<embed src="data:application/pdf;base64,{base64_pdf}" width="100%" height="600" type="application/pdf">'
                    st.markdown(pdf_display, unsafe_allow_html=True)
                    
                    with open(inf['ruta_prev_pdf'], "rb") as f2:
                        st.download_button("📥 Descargar Borrador (PDF)", f2, file_name=f"Borrador_{inf['tag']}.pdf", mime="application/pdf", key=f"dl_prev_{i}")
                else:
                    st.warning("⚠️ La vista preliminar en PDF no está disponible. Verifique la conexión con LibreOffice.")

        st.markdown("---")
        st.info("💡 **Instrucciones:** Dibuja las firmas en los recuadros usando el mouse o el dedo.")
        
        c_tec, c_cli = st.columns(2)
        with c_tec:
            st.markdown("### 🧑‍🔧 Firma del Técnico")
            st.caption(f"Técnico: {st.session_state.informes_pendientes[0]['tec1'] if st.session_state.informes_pendientes else 'N/A'}")
            canvas_tec = st_canvas(stroke_width=4, stroke_color="#000", background_color="#fff", height=200, width=400, drawing_mode="freedraw", key="canvas_tecnico")
            
        with c_cli:
            st.markdown("### 👷 Firma del Cliente")
            st.caption(f"Cliente: {st.session_state.informes_pendientes[0]['cli'] if st.session_state.informes_pendientes else 'N/A'}")
            canvas_cli = st_canvas(stroke_width=4, stroke_color="#000", background_color="#fff", height=200, width=400, drawing_mode="freedraw", key="canvas_cliente")

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🚀 Aprobar, Firmar y Subir a la Nube", type="primary", use_container_width=True):
            if canvas_tec.image_data is not None and canvas_cli.image_data is not None:
                def procesar_imagen_firma(img_data):
                    img = Image.fromarray(img_data.astype('uint8'), 'RGBA')
                    img_io = io.BytesIO()
                    img.save(img_io, format='PNG')
                    img_io.seek(0)
                    return img_io

                io_tec = procesar_imagen_firma(canvas_tec.image_data)
                io_cli = procesar_imagen_firma(canvas_cli.image_data)
                
                informes_finales = []
                with st.spinner("Fabricando documentos oficiales, inyectando firmas y transformando a PDF..."):
                    try:
                        for inf in st.session_state.informes_pendientes:
                            doc = DocxTemplate(inf['file_plantilla'])
                            context = inf['context']
                            
                            context['firma_tecnico'] = InlineImage(doc, io_tec, width=Mm(40))
                            context['firma_cliente'] = InlineImage(doc, io_cli, width=Mm(40))
                            
                            doc.render(context)
                            doc.save(inf['ruta_docx'])
                            
                            ruta_pdf_gen = convertir_a_pdf(inf['ruta_docx'])
                            
                            if ruta_pdf_gen:
                                ruta_final = ruta_pdf_gen
                                nombre_final = inf['nombre_archivo_base'].replace(".docx", ".pdf")
                            else:
                                ruta_final = inf['ruta_docx']
                                nombre_final = inf['nombre_archivo_base']
                            
                            nombre_codificado = f"{inf['area'].title()}@@{inf['tag']}@@{nombre_final}"
                            
                            # Guardamos en Google Sheets
                            tupla_lista = list(inf['tupla_db'])
                            tupla_lista[18] = ruta_final
                            guardar_registro(tuple(tupla_lista))
                            
                            informes_finales.append({"tag": inf['tag'], "tipo": inf['tipo_plan'], "ruta": ruta_final, "nombre_archivo": nombre_codificado})
                            
                        exito, mensaje_correo = enviar_carrito_por_correo(MI_CORREO_CORPORATIVO, informes_finales)
                        
                        if exito:
                            st.success("✅ ¡PERFECTO! Los documentos oficiales se firmaron, convirtieron a PDF y ya están camino a tu OneDrive.")
                            st.session_state.informes_pendientes = []  
                            st.balloons()
                        else:
                            st.error(f"Error de red: {mensaje_correo}")
                            
                    except Exception as e:
                        st.error(f"Error sistémico procesando las firmas: {e}")
            else:
                st.warning("⚠️ Asegúrate de dibujar en ambas pizarras antes de generar los PDFs finales.")

    # --- 6.2 VISTA CATÁLOGO (Dashboard interactivo) ---
    elif st.session_state.equipo_seleccionado is None:
        st.markdown("""
            <div style="margin-top: 1.5rem; margin-bottom: 2rem; text-align: center;">
                <div style="background-color: white; height: 2px; width: 100%;"></div>
                <h1 style="color: #007CA6; font-size: 4.5em; font-weight: 900; margin: 20px 0; border-bottom: none; padding: 0;">Atlas Copco</h1>
                <div style="background-color: white; height: 2px; width: 100%;"></div>
            </div>
        """, unsafe_allow_html=True)
        
        st.title("🏭 Panel de Control de Equipos")
        estados_db = obtener_estados_actuales()
        total_equipos = len(inventario_equipos)
        operativos = sum(1 for tag in inventario_equipos.keys() if estados_db.get(tag, "Operativo") == "Operativo")
        detenidos = total_equipos - operativos
        
        m1, m2, m3 = st.columns(3)
        m1.metric("📦 Total Activos Mineros", total_equipos)
        m2.metric("🟢 Equipos Operativos", operativos)
        m3.metric("🔴 Fuera de Servicio", detenidos)
        st.markdown("---")
        
        col_filtro, col_busqueda = st.columns([1.2, 2])
        with col_filtro: filtro_tipo = st.radio("🗂️ Categoría de Equipo:", ["Todos", "Compresores", "Secadores"], horizontal=True)
        with col_busqueda: busqueda = st.text_input("🔍 Buscar activo por TAG, Modelo o Área...", placeholder="Ejemplo: GA 250, 35-GC-006...").lower()
            
        st.markdown("<br>", unsafe_allow_html=True)
        columnas = st.columns(4)
        contador = 0
        for tag, (modelo, serie, area, ubicacion) in inventario_equipos.items():
            es_secador = "CD" in modelo.upper()
            if filtro_tipo == "Compresores" and es_secador: continue
            if filtro_tipo == "Secadores" and not es_secador: continue

            if busqueda in tag.lower() or busqueda in area.lower() or busqueda in modelo.lower():
                estado = estados_db.get(tag, "Operativo")
                color_bg = "#eaffea" if estado == "Operativo" else "#ffeaea"
                color_text = "#004d00" if estado == "Operativo" else "#800000"
                icono = "🟢" if estado == "Operativo" else "🔴"
                
                with columnas[contador % 4]:
                    with st.container(border=True):
                        st.markdown(f"<span style='background-color:{color_bg}; color:{color_text}; padding: 4px 8px; border-radius:4px; font-size:0.85em; font-weight:bold; letter-spacing: 0.5px;'>{icono} {estado.upper()}</span>", unsafe_allow_html=True)
                        st.markdown(f"<h3 style='margin-top:10px; margin-bottom:0;'>{tag}</h3>", unsafe_allow_html=True)
                        st.caption(f"**{modelo}** | {area.title()}")
                        st.button("📝 Ingresar", key=f"btn_{tag}", on_click=seleccionar_equipo, args=(tag,), use_container_width=True)
                contador += 1

    # --- 6.3 VISTA FORMULARIO Y GENERACIÓN ---
    else:
        tag_sel = st.session_state.equipo_seleccionado
        mod_d, ser_d, area_d, ubi_d = inventario_equipos[tag_sel]
        
        c_btn, c_tit = st.columns([1, 4])
        with c_btn: st.button("⬅️ Volver", on_click=volver_catalogo, use_container_width=True)
        with c_tit: st.markdown(f"<h1 style='margin-top:-15px;'>⚙️ Ficha de Servicio: <span style='color:#007CA6;'>{tag_sel}</span></h1>", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        tab1, tab2, tab3, tab4 = st.tabs(["📋 1. Reporte y Diagnóstico", "📚 2. Ficha Técnica", "🔍 3. Bitácora de Observaciones", "👤 4. Gestión de Área"])
        
        with tab1:
            st.markdown("### Datos de la Intervención")
            tipo_plan = st.selectbox("🛠️ Tipo de Plan / Orden:", ["Inspección", "PM03"] if "CD" in tag_sel else ["Inspección", "P1", "P2", "P3", "PM03"])
            
            c1, c2, c3, c4 = st.columns(4)
            modelo = c1.text_input("Modelo", mod_d, disabled=True)
            numero_serie = c2.text_input("N° Serie", ser_d, disabled=True)
            area = c3.text_input("Área", area_d, disabled=True)
            ubicacion = c4.text_input("Ubicación", ubi_d, disabled=True)

            c5, c6, c7, c8 = st.columns([1, 1, 1, 1.3])
            fecha = c5.text_input("Fecha Ejecución", "25 de febrero de 2026")
            tec1 = c6.text_input("Técnico 1", key="input_tec1")
            tec2 = c7.text_input("Técnico 2", key="input_tec2")
            
            with c8:
                contactos_db = obtener_contactos()
                opciones = ["➕ Escribir nuevo..."] + contactos_db
                
                if st.session_state.input_cliente in opciones: cli_idx = opciones.index(st.session_state.input_cliente)
                else: cli_idx = 1 if len(contactos_db) > 0 else 0
                
                sc1, sc2 = st.columns([4, 1])
                with sc1: cli_sel = st.selectbox("Contacto Cliente", opciones, index=cli_idx)
                with sc2:
                    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                    if cli_sel != "➕ Escribir nuevo...":
                        if st.button("❌", help="Eliminar permanentemente"):
                            eliminar_contacto(cli_sel)
                            st.session_state.input_cliente = obtener_contactos()[0] if obtener_contactos() else ""
                            st.rerun()
                
                if cli_sel == "➕ Escribir nuevo...":
                    nuevo_c = st.text_input("Nombre:", placeholder="Ej: Juan Pérez", label_visibility="collapsed")
                    if st.button("💾 Guardar y Seleccionar", use_container_width=True):
                        if nuevo_c.strip():
                            agregar_contacto(nuevo_c)
                            st.session_state.input_cliente = nuevo_c.strip().title()
                            st.rerun()
                    cli_cont = nuevo_c.strip().title()
                else:
                    cli_cont = cli_sel
                    st.session_state.input_cliente = cli_sel

            st.markdown("<hr>", unsafe_allow_html=True)
            st.markdown("### Mediciones del Equipo")
            c9, c10, c11, c12, c13, c14 = st.columns(6)
            h_m = c9.number_input("Horas Marcha Totales", step=1, value=int(st.session_state.input_h_marcha), format="%d")
            h_c = c10.number_input("Horas en Carga", step=1, value=int(st.session_state.input_h_carga), format="%d")
            unidad_p = c11.selectbox("Unidad de Presión", ["Bar", "psi"])
            p_c_str = c12.text_input("P. Carga", value=str(st.session_state.input_p_carga))
            p_d_str = c13.text_input("P. Descarga", value=str(st.session_state.input_p_descarga))
            t_salida_str = c14.text_input("Temp Salida (°C)", value=str(st.session_state.input_temp))
            
            p_c_clean = p_c_str.replace(',', '.')
            p_d_clean = p_d_str.replace(',', '.')
            t_salida_clean = t_salida_str.replace(',', '.')

            st.markdown("<hr>", unsafe_allow_html=True)
            st.markdown("### Evaluación y Diagnóstico Final")
            est_eq = st.radio("Estado de Devolución del Activo:", ["Operativo", "Fuera de servicio"], key="input_estado_eq", horizontal=True)
            est_ent = st.text_area("Descripción Condición Final:", key="input_estado", height=100)
            reco = st.text_area("Recomendaciones / Acciones Pendientes:", key="input_reco", height=100)
            
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("📥 Guardar y Añadir a la Bandeja de Firmas", type="primary", use_container_width=True):
                if "CD" in tag_sel: file_plantilla = "plantilla/secadorfueradeservicio.docx" if est_eq == "Fuera de servicio" else "plantilla/inspeccionsecador.docx"
                else:
                    if est_eq == "Fuera de servicio": file_plantilla = "plantilla/fueradeservicio.docx"
                    elif tipo_plan == "P1": file_plantilla = "plantilla/p1.docx"
                    elif tipo_plan == "P2": file_plantilla = "plantilla/p2.docx"
                    elif tipo_plan == "P3": file_plantilla = "plantilla/p3.docx"
                    else: file_plantilla = "plantilla/inspeccion.docx"
                    
                context = {"tipo_intervencion": tipo_plan, "modelo": mod_d, "tag": tag_sel, "area": area_d, "ubicacion": ubi_d, "cliente_contacto": cli_cont, "p_carga": f"{p_c_clean} {unidad_p}", "p_descarga": f"{p_d_clean} {unidad_p}", "temp_salida": t_salida_clean, "horas_marcha": int(h_m), "horas_carga": int(h_c), "tecnico_1": tec1, "tecnico_2": tec2, "estado_equipo": est_eq, "estado_entrega": est_ent, "recomendaciones": reco, "serie": ser_d, "tipo_orden": tipo_plan.upper(), "fecha": fecha, "equipo_modelo": mod_d}
                
                nombre_archivo = f"Informe_{tipo_plan}_{tag_sel}_{fecha.replace(' ','_')}.docx"
                ruta = os.path.join(RUTA_ONEDRIVE, nombre_archivo)
                
                try: temp_db = float(t_salida_clean)
                except: temp_db = 0.0
                tupla_db = (tag_sel, mod_d, ser_d, area_d, ubi_d, fecha, cli_cont, tec1, tec2, temp_db, f"{p_c_clean} {unidad_p}", f"{p_d_clean} {unidad_p}", h_m, h_c, est_ent, tipo_plan, reco, est_eq, "", st.session_state.usuario_actual)
                
                with st.spinner("Creando borrador del documento para vista preliminar..."):
                    doc_prev = DocxTemplate(file_plantilla)
                    ctx_prev = context.copy()
                    ctx_prev['firma_tecnico'] = "" 
                    ctx_prev['firma_cliente'] = ""
                    doc_prev.render(ctx_prev)
                    
                    os.makedirs(RUTA_ONEDRIVE, exist_ok=True)
                    ruta_prev_docx = os.path.join(RUTA_ONEDRIVE, f"PREVIEW_{nombre_archivo}")
                    doc_prev.save(ruta_prev_docx)
                    ruta_prev_pdf = convertir_a_pdf(ruta_prev_docx)
                
                st.session_state.informes_pendientes.append({
                    "tag": tag_sel, "area": area_d, "tec1": tec1, "cli": cli_cont, "tipo_plan": tipo_plan,
                    "file_plantilla": file_plantilla, "context": context, "tupla_db": tupla_db,
                    "ruta_docx": ruta, "nombre_archivo_base": nombre_archivo, "ruta_prev_pdf": ruta_prev_pdf
                })
                st.success("✅ Datos guardados. Agrega otro equipo o ve a la bandeja para firmar.")
                st.session_state.equipo_seleccionado = None
                st.rerun()

        with tab2:
            st.markdown(f"### 📘 Datos Técnicos y Repuestos ({mod_d})")
            with st.expander("✏️ Agregar o Corregir Datos Faltantes (N° Parte, Aceite, etc.)"):
                with st.form(key=f"form_specs_{tag_sel}"):
                    st.info(f"💡 Lo que guardes aquí se actualizará para todos los equipos modelo **{mod_d}**.")
                    c_e1, c_e2 = st.columns(2)
                    opc_claves = ["N° Parte Filtro Aceite", "N° Parte Filtro Aire", "N° Parte Kit", "N° Parte Separador", "Litros de Aceite", "Tipo de Aceite", "Cant. Filtros Aceite", "Cant. Filtros Aire", "Otro dato nuevo..."]
                    clave_sel = c_e1.selectbox("¿Qué dato vas a ingresar?", opc_claves)
                    if clave_sel == "Otro dato nuevo...": clave_final = c_e1.text_input("Escribe el nombre del dato:")
                    else: clave_final = clave_sel
                    valor_final = c_e2.text_input("Ingresa el valor (Ej: 1613 6105 00):")
                    
                    if st.form_submit_button("💾 Guardar en Base de Datos", use_container_width=True):
                        with st.spinner("Guardando en la nube..."):
                            if clave_final and valor_final:
                                guardar_especificacion_db(mod_d, clave_final.strip(), valor_final.strip())
                                st.success("✅ ¡Dato guardado en Google Sheets!")
                                st.rerun()
                            else: st.error("⚠️ Debes llenar el valor a guardar.")
            
            if mod_d in ESPECIFICACIONES:
                specs = {k: v for k, v in ESPECIFICACIONES[mod_d].items() if k != "Manual"}
                if specs:
                    cols = st.columns(3)
                    for i, (k, v) in enumerate(specs.items()):
                        with cols[i % 3]: st.markdown(f"<div style='background-color: #1e2530; padding: 15px; border-radius: 8px; margin-bottom: 15px; border-left: 4px solid #007CA6;'><span style='color: #8c9eb5; font-size: 0.85em; text-transform: uppercase; font-weight: bold;'>{k}</span><br><span style='color: white; font-size: 1.1em;'>{v}</span></div>", unsafe_allow_html=True)
                else: st.info("No hay datos técnicos registrados aún. ¡Sé el primero en agregar uno arriba!")
                    
                st.markdown("<hr>", unsafe_allow_html=True)
                st.markdown("### 📥 Documentación y Manuales")
                if "Manual" in ESPECIFICACIONES[mod_d] and os.path.exists(ESPECIFICACIONES[mod_d]["Manual"]):
                    with open(ESPECIFICACIONES[mod_d]["Manual"], "rb") as f: st.download_button(label=f"📕 Descargar Manual de {mod_d} (PDF)", data=f, file_name=ESPECIFICACIONES[mod_d]["Manual"].split('/')[-1], mime="application/pdf")
                else: st.info("ℹ️ El manual o despiece para este modelo aún no ha sido cargado en la plataforma.")
            else: st.warning(f"⚠️ No hay especificaciones técnicas base para el modelo {mod_d}.")

        with tab3:
            st.markdown(f"### 🔍 Bitácora Permanente del Equipo: {tag_sel}")
            st.info("💡 Usa este espacio para dejar notas importantes sobre el estado general del equipo.")
            
            with st.form(key=f"form_obs_{tag_sel}"):
                nueva_obs = st.text_area("Escribe una nueva observación o nota técnica para este equipo:", height=100)
                if st.form_submit_button("➕ Dejar constancia en la bitácora", use_container_width=True):
                    with st.spinner("Guardando en la nube..."):
                        if nueva_obs:
                            agregar_observacion(tag_sel, st.session_state.usuario_actual, nueva_obs)
                            st.success("✅ Observación registrada con éxito en Google Sheets.")
                            st.rerun()
                        else:
                            st.warning("⚠️ Debes escribir algo antes de guardar.")
            
            st.markdown("---")
            st.markdown("#### 📜 Historial de Observaciones Anteriores")
            df_obs = obtener_observaciones(tag_sel)
            
            if not df_obs.empty:
                for _, row in df_obs.iterrows():
                    col_obs, col_del = st.columns([11, 1])
                    with col_obs:
                        st.markdown(f"""
                            <div style='background-color: #2b303b; padding: 15px; border-radius: 8px; margin-bottom: 10px; border-left: 4px solid #FF6600;'>
                                <small style='color: #aeb9cc;'><b>👤 Técnico: {row['usuario']}</b> &nbsp;|&nbsp; 📅 Fecha: {row['fecha']}</small><br>
                                <span style='color: white; font-size: 1.05em;'>{row['texto']}</span>
                            </div>
                        """, unsafe_allow_html=True)
                    with col_del:
                        st.markdown("<div style='margin-top: 20px;'></div>", unsafe_allow_html=True)
                        if st.button("🗑️", key=f"del_obs_{row['id']}", help="Borrar esta nota"):
                            eliminar_observacion(row['id'])
                            st.rerun()
            else:
                st.caption("No hay observaciones registradas para este equipo aún. ¡Escribe la primera!")

        with tab4:
            st.markdown(f"### 👤 Información de Contactos y Seguridad del Área: {tag_sel}")
            st.info("💡 Asigna y actualiza los dueños del área, el PEA y la frecuencia radial correspondientes a este equipo.")

            with st.expander("✏️ Editar o Agregar Contacto / Dato de Seguridad"):
                with st.form(key=f"form_area_{tag_sel}"):
                    c_a1, c_a2 = st.columns(2)
                    opc_area = ["Dueño de Área (Turno 1-3)", "Dueño de Área (Turno 2-4)", "PEA", "Frecuencia Radial", "Supervisor a cargo", "Jefe de Turno", "Otro cargo..."]
                    clave_sel_area = c_a1.selectbox("¿Qué dato vas a ingresar?", opc_area)
                    if clave_sel_area == "Otro cargo...": clave_final_area = c_a1.text_input("Escribe el nombre del cargo/dato:")
                    else: clave_final_area = clave_sel_area
                    valor_final_area = c_a2.text_input("Ingresa la información:")
                    
                    if st.form_submit_button("💾 Guardar Información", use_container_width=True):
                        with st.spinner("Guardando en la nube..."):
                            if clave_final_area and valor_final_area:
                                guardar_dato_equipo(tag_sel, clave_final_area.strip(), valor_final_area.strip())
                                st.success("✅ ¡Dato actualizado exitosamente en Google Sheets!")
                                st.rerun()
                            else: st.error("⚠️ Debes llenar ambos campos.")

            datos_equipo = obtener_datos_equipo(tag_sel)
            if "Dueño de Área (Turno 1-3)" not in datos_equipo: datos_equipo["Dueño de Área (Turno 1-3)"] = "No asignado"
            if "Dueño de Área (Turno 2-4)" not in datos_equipo: datos_equipo["Dueño de Área (Turno 2-4)"] = "No asignado"
            if "PEA" not in datos_equipo: datos_equipo["PEA"] = "No asignado"
            if "Frecuencia Radial" not in datos_equipo: datos_equipo["Frecuencia Radial"] = "No asignada"

            cols_area = st.columns(2)
            for i, (k, v) in enumerate(datos_equipo.items()):
                with cols_area[i % 2]:
                    st.markdown(f"""
                        <div style='background-color: #2b303b; padding: 15px; border-radius: 8px; margin-bottom: 15px; border-left: 4px solid #FF6600;'>
                            <span style='color: #aeb9cc; font-size: 0.85em; text-transform: uppercase; font-weight: bold;'>{k}</span><br>
                            <span style='color: white; font-size: 1.1em;'>{v}</span>
                        </div>
                    """, unsafe_allow_html=True)

        st.markdown("<br><hr>", unsafe_allow_html=True)
        st.markdown("### 📋 Trazabilidad Histórica de Intervenciones (Reportes Creados)")
        df_hist = obtener_todo_el_historial(tag_sel)
        if not df_hist.empty: st.dataframe(df_hist, use_container_width=True)