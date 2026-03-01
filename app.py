import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import os, subprocess, smtplib, time, json, uuid
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from streamlit_drawable_canvas import st_canvas
from PIL import Image
import io
import gspread
from google.oauth2.service_account import Credentials
from streamlit_pdf_viewer import pdf_viewer

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
    for item in lista_informes: cuerpo += f"- TAG: {item['tag']} | Orden: {item['tipo']}\n"
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
    except Exception as e: return False, f"❌ Error al enviar el correo: {e}"

# =============================================================================
# 0.2 ESTILOS PREMIUM (Corregido y Limpio)
# =============================================================================
st.set_page_config(page_title="Atlas Spence | Gestión de Reportes", layout="wide", page_icon="⚙️")

def aplicar_estilos_premium():
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;600;800&display=swap');
        
        :root { --ac-blue: #007CA6; --ac-dark: #005675; --bhp-orange: #FF6600; --bg-card: #151a22; }
        
        html, body, [class*="css"] { font-family: 'Montserrat', sans-serif !important; }
        
        #MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}
        
        div.stButton > button:first-child {
            background: linear-gradient(135deg, var(--ac-blue) 0%, var(--ac-dark) 100%);
            color: white; border-radius: 8px; border: none; font-weight: 600; letter-spacing: 0.5px;
            padding: 0.6rem 1.2rem; transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1);
        }
        
        div.stButton > button:first-child:hover { 
            transform: translateY(-3px); box-shadow: 0 8px 25px rgba(0, 124, 166, 0.6); 
        }

        div[data-testid="column"] div.stButton > button {
            background: linear-gradient(145deg, #1a212b, #151a22) !important;
            border: 1px solid #2b3543 !important;
            border-radius: 12px !important;
            color: white !important;
            font-size: 1.4rem !important;
            font-weight: 800 !important;
            padding: 1.5rem 1rem !important;
            box-shadow: 0 4px 10px rgba(0,0,0,0.2) !important;
            transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1) !important;
            width: 100% !important;
        }
        
        div[data-testid="column"] div.stButton > button:hover {
            transform: translateY(-6px) !important;
            border-color: var(--ac-blue) !important;
            box-shadow: 0 12px 25px rgba(0, 124, 166, 0.3) !important;
            background: #1e2530 !important;
        }
        
        .stTextInput>div>div>input, .stNumberInput>div>div>input, .stSelectbox>div>div>select { 
            border-radius: 6px !important; border: 1px solid #2b3543 !important; 
            background-color: #1e2530 !important; color: white !important; transition: all 0.3s ease;
        }
        
        .stTextInput>div>div>input:focus, .stNumberInput>div>div>input:focus, .stSelectbox>div>div>select:focus { 
            border-color: var(--bhp-orange) !important; box-shadow: 0 0 10px rgba(255, 102, 0, 0.3) !important; 
        }
        
        .stTabs [data-baseweb="tab-list"] { gap: 10px; border-bottom: 2px solid #2b3543; padding-bottom: 5px; }
        .stTabs [data-baseweb="tab"] { background-color: transparent; border: none; color: #8c9eb5; font-weight: 600; }
        .stTabs [aria-selected="true"] { color: var(--bhp-orange) !important; border-bottom: 3px solid var(--bhp-orange) !important; }
        </style>
    """, unsafe_allow_html=True)

# =============================================================================
# 1. DATOS MAESTROS (INVENTARIO Y USUARIOS)
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
    "80-GC-001": ["GA 18", "API335343", "Taller", "Taller"]
}

# =============================================================================
# 2. CONEXIÓN OPTIMIZADA A GOOGLE SHEETS
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
        try: return doc.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound: return doc.add_worksheet(title=sheet_name, rows="1000", cols="20")
    except Exception as e: return None

@st.cache_data(ttl=120)
def obtener_datos_equipo(tag):
    datos = {}
    try:
        sheet = get_sheet("datos_equipo")
        if sheet:
            data = sheet.get_all_values()
            for row in data:
                if len(row) >= 3 and row[0] == tag: datos[row[1]] = row[2]
    except: pass
    return datos
@st.cache_data(ttl=120)
def obtener_observaciones(tag):
    try:
        sheet = get_sheet("observaciones")
        if sheet:
            data = sheet.get_all_values()
            obs = [{"id": r[0], "fecha": r[2], "usuario": r[3], "texto": r[4]} for r in data if len(r) >= 6 and r[1] == tag and r[5] == "ACTIVO"]
            if obs: return pd.DataFrame(obs).iloc[::-1]
    except: pass
    return pd.DataFrame(columns=["id", "fecha", "usuario", "texto"])
@st.cache_data(ttl=120)
def obtener_especificaciones(defaults):
    specs = {k: dict(v) for k, v in defaults.items()}
    try:
        sheet = get_sheet("especificaciones")
        if sheet:
            data = sheet.get_all_values()
            for row in data:
                if len(row) >= 3:
                    mod, clave, valor = row[0], row[1], row[2]
                    if mod not in specs: specs[mod] = {}
                    specs[mod][clave] = valor
    except: pass
    return specs
@st.cache_data(ttl=120)
def obtener_contactos():
    try:
        sheet = get_sheet("contactos")
        if sheet:
            data = sheet.get_all_values()
            contactos = [row[0] for row in data if len(row) > 1 and row[1] == "ACTIVO"]
            if contactos: return sorted(list(set(contactos)))
    except: pass
    return ["Lorena Rojas"]
@st.cache_data(ttl=120)
def buscar_ultimo_registro(tag):
    try:
        sheet = get_sheet("intervenciones")
        if sheet:
            data = sheet.get_all_values()
            for row in reversed(data):
                if len(row) >= 20 and row[0] == tag: return (row[5], row[6], row[9], row[14], row[15], row[7], row[8], row[10], row[11], row[12], row[13], row[16], row[17])
    except: pass
    return None
@st.cache_data(ttl=120)
def obtener_todo_el_historial(tag):
    try:
        sheet = get_sheet("intervenciones")
        if sheet:
            data = sheet.get_all_values()
            hist = [{"fecha": r[5], "tipo_intervencion": r[15], "estado_equipo": r[17], "Cuenta Usuario": r[19], "horas_marcha": r[12], "p_carga": r[10], "temp_salida": r[9]} for r in data if len(r) >= 20 and r[0] == tag]
            if hist: return pd.DataFrame(hist).iloc[::-1]
    except: pass
    return pd.DataFrame()
@st.cache_data(ttl=120)
def obtener_estados_actuales():
    estados = {}
    try:
        sheet = get_sheet("intervenciones")
        if sheet:
            data = sheet.get_all_values()
            for row in data:
                if len(row) >= 18: estados[row[0]] = row[17]
    except: pass
    return estados

def guardar_dato_equipo(tag, clave, valor):
    try:
        sheet = get_sheet("datos_equipo")
        if sheet: sheet.append_row([tag, clave, valor]); st.cache_data.clear()
    except: pass
def agregar_observacion(tag, usuario, texto):
    if not texto.strip(): return
    fecha_actual = pd.Timestamp.now().strftime("%d/%m/%Y %H:%M")
    id_obs = str(uuid.uuid4())[:8]
    try:
        sheet = get_sheet("observaciones")
        if sheet: sheet.append_row([id_obs, tag, fecha_actual, usuario.title(), texto.strip(), "ACTIVO"]); st.cache_data.clear()
    except: pass
def eliminar_observacion(id_obs):
    try:
        sheet = get_sheet("observaciones")
        if sheet:
            cell = sheet.find(id_obs)
            if cell: sheet.update_cell(cell.row, 6, "ELIMINADO"); st.cache_data.clear()
    except: pass
def guardar_especificacion_db(modelo, clave, valor):
    try:
        sheet = get_sheet("especificaciones")
        if sheet: sheet.append_row([modelo, clave, valor]); st.cache_data.clear()
    except: pass
def agregar_contacto(nombre):
    if not nombre.strip(): return
    try:
        sheet = get_sheet("contactos")
        if sheet: sheet.append_row([nombre.strip().title(), "ACTIVO"]); st.cache_data.clear()
    except: pass
def eliminar_contacto(nombre):
    try:
        sheet = get_sheet("contactos")
        if sheet:
            cells = sheet.findall(nombre)
            for cell in cells: sheet.update_cell(cell.row, 2, "ELIMINADO"); st.cache_data.clear()
    except: pass
def guardar_registro(data_tuple):
    for intento in range(3):
        try:
            sheet = get_sheet("intervenciones")
            if sheet is not None:
                row = [str(x) for x in data_tuple]
                num_filas = len(sheet.get_all_values()) + 1
                sheet.insert_row(row, index=num_filas)
                st.cache_data.clear(); return True
        except Exception as e: time.sleep(5)
    return False
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
        pythoncom.CoInitialize(); convert(ruta_absoluta, ruta_pdf)
        if os.path.exists(ruta_pdf): return ruta_pdf
    except: pass
    return None

# =============================================================================
# 4. INICIALIZACIÓN DE VARIABLES DE SESIÓN
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
    st.session_state.equipo_seleccionado = tag; st.session_state.vista_firmas = False
    reg = buscar_ultimo_registro(tag)
    if reg:
        st.session_state.input_cliente = reg[1]
        st.session_state.input_tec1 = reg[5]; st.session_state.input_tec2 = reg[6]
        st.session_state.input_estado = reg[3]
        st.session_state.input_reco = reg[11] if reg[11] else ""
        st.session_state.input_estado_eq = reg[12] if reg[12] else "Operativo"
        st.session_state.input_h_marcha = int(reg[9]) if reg[9] else 0; st.session_state.input_h_carga = int(reg[10]) if reg[10] else 0
        st.session_state.input_temp = str(reg[2]).replace(',', '.') if reg[2] is not None else "70.0"
        try: st.session_state.input_p_carga = str(reg[7]).split()[0].replace(',', '.')
        except: st.session_state.input_p_carga = "7.0"
        try: st.session_state.input_p_descarga = str(reg[8]).split()[0].replace(',', '.')
        except: st.session_state.input_p_descarga = "7.5"
    else: st.session_state.update({'input_estado_eq': "Operativo", 'input_estado': "", 'input_reco': ""})
def volver_catalogo(): 
    st.session_state.equipo_seleccionado = None; st.session_state.vista_firmas = False

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
                    if u_in in USUARIOS and USUARIOS[u_in] == p_in: st.session_state.update({'logged_in': True, 'usuario_actual': u_in}); st.rerun()
                    else: st.error("❌ Credenciales inválidas.")

# =============================================================================
# 6. PANTALLA PRINCIPAL: APLICACIÓN AUTENTICADA
# =============================================================================
else:
    with st.sidebar:
        st.markdown("<h2 style='text-align: center; border-bottom:none; margin-top: -20px;'><span style='color:#007CA6;'>Atlas Copco</span> <span style='color:#FF6600;'>Spence</span></h2>", unsafe_allow_html=True)
        st.markdown(f"**Usuario Activo:**<br>{st.session_state.usuario_actual.title()}", unsafe_allow_html=True)
        if len(st.session_state.informes_pendientes) > 0:
            st.markdown("---")
            st.warning(f"📝 Tienes {len(st.session_state.informes_pendientes)} reportes esperando firmas.")
            if st.button("✍️ Ir a Pizarra de Firmas", use_container_width=True, type="primary"): st.session_state.vista_firmas = True; st.session_state.equipo_seleccionado = None; st.rerun()
        st.markdown("---")
        if st.button("🚪 Cerrar Sesión", use_container_width=True): st.session_state.logged_in = False; st.rerun()

    # --- 6.1 VISTA DE FIRMAS Y ENVÍO MÚLTIPLE ---
    if st.session_state.vista_firmas:
        c_v1, c_v2 = st.columns([1,4])
        with c_v1: 
            if st.button("⬅️ Volver", use_container_width=True): volver_catalogo(); st.rerun()
        with c_v2: st.markdown("<h1 style='margin-top:-15px;'>✍️ Pizarra de Firmas Digital</h1>", unsafe_allow_html=True)
        st.markdown("---"); st.markdown(f"### 📑 Revisión de Informes ({len(st.session_state.informes_pendientes)})"); st.info("👀 **Para el Cliente:** Por favor, revise el documento oficial antes de firmar.")
        
        for i, inf in enumerate(st.session_state.informes_pendientes):
            c_exp, c_del = st.columns([12, 1])
            with c_exp:
                with st.expander(f"📄 Ver documento preliminar: {inf['tag']} ({inf['tipo_plan']})"):
                    if inf.get('ruta_prev_pdf') and os.path.exists(inf['ruta_prev_pdf']):
                        try: pdf_viewer(inf['ruta_prev_pdf'], width=700, height=600)
                        except Exception as e: st.error(f"No se pudo desplegar el visor: {e}")
                        
                        st.markdown("<br>", unsafe_allow_html=True)
                        # DOS BOTONES: PDF y Word
                        col_dl1, col_dl2 = st.columns(2)
                        with col_dl1:
                            with open(inf['ruta_prev_pdf'], "rb") as f2: 
                                st.download_button("📥 Descargar Borrador (PDF)", f2, file_name=f"Borrador_{inf['tag']}.pdf", mime="application/pdf", key=f"dl_prev_pdf_{i}", use_container_width=True)
                        with col_dl2:
                            ruta_prev_docx = os.path.join(RUTA_ONEDRIVE, f"PREVIEW_{inf['nombre_archivo_base']}")
                            if os.path.exists(ruta_prev_docx):
                                with open(ruta_prev_docx, "rb") as f3: 
                                    st.download_button("📝 Descargar Borrador (Word)", f3, file_name=f"Borrador_{inf['tag']}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"dl_prev_docx_{i}", use_container_width=True)
                    else: st.warning("⚠️ La vista preliminar no está disponible.")
            with c_del:
                st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)
                if st.button("❌", key=f"del_inf_{i}", help="Quitar este informe de la bandeja"):
                    st.session_state.informes_pendientes.pop(i)
                    if len(st.session_state.informes_pendientes) == 0: st.session_state.vista_firmas = False; st.session_state.equipo_seleccionado = None
                    st.rerun()
                    
        st.markdown("---"); st.info("💡 **Instrucciones:** Dibuja las firmas en los recuadros usando el mouse o el dedo.")
        c_tec, c_cli = st.columns(2)
        with c_tec:
            st.markdown("### 🧑‍🔧 Firma del Técnico"); st.caption(f"Técnico: {st.session_state.informes_pendientes[0]['tec1'] if st.session_state.informes_pendientes else 'N/A'}")
            canvas_tec = st_canvas(stroke_width=4, stroke_color="#000", background_color="#fff", height=200, width=400, drawing_mode="freedraw", key="canvas_tecnico")
        with c_cli:
            st.markdown("### 👷 Firma del Cliente"); st.caption(f"Cliente: {st.session_state.informes_pendientes[0]['cli'] if st.session_state.informes_pendientes else 'N/A'}")
            canvas_cli = st_canvas(stroke_width=4, stroke_color="#000", background_color="#fff", height=200, width=400, drawing_mode="freedraw", key="canvas_cliente")
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🚀 Aprobar, Firmar y Subir a la Nube", type="primary", use_container_width=True):
            if canvas_tec.image_data is not None and canvas_cli.image_data is not None:
                def procesar_imagen_firma(img_data):
                    img = Image.fromarray(img_data.astype('uint8'), 'RGBA'); img_io = io.BytesIO(); img.save(img_io, format='PNG'); img_io.seek(0); return img_io
                io_tec = procesar_imagen_firma(canvas_tec.image_data); io_cli = procesar_imagen_firma(canvas_cli.image_data); informes_finales = []
                with st.spinner("Fabricando documentos oficiales, inyectando firmas y transformando a PDF..."):
                    try:
                        for inf in st.session_state.informes_pendientes:
                            doc = DocxTemplate(inf['file_plantilla']); context = inf['context']
                            context['firma_tecnico'] = InlineImage(doc, io_tec, width=Mm(40)); context['firma_cliente'] = InlineImage(doc, io_cli, width=Mm(40)); doc.render(context); doc.save(inf['ruta_docx']); ruta_pdf_gen = convertir_a_pdf(inf['ruta_docx'])
                            if ruta_pdf_gen: ruta_final = ruta_pdf_gen; nombre_final = inf['nombre_archivo_base'].replace(".docx", ".pdf")
                            else: ruta_final = inf['ruta_docx']; nombre_final = inf['nombre_archivo_base']
                            tupla_lista = list(inf['tupla_db']); tupla_lista[18] = ruta_final; guardado_ok = guardar_registro(tuple(tupla_lista))
                            if not guardado_ok: st.error(f"⚠️ El PDF de {inf['tag']} se generó y envió, pero la base de datos de Google superó su límite. Verifica el catálogo en 1 minuto.")
                            informes_finales.append({"tag": inf['tag'], "tipo": inf['tipo_plan'], "ruta": ruta_final, "nombre_archivo": f"{inf['area'].title()}@@{inf['tag']}@@{nombre_final}"})
                        exito, mensaje_correo = enviar_carrito_por_correo(MI_CORREO_CORPORATIVO, informes_finales)
                        if exito: st.success("✅ ¡PERFECTO! Los documentos oficiales se firmaron, convirtieron a PDF y ya están camino a tu OneDrive."); st.session_state.informes_pendientes = []; st.balloons()
                        else: st.error(f"Error de red: {mensaje_correo}")
                    except Exception as e: st.error(f"Error sistémico procesando las firmas: {e}")
            else: st.warning("⚠️ Asegúrate de dibujar en ambas pizarras antes de generar los PDFs finales.")

    # --- 6.2 VISTA CATÁLOGO (DASHBOARD CINETICO Y PREMIUM) ---
    elif st.session_state.equipo_seleccionado is None:
        st.markdown("""
            <div style="margin-top: 1rem; margin-bottom: 2.5rem; text-align: center; background: linear-gradient(90deg, rgba(0,124,166,0) 0%, rgba(0,124,166,0.1) 50%, rgba(0,124,166,0) 100%); padding: 20px; border-radius: 15px;">
                <h1 style="color: #007CA6; font-size: 4em; font-weight: 800; margin: 0; letter-spacing: -1px; text-transform: uppercase;">Atlas Copco <span style="color: #FF6600;">Spence</span></h1>
                <p style="color: #8c9eb5; font-size: 1.2em; font-weight: 300; margin-top: -10px;">Sistema Integrado de Control de Activos • Hidrometalurgia</p>
            </div>
        """, unsafe_allow_html=True)
        estados_db = obtener_estados_actuales(); total_equipos = len(inventario_equipos); operativos = sum(1 for tag in inventario_equipos.keys() if estados_db.get(tag, "Operativo") == "Operativo"); fuera_servicio = total_equipos - operativos
        
        m1, m2, m3 = st.columns(3)
        with m1: st.markdown(f"<div style='background: #1e2530; border-left: 5px solid #8c9eb5; padding: 20px; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.2); text-align: center;'><p style='color: #8c9eb5; margin:0; font-size:1rem; font-weight:600; text-transform:uppercase;'>📦 Total Activos</p><h2 style='color: white; margin:0; font-size:2.5rem; font-weight:800;'>{total_equipos}</h2></div>", unsafe_allow_html=True)
        with m2: st.markdown(f"<div style='background: #1e2530; border-left: 5px solid #00e676; padding: 20px; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,230,118,0.1); text-align: center;'><p style='color: #8c9eb5; margin:0; font-size:1rem; font-weight:600; text-transform:uppercase;'>🟢 Operativos</p><h2 style='color: #00e676; margin:0; font-size:2.5rem; font-weight:800;'>{operativos}</h2></div>", unsafe_allow_html=True)
        with m3: st.markdown(f"<div style='background: #1e2530; border-left: 5px solid #ff1744; padding: 20px; border-radius: 10px; box-shadow: 0 4px 15px rgba(255,23,68,0.1); text-align: center;'><p style='color: #8c9eb5; margin:0; font-size:1rem; font-weight:600; text-transform:uppercase;'>🔴 Fuera de Servicio</p><h2 style='color: #ff1744; margin:0; font-size:2.5rem; font-weight:800;'>{fuera_servicio}</h2></div>", unsafe_allow_html=True)
        st.markdown("<br><hr style='border-color: #2b3543;'>", unsafe_allow_html=True)
        
        col_filtro, col_busqueda = st.columns([1.2, 2])
        with col_filtro: filtro_tipo = st.radio("🗂️ Categoría de Equipo:", ["Todos", "Compresores", "Secadores"], horizontal=True)
        with col_busqueda: busqueda = st.text_input("🔍 Buscar activo por TAG, Modelo o Área...", placeholder="Ejemplo: GA 250, 35-GC-006...").lower()
        st.markdown("<br>", unsafe_allow_html=True); columnas = st.columns(4); contador = 0
        
        for tag, (modelo, serie, area, ubicacion) in inventario_equipos.items():
            es_secador = "CD" in modelo.upper()
            if filtro_tipo == "Compresores" and es_secador: continue
            if filtro_tipo == "Secadores" and not es_secador: continue
            if busqueda in tag.lower() or busqueda in area.lower() or busqueda in modelo.lower():
                estado = estados_db.get(tag, "Operativo")
                if estado == "Operativo":
                    color_borde = "#00e676"; badge_html = "<span style='background: rgba(0,230,118,0.15); color: #00e676; border: 1px solid #00e676; padding: 4px 10px; border-radius: 20px; font-size: 0.75rem; font-weight: 700; box-shadow: 0 0 10px rgba(0,230,118,0.4); text-transform: uppercase; letter-spacing: 1px;'>OPERATIVO</span>"
                else:
                    color_borde = "#ff1744"; badge_html = "<span style='background: rgba(255,23,68,0.15); color: #ff1744; border: 1px solid #ff1744; padding: 4px 10px; border-radius: 20px; font-size: 0.75rem; font-weight: 700; box-shadow: 0 0 10px rgba(255,23,68,0.4); text-transform: uppercase; letter-spacing: 1px;'>FUERA DE SERVICIO</span>"
                
                with columnas[contador % 4]:
                    st.markdown(f"""
                    <div class="tarjeta-equipo" style="border-top: 4px solid {color_borde};">
                        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px;">{badge_html}</div>
                        <h3 style="color: white; margin: 0 0 5px 0; font-size: 1.4rem; font-weight: 800;">{tag}</h3>
                        <p style="color: #8c9eb5; margin: 0; font-size: 0.9rem; font-weight: 400;"><strong style="color:#007CA6;">{modelo}</strong> &bull; {area.title()}</p>
                    </div>""", unsafe_allow_html=True)
                    st.button(" ", key=f"btn_{tag}", on_click=seleccionar_equipo, args=(tag,), use_container_width=True)
                contador += 1

    # --- 6.3 VISTA FORMULARIO Y GENERACIÓN ---
    else:
        tag_sel = st.session_state.equipo_seleccionado; mod_d, ser_d, area_d, ubi_d = inventario_equipos[tag_sel]
        c_btn, c_tit = st.columns([1, 4])
        with c_btn: st.button("⬅️ Volver", on_click=volver_catalogo, use_container_width=True)
        with c_tit: st.markdown(f"<h1 style='margin-top:-15px;'>⚙️ Ficha de Servicio: <span style='color:#007CA6;'>{tag_sel}</span></h1>", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True); tab1, tab2, tab3, tab4 = st.tabs(["📋 1. Reporte y Diagnóstico", "📚 2. Ficha Técnica", "🔍 3. Bitácora de Observaciones", "👤 4. Gestión de Área"])
        with tab1:
            st.markdown("### Datos de la Intervención"); tipo_plan = st.selectbox("🛠️ Tipo de Plan / Orden:", ["Inspección", "PM03"] if "CD" in tag_sel else ["Inspección", "P1", "P2", "P3", "PM03"]); c1, c2, c3, c4 = st.columns(4); modelo = c1.text_input("Modelo", mod_d, disabled=True); numero_serie = c2.text_input("N° Serie", ser_d, disabled=True); area = c3.text_input("Área", area_d, disabled=True); ubicacion = c4.text_input("Ubicación", ubi_d, disabled=True); c5, c6, c7, c8 = st.columns([1, 1, 1, 1.3]); fecha = c5.text_input("Fecha Ejecución", pd.Timestamp.now().strftime("%d de %B de %Y")); tec1 = c6.text_input("Técnico 1", key="input_tec1"); tec2 = c7.text_input("Técnico 2", key="input_tec2")
            with c8:
                contactos_db = obtener_contactos(); opciones = ["➕ Escribir nuevo..."] + contactos_db
                cli_idx = opciones.index(st.session_state.input_cliente) if st.session_state.input_cliente in opciones else 1 if len(contactos_db) > 0 else 0
                sc1, sc2 = st.columns([4, 1])
                with sc1: cli_sel = st.selectbox("Contacto Cliente", opciones, index=cli_idx)
                with sc2:
                    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                    if cli_sel != "➕ Escribir nuevo...":
                        if st.button("❌", help="Eliminar permanentemente"): eliminar_contacto(cli_sel); st.session_state.input_cliente = obtener_contactos()[0] if obtener_contactos() else ""; st.rerun()
                if cli_sel == "➕ Escribir nuevo...":
                    nuevo_c = st.text_input("Nombre:", placeholder="Ej: Juan Pérez", label_visibility="collapsed")
                    if st.button("💾 Guardar y Seleccionar", use_container_width=True):
                        if nuevo_c.strip(): agregar_contacto(nuevo_c); st.session_state.input_cliente = nuevo_c.strip().title(); st.rerun()
                    cli_cont = nuevo_c.strip().title()
                else: cli_cont = cli_sel; st.session_state.input_cliente = cli_sel
            st.markdown("<hr>", unsafe_allow_html=True); st.markdown("### Mediciones del Equipo"); c9, c10, c11, c12, c13, c14 = st.columns(6); h_m = c9.number_input("Horas Marcha Totales", step=1, value=int(st.session_state.input_h_marcha), format="%d"); h_c = c10.number_input("Horas en Carga", step=1, value=int(st.session_state.input_h_carga), format="%d"); unidad_p = c11.selectbox("Unidad de Presión", ["Bar", "psi"]); p_c_str = c12.text_input("P. Carga", value=str(st.session_state.input_p_carga)); p_d_str = c13.text_input("P. Descarga", value=str(st.session_state.input_p_descarga)); t_salida_str = c14.text_input("Temp Salida (°C)", value=str(st.session_state.input_temp)); p_c_clean = p_c_str.replace(',', '.'); p_d_clean = p_d_str.replace(',', '.'); t_salida_clean = t_salida_str.replace(',', '.')
            st.markdown("<hr>", unsafe_allow_html=True); st.markdown("### Evaluación y Diagnóstico Final"); est_eq = st.radio("Estado de Devolución del Activo:", ["Operativo", "Fuera de servicio"], key="input_estado_eq", horizontal=True); est_ent = st.text_area("Descripción Condición Final:", key="input_estado", height=100); reco = st.text_area("Recomendaciones / Acciones Pendientes:", key="input_reco", height=100); st.markdown("<br>", unsafe_allow_html=True)
            if st.button("📥 Guardar y Añadir a la Bandeja de Firmas", type="primary", use_container_width=True):
                if "CD" in tag_sel: file_plantilla = "plantilla/secadorfueradeservicio.docx" if est_eq == "Fuera de servicio" else "plantilla/inspeccionsecador.docx"
                else: file_plantilla = "plantilla/fueradeservicio.docx" if est_eq == "Fuera de servicio" else f"plantilla/{tipo_plan.lower()}.docx" if tipo_plan in ["P1", "P2", "P3"] else "plantilla/inspeccion.docx"
                context = {"tipo_intervencion": tipo_plan, "modelo": mod_d, "tag": tag_sel, "area": area_d, "ubicacion": ubi_d, "cliente_contacto": cli_cont, "p_carga": f"{p_c_clean} {unidad_p}", "p_descarga": f"{p_d_clean} {unidad_p}", "temp_salida": t_salida_clean, "horas_marcha": int(h_m), "horas_carga": int(h_c), "tecnico_1": tec1, "tecnico_2": tec2, "estado_equipo": est_eq, "estado_entrega": est_ent, "recomendaciones": reco, "serie": ser_d, "tipo_orden": tipo_plan.upper(), "fecha": fecha, "equipo_modelo": mod_d}; nombre_archivo = f"Informe_{tipo_plan}_{tag_sel}_{fecha.replace(' ','_')}.docx"; ruta = os.path.join(RUTA_ONEDRIVE, nombre_archivo); temp_db = float(t_salida_clean) if t_salida_clean.replace('.', '', 1).isdigit() else 0.0; tupla_db = (tag_sel, mod_d, ser_d, area_d, ubi_d, fecha, cli_cont, tec1, tec2, temp_db, f"{p_c_clean} {unidad_p}", f"{p_d_clean} {unidad_p}", h_m, h_c, est_ent, tipo_plan, reco, est_eq, "", st.session_state.usuario_actual)
                with st.spinner("Creando borrador del documento para vista preliminar..."):
                    doc_prev = DocxTemplate(file_plantilla); ctx_prev = context.copy(); ctx_prev['firma_tecnico'] = ""; ctx_prev['firma_cliente'] = ""; doc_prev.render(ctx_prev); os.makedirs(RUTA_ONEDRIVE, exist_ok=True); ruta_prev_docx = os.path.join(RUTA_ONEDRIVE, f"PREVIEW_{nombre_archivo}"); doc_prev.save(ruta_prev_docx); ruta_prev_pdf = convertir_a_pdf(ruta_prev_docx)
                st.session_state.informes_pendientes.append({"tag": tag_sel, "area": area_d, "tec1": tec1, "cli": cli_cont, "tipo_plan": tipo_plan, "file_plantilla": file_plantilla, "context": context, "tupla_db": tupla_db, "ruta_docx": ruta, "nombre_archivo_base": nombre_archivo, "ruta_prev_pdf": ruta_prev_pdf}); st.success("✅ Datos guardados. Agrega otro equipo o ve a la bandeja para firmar."); st.session_state.equipo_seleccionado = None; st.rerun()
        with tab2:
            st.markdown(f"### 📘 Datos Técnicos y Repuestos ({mod_d})")
            with st.expander("✏️ Agregar o Corregir Datos Faltantes"):
                with st.form(key=f"form_specs_{tag_sel}"):
                    c_e1, c_e2 = st.columns(2); opc_claves = ["N° Parte Filtro Aceite", "N° Parte Filtro Aire", "N° Parte Kit", "N° Parte Separador", "Litros de Aceite", "Tipo de Aceite", "Cant. Filtros Aceite", "Cant. Filtros Aire", "Otro dato nuevo..."]; clave_sel = c_e1.selectbox("¿Qué dato vas a ingresar?", opc_claves); clave_final = c_e1.text_input("Escribe el nombre del dato:") if clave_sel == "Otro dato nuevo..." else clave_sel; valor_final = c_e2.text_input("Ingresa el valor:")
                    if st.form_submit_button("💾 Guardar en Base de Datos", use_container_width=True):
                        if clave_final and valor_final: guardar_especificacion_db(mod_d, clave_final.strip(), valor_final.strip()); st.success("✅ ¡Dato guardado!"); st.rerun()
            if mod_d in ESPECIFICACIONES:
                specs = {k: v for k, v in ESPECIFICACIONES[mod_d].items() if k != "Manual"}
                if specs:
                    cols = st.columns(3)
                    for i, (k, v) in enumerate(specs.items()):
                        with cols[i % 3]: st.markdown(f"<div style='background-color: #1e2530; padding: 15px; border-radius: 8px; margin-bottom: 15px; border-left: 4px solid #007CA6;'><span style='color: #8c9eb5; font-size: 0.85em; text-transform: uppercase; font-weight: bold;'>{k}</span><br><span style='color: white; font-size: 1.1em;'>{v}</span></div>", unsafe_allow_html=True)
                st.markdown("<hr>", unsafe_allow_html=True); st.markdown("### 📥 Documentación y Manuales")
                if "Manual" in ESPECIFICACIONES[mod_d] and os.path.exists(ESPECIFICACIONES[mod_d]["Manual"]):
                    with open(ESPECIFICACIONES[mod_d]["Manual"], "rb") as f: st.download_button(label=f"📕 Descargar Manual de {mod_d} (PDF)", data=f, file_name=ESPECIFICACIONES[mod_d]["Manual"].split('/')[-1], mime="application/pdf")
        with tab3:
            st.markdown(f"### 🔍 Bitácora Permanente del Equipo: {tag_sel}")
            with st.form(key=f"form_obs_{tag_sel}"):
                nueva_obs = st.text_area("Escribe una nueva observación:", height=100)
                if st.form_submit_button("➕ Dejar constancia en la bitácora", use_container_width=True):
                    if nueva_obs: agregar_observacion(tag_sel, st.session_state.usuario_actual, nueva_obs); st.success("✅ Observación registrada."); st.rerun()
            st.markdown("---"); df_obs = obtener_observaciones(tag_sel)
            if not df_obs.empty:
                for _, row in df_obs.iterrows():
                    col_obs, col_del = st.columns([11, 1])
                    with col_obs: st.markdown(f"<div style='background-color: #2b303b; padding: 15px; border-radius: 8px; margin-bottom: 10px; border-left: 4px solid #FF6600;'><small style='color: #aeb9cc;'><b>👤 Técnico: {row['usuario']}</b> &nbsp;|&nbsp; 📅 Fecha: {row['fecha']}</small><br><span style='color: white; font-size: 1.05em;'>{row['texto']}</span></div>", unsafe_allow_html=True)
                    with col_del:
                        if st.button("🗑️", key=f"del_obs_{row['id']}"): eliminar_observacion(row['id']); st.rerun()
        with tab4:
            st.markdown(f"### 👤 Información de Contactos y Seguridad del Área: {tag_sel}")
            with st.expander("✏️ Editar o Agregar Contacto / Dato de Seguridad"):
                with st.form(key=f"form_area_{tag_sel}"):
                    c_a1, c_a2 = st.columns(2); opc_area = ["Dueño de Área (Turno 1-3)", "Dueño de Área (Turno 2-4)", "PEA", "Frecuencia Radial", "Supervisor a cargo", "Jefe de Turno", "Otro cargo..."]; clave_sel_area = c_a1.selectbox("¿Qué dato vas a ingresar?", opc_area); clave_final_area = c_a1.text_input("Escribe el nombre del cargo:") if clave_sel_area == "Otro cargo..." else clave_sel_area; valor_final_area = c_a2.text_input("Ingresa la información:")
                    if st.form_submit_button("💾 Guardar Información", use_container_width=True):
                        if clave_final_area and valor_final_area: guardar_dato_equipo(tag_sel, clave_final_area.strip(), valor_final_area.strip()); st.success("✅ Dato actualizado!"); st.rerun()
            datos_equipo = obtener_datos_equipo(tag_sel); cols_area = st.columns(2)
            for i, (k, v) in enumerate(datos_equipo.items()):
                with cols_area[i % 2]: st.markdown(f"<div style='background-color: #2b303b; padding: 15px; border-radius: 8px; margin-bottom: 15px; border-left: 4px solid #FF6600;'><span style='color: #aeb9cc; font-size: 0.85em; text-transform: uppercase; font-weight: bold;'>{k}</span><br><span style='color: white; font-size: 1.1em;'>{v}</span></div>", unsafe_allow_html=True)
        st.markdown("<br><hr>", unsafe_allow_html=True); st.markdown("### 📋 Trazabilidad Histórica de Intervenciones"); df_hist = obtener_todo_el_historial(tag_sel)
        if not df_hist.empty: st.dataframe(df_hist, use_container_width=True)