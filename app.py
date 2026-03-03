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
import datetime
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
# 0.2 ESTILOS PREMIUM (DISEÑO NATIVO Y BOTÓN ELEGANTE)
# =============================================================================
st.set_page_config(page_title="Atlas Spence | Gestión de Reportes", layout="wide", page_icon="⚙️", initial_sidebar_state="expanded")

def aplicar_estilos_premium():
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;600;800&display=swap');
        :root { --ac-blue: #007CA6; --ac-dark: #005675; --bhp-orange: #FF6600; }
        
        html, body, p, h1, h2, h3, h4, h5, h6, span, div { font-family: 'Montserrat', sans-serif; }
        
        #MainMenu {visibility: hidden;} 
        footer {visibility: hidden;} 
        
        /* ========================================================= */
        /* 🔥 BOTÓN DE MENÚ LATERAL: ELEGANTE Y DE ALTO CONTRASTE    */
        /* ========================================================= */
        [data-testid="collapsedControl"] {
            display: flex !important;
            visibility: visible !important;
            opacity: 1 !important;
            z-index: 999999 !important;
            background-color: #00BFFF !important; /* Celeste vibrante de alto contraste */
            border-radius: 8px !important;
            box-shadow: 0 4px 15px rgba(0, 191, 255, 0.4) !important;
            margin-top: 15px !important;
            margin-left: 15px !important;
            transition: all 0.3s ease !important;
        }
        [data-testid="collapsedControl"]:hover {
            background-color: var(--ac-blue) !important; /* Azul Atlas Copco al pasar el mouse */
            box-shadow: 0 6px 20px rgba(0, 124, 166, 0.6) !important;
            transform: scale(1.05) !important;
        }
        [data-testid="collapsedControl"] svg {
            fill: white !important;
            stroke: white !important;
        }
        /* ========================================================= */
        
        div.stButton > button:first-child {
            background: linear-gradient(135deg, var(--ac-blue) 0%, var(--ac-dark) 100%);
            color: white; border-radius: 8px; border: none; font-weight: 600; padding: 0.6rem 1.2rem;
            transition: all 0.3s ease; box-shadow: 0 4px 15px rgba(0, 124, 166, 0.4);
        }
        div.stButton > button:first-child:hover { transform: translateY(-2px); box-shadow: 0 8px 20px rgba(0, 124, 166, 0.6); }
        
        [data-testid="stVerticalBlockBorderWrapper"] {
            background: linear-gradient(145deg, #1a212b, #151a22) !important;
            border-radius: 12px !important; border: 1px solid #2b3543 !important;
            transition: transform 0.3s ease, box-shadow 0.3s ease, border-color 0.3s ease !important;
        }
        [data-testid="stVerticalBlockBorderWrapper"]:hover {
            transform: translateY(-6px) !important;
            box-shadow: 0 10px 25px rgba(0, 124, 166, 0.25) !important;
            border-color: var(--ac-blue) !important;
        }
        
        div[class^="st-key-btn_"] button {
            background: transparent !important; border: 1px solid rgba(255,255,255,0.05) !important;
            color: white !important; font-size: 1.6rem !important; font-weight: 800 !important;
            padding: 1.2rem !important; border-radius: 8px !important; box-shadow: none !important;
        }
        div[class^="st-key-btn_"] button:hover {
            background: rgba(0, 124, 166, 0.2) !important; border-color: var(--ac-blue) !important;
            color: #fff !important; box-shadow: inset 0 0 15px rgba(0,124,166,0.3) !important;
        }
        
        .stTextInput>div>div>input, .stNumberInput>div>div>input, .stSelectbox>div>div>select { 
            border-radius: 6px !important; border: 1px solid #2b3543 !important; 
            background-color: #1e2530 !important; color: white !important;
        }
        .stTextInput>div>div>input:focus, .stNumberInput>div>div>input:focus, .stSelectbox>div>div>select:focus { 
            border-color: var(--bhp-orange) !important; box-shadow: 0 0 8px rgba(255, 102, 0, 0.3) !important; 
        }
        .stTabs [data-baseweb="tab-list"] { border-bottom: 2px solid #2b3543; }
        .stTabs [aria-selected="true"] { color: var(--bhp-orange) !important; border-bottom: 3px solid var(--bhp-orange) !important; }
        </style>
    """, unsafe_allow_html=True)
aplicar_estilos_premium()

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
    "Taller": ["GA 18", "API335343", "Taller", "Taller"]
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
import datetime
from google.oauth2.service_account import Credentials
from streamlit_pdf_viewer import pdf_viewer

# =============================================================================
# 0. CONFIGURACIÓN INICIAL Y ESTILOS
# =============================================================================
RUTA_ONEDRIVE = "Reportes_Temporales" 
MI_CORREO_CORPORATIVO = "ignacio.a.morales@atlascopco.com"
CORREO_REMITENTE = "informeatlas.spence@gmail.com"
PASSWORD_APLICACION = "jbumdljbdpyomnna"

st.set_page_config(page_title="Atlas Spence | Planificación", layout="wide", page_icon="⚙️", initial_sidebar_state="expanded")

def aplicar_estilos_premium():
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;600;800&display=swap');
        :root { --ac-blue: #007CA6; --ac-dark: #005675; --bhp-orange: #FF6600; }
        html, body, p, h1, h2, h3, h4, h5, h6, span, div { font-family: 'Montserrat', sans-serif; }
        #MainMenu {visibility: hidden;} footer {visibility: hidden;} 
        
        [data-testid="collapsedControl"] {
            display: flex !important; visibility: visible !important; opacity: 1 !important;
            z-index: 999999 !important; background-color: #00BFFF !important; 
            border-radius: 8px !important; box-shadow: 0 4px 15px rgba(0, 191, 255, 0.4) !important;
            margin-top: 15px !important; margin-left: 15px !important; transition: all 0.3s ease !important;
        }
        [data-testid="collapsedControl"]:hover { background-color: var(--ac-blue) !important; transform: scale(1.05) !important; }
        [data-testid="collapsedControl"] svg { fill: white !important; stroke: white !important; }
        
        div.stButton > button:first-child { background: linear-gradient(135deg, var(--ac-blue) 0%, var(--ac-dark) 100%); color: white; border-radius: 8px; border: none; font-weight: 600; }
        [data-testid="stVerticalBlockBorderWrapper"] { background: #1a212b !important; border-radius: 12px !important; border: 1px solid #2b3543 !important; }
        .stTextInput>div>div>input, .stNumberInput>div>div>input, .stSelectbox>div>div>select { border-radius: 6px !important; background-color: #1e2530 !important; color: white !important; }
        </style>
    """, unsafe_allow_html=True)
aplicar_estilos_premium()

# =============================================================================
# 1. DATOS MAESTROS Y CONEXIÓN A GOOGLE SHEETS
# =============================================================================
USUARIOS = {"ignacio morales": "spence2026", "emian": "spence2026", "ignacio veas": "spence2026", "admin": "admin123"}
DEFAULT_SPECS = {
    "GA 18": {"Litros de Aceite": "14.1 L", "Manual": "manuales/manual_ga18.pdf"},
    "GA 30": {"Litros de Aceite": "14.6 L", "Manual": "manuales/manual_ga30.pdf"},
    "GA 37": {"Litros de Aceite": "14.6 L", "Manual": "manuales/manual_ga37.pdf"},
    "GA 45": {"Litros de Aceite": "17.9 L", "Manual": "manuales/manual_ga45.pdf"},
    "GA 75": {"Litros de Aceite": "35.2 L", "Manual": "manuales/manual_ga75.pdf"},
    "GA 90": {"Litros de Aceite": "69 L", "Manual": "manuales/manual_ga90.pdf"},
    "GA 132": {"Litros de Aceite": "93 L", "Manual": "manuales/manual_ga132.pdf"},
    "GA 250": {"Litros de Aceite": "130 L", "Manual": "manuales/manual_ga250.pdf"},
    "ZT 37": {"Litros de Aceite": "23 L", "Manual": "manuales/manual_zt37.pdf"},
    "CD 80+": {"Filtro de Gases": "DD/PD 80", "Manual": "manuales/manual_cd80.pdf"},
    "CD 630": {"Filtro de Gases": "DD/PD 630", "Manual": "manuales/manual_cd630.pdf"}
}

inventario_equipos = {
    "20-GC-001": ["GA 75", "AII482673", "truck shop", "mina"], "20-GC-002": ["GA 75", "AII482674", "truck shop", "mina"], "20-GC-003": ["GA 90", "AIF095178", "truck shop", "mina"], "20-GC-004": ["GA 37", "AII390776", "truck shop", "mina"],
    "35-GC-006": ["GA 250", "AIF095420", "chancado secundario", "área seca"], "35-GC-007": ["GA 250", "AIF095421", "chancado secundario", "área seca"], "35-GC-008": ["GA 250", "AIF095302", "chancado secundario", "área seca"],
    "50-GC-001": ["GA 45", "API542705", "planta SX", "área húmeda"], "50-GC-002": ["GA 45", "API542706", "planta SX", "área húmeda"], "50-GC-003": ["ZT 37", "API791692", "planta SX", "área húmeda"], "50-GC-004": ["ZT 37", "API791693", "planta SX", "área húmeda"], "50-CD-001": ["CD 80+", "API095825", "planta SX", "área húmeda"], "50-CD-002": ["CD 80+", "API095826", "planta SX", "área húmeda"],
    "55-GC-015": ["GA 30", "API501440", "planta borra", "área húmeda"],
    "65-GC-009": ["GA 250", "APF253608", "patio de estanques", "área húmeda"], "65-GC-011": ["GA 250", "APF253581", "patio de estanques", "área húmeda"], "65-CD-011": ["CD 630", "WXF300015", "patio de estanques", "área húmeda"], "65-CD-012": ["CD 630", "WXF300016", "patio de estanques", "área húmeda"],
    "70-GC-013": ["GA 132", "AIF095296", "descarga de acido", "área húmeda"], "70-GC-014": ["GA 132", "AIF095297", "descarga de acido", "área húmeda"],
    "Taller": ["GA 18", "API335343", "Taller", "Taller"]
}

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
        except: return doc.add_worksheet(title=sheet_name, rows="1000", cols="50")
    except: return None

# =============================================================================
# 2. CONSTRUCCIÓN DE LA MATRIZ BASE (ESTRICTA: Pauta | WK Plan | Estado | WK Real)
# =============================================================================
def generar_planificacion_base():
    tags = [
        ("70-GC-013", "GA 132", "Descarga Acido"), ("70-GC-014", "GA 132", "Descarga Acido"),
        ("50-GC-001", "GA 45", "Planta SX"), ("50-GC-002", "GA 45", "Planta SX"),
        ("50-GC-003", "ZT 37", "Planta SX"), ("50-GC-004", "ZT 37", "Planta SX"),
        ("50-CD-001", "CD 80+", "Planta SX"), ("50-CD-002", "CD 80+", "Planta SX"),
        ("55-GC-015", "GA 30", "Planta Borra"),
        ("65-GC-011", "GA 250", "Patio Estanques"), ("65-GC-009", "GA 250", "Patio Estanques"),
        ("65-CD-011", "CD 630", "Patio Estanques"), ("65-CD-012", "CD 630", "Patio Estanques"),
        ("35-GC-006", "GA 250", "Chancado Sec."), ("35-GC-007", "GA 250", "Chancado Sec."), ("35-GC-008", "GA 250", "Chancado Sec."),
        ("20-GC-004", "GA 37", "Truck Shop"), ("20-GC-001", "GA 75", "Truck Shop"),
        ("20-GC-002", "GA 75", "Truck Shop"), ("20-GC-003", "GA 90", "Truck Shop"),
        ("Taller", "GA 18", "Taller")
    ]
    meses = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
    
    # Excepciones manuales extraídas del Excel
    excepciones = {
        "70-GC-013": {"Feb_Pauta": "P1", "Feb_Est": "Hecho", "Feb_WK_Real": "WK7", "Mar_Est": "Hecho", "Mar_WK_Real": "W10", "Abr_Pauta": "P4", "Jun_Pauta": "P1", "Ago_Pauta": "P2", "Oct_Pauta": "P1", "Dic_Pauta": "P3"},
        "70-GC-014": {"Ene_Pauta": "P2", "Ene_Est": "Hecho", "Ene_WK_Real": "WK3", "Mar_Pauta": "P1", "Mar_Est": "Hecho", "Mar_WK_Real": "W10", "May_Pauta": "P3", "Jul_Pauta": "P1", "Sep_Pauta": "P2", "Nov_Pauta": "P1"},
        "50-GC-001": {"Feb_Pauta": "P1", "Feb_Est": "Hecho", "Feb_WK_Real": "WK4", "Mar_Est": "Hecho", "Mar_WK_Real": "W10", "Abr_Pauta": "P3", "Jun_Pauta": "P1", "Ago_Pauta": "P2", "Oct_Pauta": "P1", "Dic_Pauta": "P3"},
        "50-GC-002": {"Ene_Pauta": "P2", "Feb_Est": "Hecho", "Feb_WK_Real": "WK4", "Mar_Pauta": "P1", "Mar_Est": "Hecho", "Mar_WK_Real": "W10", "May_Pauta": "P3", "Jul_Pauta": "P1", "Sep_Pauta": "P2", "Nov_Pauta": "P1"},
        "50-GC-003": {"Feb_Pauta": "P1", "Feb_WK_Real": "WK7", "Mar_WK_Real": "WK9", "Abr_Pauta": "P4", "Jun_Pauta": "P1", "Ago_Pauta": "P2", "Oct_Pauta": "P1", "Dic_Pauta": "P3"},
        "50-GC-004": {"Ene_Pauta": "P2", "Mar_WK_Real": "WK8", "May_Pauta": "P4", "Jul_Pauta": "P1", "Sep_Pauta": "P2", "Nov_Pauta": "P1"},
        "50-CD-001": {"Ene_Pauta": "P4", "Mar_Est": "Hecho", "Mar_WK_Real": "WK8", "Jul_Pauta": "P2"},
        "50-CD-002": {"Ene_Pauta": "P4", "Mar_Est": "Hecho", "Mar_WK_Real": "WK8", "Jul_Pauta": "P2"},
        "55-GC-015": {"Feb_Pauta": "P1", "Feb_Est": "Hecho", "Feb_WK_Real": "WK6", "Mar_Est": "Hecho", "Mar_WK_Real": "WK11", "Abr_Pauta": "P4", "Jun_Pauta": "P1", "Ago_Pauta": "P2", "Oct_Pauta": "P1", "Dic_Pauta": "P3"},
        "65-GC-011": {"Feb_Pauta": "P1", "Feb_Est": "Hecho", "Feb_WK_Real": "WK5", "Mar_Est": "Hecho", "Mar_WK_Real": "WK11", "Abr_Pauta": "P1", "Jun_Pauta": "P2", "Ago_Pauta": "P1", "Oct_Pauta": "P1", "Dic_Pauta": "P4"},
        "65-GC-009": {"Ene_Pauta": "P1", "Mar_Pauta": "P4", "Mar_Est": "Hecho", "Mar_WK_Real": "WK8", "May_Pauta": "P1", "Jul_Pauta": "P1", "Sep_Pauta": "P2", "Nov_Pauta": "P1"},
        "65-CD-011": {"Feb_Pauta": "P2", "Mar_Est": "Hecho", "Mar_WK_Real": "WK8", "May_Pauta": "P2", "Ago_Pauta": "P2", "Nov_Pauta": "P2"},
        "65-CD-012": {"Feb_Pauta": "P2", "Mar_Est": "Hecho", "Mar_WK_Real": "WK8", "May_Pauta": "P2", "Ago_Pauta": "P2", "Nov_Pauta": "P2"},
        "35-GC-006": {"Ene_Pauta": "P1", "Feb_Pauta": "P1", "Mar_Pauta": "P2", "Mar_WK_Real": "WK11", "Abr_Pauta": "P1", "May_Pauta": "P1", "Jun_Pauta": "P2", "Jul_Pauta": "P1", "Ago_Pauta": "P1", "Sep_Pauta": "P4", "Oct_Pauta": "P1", "Nov_Pauta": "P1", "Dic_Pauta": "P2"},
        "35-GC-007": {"Ene_Pauta": "P3", "Ene_Est": "Hecho", "Feb_Pauta": "P1", "Feb_Est": "Hecho", "Feb_WK_Real": "W6", "Mar_Pauta": "P1", "Mar_Est": "Hecho", "Mar_WK_Real": "WK11", "Abr_Pauta": "P2", "May_Pauta": "P1", "Jun_Pauta": "P1", "Jul_Pauta": "P2", "Ago_Pauta": "P1", "Sep_Pauta": "P1", "Oct_Pauta": "P4", "Nov_Pauta": "P1", "Dic_Pauta": "P1"},
        "35-GC-008": {"Ene_Pauta": "P1", "Feb_Pauta": "P2", "Feb_Est": "Hecho", "Feb_WK_Real": "W6", "Mar_Pauta": "P1", "Mar_Est": "Hecho", "Mar_WK_Real": "WK11", "Abr_Pauta": "P1", "May_Pauta": "P2", "Jun_Pauta": "P1", "Jul_Pauta": "P1", "Ago_Pauta": "P4", "Sep_Pauta": "P1", "Oct_Pauta": "P1", "Nov_Pauta": "P2", "Dic_Pauta": "P1"},
        "20-GC-004": {"Feb_Pauta": "P1", "Feb_WK_Real": "WK5", "Mar_Pauta": "P1", "Mar_Est": "Hecho", "Mar_WK_Real": "WK10", "May_Pauta": "P4", "Ago_Pauta": "P1", "Nov_Pauta": "P2"},
        "20-GC-001": {"Feb_Pauta": "P1", "Feb_Est": "Hecho", "Feb_WK_Real": "WK4", "Mar_Est": "Hecho", "Mar_WK_Real": "WK10", "Abr_Pauta": "P4", "Jun_Pauta": "P1", "Ago_Pauta": "P2", "Oct_Pauta": "P1", "Dic_Pauta": "P3"},
        "20-GC-002": {"Feb_Pauta": "P1", "Feb_WK_Real": "WK4", "Mar_Est": "Hecho", "Mar_WK_Real": "WK10", "Abr_Pauta": "P4", "Jun_Pauta": "P1", "Ago_Pauta": "P2", "Oct_Pauta": "P1", "Dic_Pauta": "P3"},
        "20-GC-003": {"Feb_Pauta": "P1", "Feb_Est": "Hecho", "Feb_WK_Real": "W7", "Mar_Est": "Hecho", "Mar_WK_Real": "WK10", "Abr_Pauta": "P4", "Jun_Pauta": "P1", "Ago_Pauta": "P2", "Oct_Pauta": "P1", "Dic_Pauta": "P3"},
        "Taller": {"Feb_Pauta": "P2", "Feb_Est": "Hecho", "Feb_WK_Real": "WK5"}
    }

    datos = []
    for t, eq, ar in tags:
        row = {"TAG": t, "Equipo": eq, "Área": ar}
        for m in meses:
            # Por defecto
            row[f"{m} Pauta"] = "INSP"
            row[f"{m} WK Plan"] = ""
            row[f"{m} Estado"] = "Pendiente"
            row[f"{m} WK Real"] = ""
            
            # Aplicar excepciones si existen
            if t in excepciones:
                row[f"{m} Pauta"] = excepciones[t].get(f"{m}_Pauta", "INSP")
                row[f"{m} WK Plan"] = excepciones[t].get(f"{m}_W_Plan", "")
                row[f"{m} Estado"] = excepciones[t].get(f"{m}_Est", "Pendiente")
                row[f"{m} WK Real"] = excepciones[t].get(f"{m}_WK_Real", "")
                
        datos.append(row)
    return pd.DataFrame(datos)

@st.cache_data(ttl=60)
def cargar_planificacion():
    try:
        sheet = get_sheet("planificacion")
        if sheet:
            data = sheet.get_all_values()
            if len(data) > 1:
                df = pd.DataFrame(data[1:], columns=data[0])
                if "Ene WK Plan" in df.columns: return df
    except Exception as e: pass
    return generar_planificacion_base()

def guardar_planificacion(df):
    try:
        sheet = get_sheet("planificacion")
        if sheet:
            sheet.clear() 
            datos_a_guardar = [df.columns.values.tolist()] + df.values.tolist()
            sheet.append_rows(datos_a_guardar)
            st.cache_data.clear() 
    except Exception as e:
        st.error(f"Error al conectar con la Nube: {e}")

def aplicar_estilos_matriz(df):
    """Aplica colores a la fila entera evaluando la columna Estado y Pauta."""
    def style_row(row):
        styles = [''] * len(row)
        for i, col in enumerate(row.index):
            val = str(row[col]).upper()
            
            if col.endswith('Pauta'):
                if 'P1' in val: styles[i] = 'background-color: #0c2d48; color: #66c2ff; font-weight: bold; text-align: center;'
                elif 'P2' in val: styles[i] = 'background-color: #4a2c00; color: #ffb04c; font-weight: bold; text-align: center;'
                elif 'P3' in val: styles[i] = 'background-color: #301047; color: #d78aff; font-weight: bold; text-align: center;'
                elif 'P4' in val: styles[i] = 'background-color: #471015; color: #ff8a93; font-weight: bold; text-align: center;'
                else: styles[i] = 'color: #8c9eb5; text-align: center;'
                
            elif col.endswith('Estado'):
                if val == 'HECHO': styles[i] = 'background-color: #063f22; color: #6ee7b7; font-weight: bold; text-align: center;'
                elif val == 'PENDIENTE': styles[i] = 'background-color: #423205; color: #fde047; font-weight: bold; text-align: center;'
                else: styles[i] = 'color: white; text-align: center;'
                    
            elif col.endswith('WK Plan') or col.endswith('WK Real'):
                styles[i] = 'color: #aeb9cc; text-align: center;'
                
        return styles
    return df.style.apply(style_row, axis=1)

# =============================================================================
# 3. INTERFAZ PRINCIPAL
# =============================================================================
ESPECIFICACIONES = obtener_especificaciones(DEFAULT_SPECS)
default_states = {'logged_in': False, 'usuario_actual': "", 'equipo_seleccionado': None, 'vista_actual': "catalogo"}
for key, value in default_states.items():
    if key not in st.session_state: st.session_state[key] = value
if 'informes_pendientes' not in st.session_state: st.session_state.informes_pendientes = []

if not st.session_state.logged_in:
    st.markdown("<br><br><br>", unsafe_allow_html=True); _, col_centro, _ = st.columns([1, 1.5, 1])
    with col_centro:
        with st.container(border=True):
            st.markdown("<h1 style='text-align: center; border-bottom:none;'>⚙️ <span style='color:#007CA6;'>Atlas</span> <span style='color:#FF6600;'>Spence</span></h1>", unsafe_allow_html=True)
            with st.form("form_login"):
                u_in = st.text_input("Usuario").lower(); p_in = st.text_input("Contraseña", type="password")
                if st.form_submit_button("Acceder", type="primary", use_container_width=True):
                    if u_in in USUARIOS and USUARIOS[u_in] == p_in: 
                        st.session_state.update({'logged_in': True, 'usuario_actual': u_in})
                        st.rerun()
else:
    with st.sidebar:
        st.markdown("<h2 style='text-align: center; margin-top: -20px;'><span style='color:#007CA6;'>Atlas</span> <span style='color:#FF6600;'>Spence</span></h2>", unsafe_allow_html=True)
        if st.button("🏭 Catálogo de Activos", use_container_width=True): st.session_state.vista_actual = "catalogo"; st.rerun()
        if st.button("📅 Planificación Anual", use_container_width=True, type="primary"): st.session_state.vista_actual = "planificacion"; st.rerun()
        if st.button("🚪 Cerrar Sesión", use_container_width=True): st.session_state.logged_in = False; st.rerun()

    # --- VISTA PLANIFICACIÓN MATRIZ ---
    if st.session_state.vista_actual == "planificacion":
        df_plan = cargar_planificacion()
        
        st.markdown(f"""
            <div style="margin-top: 1rem; margin-bottom: 1rem; background: linear-gradient(90deg, rgba(0,124,166,0.1) 0%, rgba(0,124,166,0.2) 50%, rgba(0,124,166,0.1) 100%); padding: 20px; border-radius: 15px; border-left: 5px solid var(--ac-blue);">
                <h2 style="color: white; margin: 0;">📅 Matriz de Planificación Anual</h2>
                <p style="color: #8c9eb5; margin: 0; font-weight: 600;">Estructura Oficial de Confiabilidad: Pauta | WK Plan | Estado | WK Real</p>
            </div>
        """, unsafe_allow_html=True)
        
        col_fil1, col_fil2, col_fil3 = st.columns([1, 1, 1.5])
        with col_fil1:
            areas_disp = ["Todas"] + sorted(list(df_plan["Área"].unique()))
            filtro_area = st.selectbox("🏢 Filtrar por Área:", areas_disp)
        with col_fil2:
            modo_edicion_matriz = st.toggle("✏️ Habilitar Edición de Matriz")
        with col_fil3:
            st.markdown("<div style='margin-top:30px;'></div>", unsafe_allow_html=True)
            if modo_edicion_matriz: st.info("Haz doble clic para abrir el menú desplegable en Pauta y Estado.")
            
        df_mostrar = df_plan.copy()
        if filtro_area != "Todas": df_mostrar = df_mostrar[df_mostrar["Área"] == filtro_area]
        
        meses_list = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
        
        # Configurar Columnas Desplegables
        config_cols = {
            "TAG": st.column_config.TextColumn(disabled=True, width="small"),
            "Equipo": st.column_config.TextColumn(disabled=True, width="small"),
            "Área": st.column_config.TextColumn(disabled=True, width="small")
        }
        for m in meses_list:
            config_cols[f"{m} Pauta"] = st.column_config.SelectboxColumn("Pauta", options=["INSP", "P1", "P2", "P3", "P4"], width="small")
            config_cols[f"{m} WK Plan"] = st.column_config.TextColumn("W_Plan", width="small")
            config_cols[f"{m} Estado"] = st.column_config.SelectboxColumn("Est", options=["Pendiente", "Hecho"], width="small")
            config_cols[f"{m} WK Real"] = st.column_config.TextColumn("W_Real", width="small")

        if modo_edicion_matriz:
            df_editado = st.data_editor(df_mostrar, use_container_width=True, hide_index=True, height=750, column_config=config_cols)
            if st.button("💾 Guardar Matriz en Nube", type="primary", use_container_width=True):
                df_final_guardar = df_plan.copy()
                df_editado_str = df_editado.astype(str)
                df_final_guardar.update(df_editado_str)
                guardar_planificacion(df_final_guardar)
                st.success("✅ ¡Base de Datos actualizada con éxito!")
                st.rerun()
        else:
            df_estilizado_view = aplicar_estilos_matriz(df_mostrar)
            st.dataframe(df_estilizado_view, use_container_width=True, hide_index=True, height=750, column_config=config_cols)

    # --- 6.2 VISTA CATÁLOGO (SIEMPRE DISPONIBLE) ---
    elif st.session_state.vista_actual == "catalogo":
        st.markdown("<h1 style='text-align: center; color: #007CA6;'>Atlas Copco <span style='color: #FF6600;'>Spence</span></h1>", unsafe_allow_html=True)
        st.markdown("<hr>", unsafe_allow_html=True)
        columnas = st.columns(4); contador = 0
        for tag, (modelo, serie, area, ubicacion) in inventario_equipos.items():
            with columnas[contador % 4]:
                with st.container(border=True):
                    st.button(f"{tag}", key=f"btn_{tag}", use_container_width=True)
                    st.markdown(f"<p style='color: #8c9eb5; margin-top: 5px; font-size: 0.85rem; text-align: center;'>{modelo} • {area.title()}</p>", unsafe_allow_html=True)
            contador += 1