import streamlit as st

# 🔥 CONFIGURACIÓN DE PÁGINA (Debe ser la línea 2)
st.set_page_config(page_title="Atlas Spence | Gestión de Reportes", layout="wide", page_icon="⚙️", initial_sidebar_state="expanded")

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
import calendar
import re
from google.oauth2.service_account import Credentials
from streamlit_pdf_viewer import pdf_viewer

# =============================================================================
# 0.1 CONFIGURACIÓN DE NUBE Y CORREO
# =============================================================================
RUTA_ONEDRIVE = "Reportes_Temporales" 
RUTA_APROBADOS = "Reportes_Aprobados"
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
# 0.2 ESTILOS PREMIUM (BORDES CIRCULARES)
# =============================================================================
def aplicar_estilos_premium():
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;600;800&display=swap');
        :root { --ac-blue: #007CA6; --ac-dark: #005675; --bhp-orange: #FF6600; }
        html, body, p, h1, h2, h3, h4, h5, h6, span, div { font-family: 'Montserrat', sans-serif; }
        
        @keyframes cinematicFadeIn { 0% { opacity: 0; transform: translateY(15px); } 100% { opacity: 1; transform: translateY(0); } }
        .main .block-container { animation: cinematicFadeIn 0.5s cubic-bezier(0.22, 1, 0.36, 1) forwards; }
        
        [data-testid="stStatusWidget"] { visibility: hidden !important; display: none !important; }
        [data-testid="stToolbar"] { visibility: hidden !important; display: none !important; } 
        a[href*="github.com"] { display: none !important; visibility: hidden !important; }
        [data-testid="viewerBadge"], div[class^="viewerBadge_container"], footer { display: none !important; }
        
        div.stButton > button:first-child { background: linear-gradient(135deg, var(--ac-blue) 0%, var(--ac-dark) 100%); color: white; border-radius: 20px; border: none; font-weight: 600; padding: 0.6rem 1.2rem; transition: all 0.3s ease; box-shadow: 0 4px 15px rgba(0, 124, 166, 0.4); }
        div.stButton > button:first-child:hover { transform: translateY(-2px); box-shadow: 0 8px 20px rgba(0, 124, 166, 0.6); }
        
        /* 🔥 BORDES REDONDEADOS / CIRCULARES PARA PANELES Y FORMULARIOS 🔥 */
        [data-testid="stVerticalBlockBorderWrapper"] { background: linear-gradient(145deg, #1a212b, #151a22) !important; border-radius: 30px !important; border: 1px solid #2b3543 !important; transition: transform 0.3s ease, box-shadow 0.3s ease, border-color 0.3s ease !important; }
        [data-testid="stVerticalBlockBorderWrapper"]:hover { transform: translateY(-6px) !important; box-shadow: 0 10px 25px rgba(0, 124, 166, 0.25) !important; border-color: var(--ac-blue) !important; }
        [data-testid="stForm"] { border-radius: 25px !important; border: 1px solid #2b3543 !important; }
        
        .stTextInput>div>div>input, .stNumberInput>div>div>input, .stSelectbox>div>div>select, .stDateInput>div>div>input { border-radius: 12px !important; border: 1px solid #2b3543 !important; background-color: #1e2530 !important; color: white !important; }
        .stTextInput>div>div>input:focus, .stNumberInput>div>div>input:focus, .stSelectbox>div>div>select:focus, .stDateInput>div>div>input:focus { border-color: var(--bhp-orange) !important; box-shadow: 0 0 8px rgba(255, 102, 0, 0.3) !important; }
        .stTabs [data-baseweb="tab-list"] { border-bottom: 2px solid #2b3543; }
        .stTabs [aria-selected="true"] { color: var(--bhp-orange) !important; border-bottom: 3px solid var(--bhp-orange) !important; }
        
        /* 🔥 OCULTAR AVISO DE "Press Enter to apply" 🔥 */
        div[data-testid="InputInstructions"], 
        div[data-testid="stNumberInput"] small { 
            display: none !important; 
            visibility: hidden !important; 
        }

        /* 🚀 BOTÓN FÍSICO FLOTANTE PARA ABRIR LA BARRA LATERAL (SI LA CIERRAN) 🚀 */
        [data-testid="collapsedControl"] {
            background: linear-gradient(135deg, var(--bhp-orange) 0%, #cc5200 100%) !important;
            border-radius: 50% !important;
            box-shadow: 0 4px 15px rgba(255, 102, 0, 0.5) !important;
            top: 25px !important;
            left: 25px !important;
            width: 45px !important;
            height: 45px !important;
            z-index: 999999 !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            transition: transform 0.3s ease !important;
            opacity: 1 !important;
        }
        [data-testid="collapsedControl"]:hover {
            transform: scale(1.1) !important;
            background: linear-gradient(135deg, #ff7a22 0%, var(--bhp-orange) 100%) !important;
        }
        [data-testid="collapsedControl"] svg {
            fill: white !important;
            color: white !important;
            width: 25px !important;
            height: 25px !important;
        }
        
        /* 🚀 BOTÓN FÍSICO PARA CERRAR DENTRO DE LA BARRA LATERAL 🚀 */
        [data-testid="stSidebar"] button[kind="header"] {
            background-color: rgba(255, 255, 255, 0.1) !important;
            border-radius: 50% !important;
            width: 35px !important;
            height: 35px !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            transition: background-color 0.3s !important;
        }
        [data-testid="stSidebar"] button[kind="header"]:hover {
            background-color: var(--bhp-orange) !important;
        }
        [data-testid="stSidebar"] button[kind="header"] svg {
            fill: white !important;
        }
        </style>
    """, unsafe_allow_html=True)
aplicar_estilos_premium()

# =============================================================================
# 1. DATOS MAESTROS Y ROLES (RBAC)
# =============================================================================
USUARIOS = {"ignacio morales": "spence2026", "emian": "spence2026", "ignacio veas": "spence2026", "yerko villarroel": "spence2026", "admin": "admin123"}
ADMIN_USERS = ["ignacio morales", "admin"]

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
    "20-GC-001": ["GA 75", "AII482673", "truck shop", "Mina"], "20-GC-002": ["GA 75", "AII482674", "truck shop", "Mina"], "20-GC-003": ["GA 90", "AIF095178", "truck shop", "Mina"], "20-GC-004": ["GA 37", "AII390776", "truck shop", "Mina"],
    "35-GC-006": ["GA 250", "AIF095420", "chancado secundario", "Área Seca"], "35-GC-007": ["GA 250", "AIF095421", "chancado secundario", "Área Seca"], "35-GC-008": ["GA 250", "AIF095302", "chancado secundario", "Área Seca"],
    "50-GC-001": ["GA 45", "API542705", "planta SX", "Área Húmeda"], "50-GC-002": ["GA 45", "API542706", "planta SX", "Área Húmeda"], "50-GC-003": ["ZT 37", "API791692", "planta SX", "Área Húmeda"], "50-GC-004": ["ZT 37", "API791693", "planta SX", "Área Húmeda"], "50-CD-001": ["CD 80+", "API095825", "planta SX", "Área Húmeda"], "50-CD-002": ["CD 80+", "API095826", "planta SX", "Área Húmeda"],
    "55-GC-015": ["GA 30", "API501440", "planta borra", "Área Húmeda"],
    "65-GC-009": ["GA 250", "APF253608", "patio de estanques", "Área Húmeda"], "65-GC-011": ["GA 250", "APF253581", "patio de estanques", "Área Húmeda"], "65-CD-011": ["CD 630", "WXF300015", "patio de estanques", "Área Húmeda"], "65-CD-012": ["CD 630", "WXF300016", "patio de estanques", "Área Húmeda"],
    "70-GC-013": ["GA 132", "AIF095296", "descarga de acido", "Área Húmeda"], "70-GC-014": ["GA 132", "AIF095297", "descarga de acido", "Área Húmeda"],
    "Taller": ["GA 18", "API335343", "Taller", "Laboratorio"] 
}

# =============================================================================
# 2. CONEXIÓN A GOOGLE SHEETS
# =============================================================================
@st.cache_resource(show_spinner=False)
def get_gspread_client():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    try: creds_dict = json.loads(os.environ["gcp_json"])
    except: creds_dict = json.loads(st.secrets["gcp_json"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return gspread.authorize(creds)

def get_sheet(sheet_name):
    try:
        client = get_gspread_client()
        doc = client.open("BaseDatos")
        try: return doc.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound: return doc.add_worksheet(title=sheet_name, rows="1000", cols="20")
    except Exception as e: return None

# =============================================================================
# 3. FUNCIONES DE BASE DE DATOS
# =============================================================================
@st.cache_data(ttl=120, show_spinner=False)
def obtener_estados_actuales():
    estados = {}
    try:
        sheet_int = get_sheet("intervenciones")
        if sheet_int:
            data_int = sheet_int.get_all_values()
            for row in data_int:
                if len(row) >= 18: estados[row[0]] = row[17]
        sheet = get_sheet("estados_equipos")
        if sheet:
            data = sheet.get_all_values()
            for row in data:
                if len(row) >= 2: estados[row[0]] = row[1] 
    except: pass
    return estados

def actualizar_estado_equipo_en_nube(tag, nuevo_estado):
    try:
        sheet = get_sheet("estados_equipos")
        if sheet:
            registros = sheet.get_all_values(); fila_encontrada = -1
            for i, fila in enumerate(registros):
                if len(fila) > 0 and fila[0] == tag: fila_encontrada = i + 1; break
            if fila_encontrada != -1: sheet.update_cell(fila_encontrada, 2, nuevo_estado)
            else: sheet.append_row([tag, nuevo_estado])
            st.cache_data.clear() 
    except Exception as e: pass

@st.cache_data(ttl=120, show_spinner=False)
def obtener_datos_equipo(tag):
    datos = {}; sheet = get_sheet("datos_equipo")
    if sheet:
        data = sheet.get_all_values()
        for r in data:
            if len(r) >= 3 and r[0] == tag: datos[r[1]] = r[2]
    return datos

@st.cache_data(ttl=120, show_spinner=False)
def obtener_observaciones(tag):
    sheet = get_sheet("observaciones")
    if sheet:
        data = sheet.get_all_values()
        obs = [{"id": r[0], "fecha": r[2], "usuario": r[3], "texto": r[4]} for r in data if len(r) >= 6 and r[1] == tag and r[5] == "ACTIVO"]
        if obs: return pd.DataFrame(obs).iloc[::-1]
    return pd.DataFrame(columns=["id", "fecha", "usuario", "texto"])

@st.cache_data(ttl=120, show_spinner=False)
def obtener_contactos():
    sheet = get_sheet("contactos")
    if sheet:
        data = sheet.get_all_values()
        contactos = [r[0] for r in data if len(r) > 1 and r[1] == "ACTIVO"]
        if contactos: return sorted(list(set(contactos)))
    return ["Lorena Rojas"]

@st.cache_data(ttl=120, show_spinner=False)
def obtener_tecnicos():
    sheet = get_sheet("tecnicos")
    if sheet:
        data = sheet.get_all_values()
        tecnicos = [r[0] for r in data if len(r) > 1 and r[1] == "ACTIVO"]
        if tecnicos: return sorted(list(set(tecnicos)))
    return [st.session_state.get('usuario_actual', '').title()] if st.session_state.get('usuario_actual') else ["Ignacio Morales"]

@st.cache_data(ttl=120, show_spinner=False)
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

@st.cache_data(ttl=120, show_spinner=False)
def buscar_ultimo_registro(tag):
    sheet = get_sheet("intervenciones")
    if sheet:
        data = sheet.get_all_values()
        for row in reversed(data):
            if len(row) >= 20 and row[0] == tag: return (row[5], row[6], row[9], row[14], row[15], row[7], row[8], row[10], row[11], row[12], row[13], row[16], row[17])
    return None

@st.cache_data(ttl=120, show_spinner=False)
def obtener_todo_el_historial(tag):
    try:
        sheet = get_sheet("intervenciones")
        if sheet:
            data = sheet.get_all_values()
            hist = [{"fecha": r[5], "tipo_intervencion": r[15], "estado_equipo": r[17], "Cuenta Usuario": r[19], "horas_marcha": r[12], "p_carga": r[10], "temp_salida": r[9]} for r in data if len(r) >= 20 and r[0] == tag]
            if hist: return pd.DataFrame(hist).iloc[::-1]
    except: pass
    return pd.DataFrame()

@st.cache_data(ttl=120, show_spinner=False)
def obtener_historial_global():
    try:
        sheet = get_sheet("intervenciones")
        if sheet:
            data = sheet.get_all_values()
            hist = []
            for r in reversed(data):
                if len(r) >= 20 and r[0] != "TAG":
                    hist.append({
                        "tag": r[0], "modelo": r[1], "area": r[3], 
                        "fecha": r[5], "tecnico": r[7], "tipo": r[15], 
                        "estado": r[17], "condicion": r[14], "reco": r[16]
                    })
                    if len(hist) >= 50: break
            return hist
    except: pass
    return []

def eliminar_registro_intervencion(tag, fecha, tipo):
    try:
        sheet = get_sheet("intervenciones")
        if sheet:
            records = sheet.get_all_values()
            for i in range(len(records)-1, 0, -1):
                row = records[i]
                if row[0] == tag and row[5] == fecha and row[15] == tipo:
                    sheet.delete_rows(i + 1)
                    st.cache_data.clear()
                    return True
    except Exception as e: pass
    return False

def guardar_dato_equipo(tag, clave, valor):
    sheet = get_sheet("datos_equipo")
    if sheet: sheet.append_row([tag, clave, valor]); st.cache_data.clear()

def guardar_registro(data_tuple):
    for intento in range(3):
        try:
            sheet = get_sheet("intervenciones")
            if sheet is not None:
                row = [str(x) for x in data_tuple]
                sheet.insert_row(row, index=len(sheet.get_all_values()) + 1)
                st.cache_data.clear(); return True
        except: time.sleep(5)
    return False

def agregar_observacion(tag, usuario, texto):
    if not texto.strip(): return
    fecha_actual = pd.Timestamp.now().strftime("%d/%m/%Y %H:%M"); id_obs = str(uuid.uuid4())[:8]
    sheet = get_sheet("observaciones")
    if sheet: sheet.append_row([id_obs, tag, fecha_actual, usuario.title(), texto.strip(), "ACTIVO"]); st.cache_data.clear()

def eliminar_observacion(id_obs):
    sheet = get_sheet("observaciones")
    if sheet:
        cell = sheet.find(id_obs)
        if cell: sheet.update_cell(cell.row, 6, "ELIMINADO"); st.cache_data.clear()

def agregar_contacto(nombre):
    if not nombre.strip(): return
    sheet = get_sheet("contactos")
    if sheet: sheet.append_row([nombre.strip().title(), "ACTIVO"]); st.cache_data.clear()

def eliminar_contacto(nombre):
    sheet = get_sheet("contactos")
    if sheet:
        cells = sheet.findall(nombre)
        for cell in cells: sheet.update_cell(cell.row, 2, "ELIMINADO"); st.cache_data.clear()

def agregar_tecnico(nombre):
    if not nombre.strip(): return
    sheet = get_sheet("tecnicos")
    if sheet: sheet.append_row([nombre.strip().title(), "ACTIVO"]); st.cache_data.clear()

def eliminar_tecnico(nombre):
    sheet = get_sheet("tecnicos")
    if sheet:
        cells = sheet.findall(nombre)
        for cell in cells: sheet.update_cell(cell.row, 2, "ELIMINADO"); st.cache_data.clear()

def guardar_especificacion_db(modelo, clave, valor):
    sheet = get_sheet("especificaciones")
    if sheet: sheet.append_row([modelo, clave, valor]); st.cache_data.clear()
    # =============================================================================
# 4. FUNCIONES AUXILIARES GLOBALES Y CEREBRO DE FECHAS
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

def obtener_fecha_hoy_esp():
    meses = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
    ahora = pd.Timestamp.now()
    return f"{ahora.day} de {meses[ahora.month]} de {ahora.year}"

def parse_fecha(f_str):
    try:
        s = str(f_str).lower().strip()
        meses = {"ene":1,"feb":2,"mar":3,"abr":4,"may":5,"jun":6,"jul":7,"ago":8,"sep":9,"oct":10,"nov":11,"dic":12}
        m1 = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)
        if m1: return datetime.date(int(m1.group(1)), int(m1.group(2)), int(m1.group(3)))
        m2 = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})", s)
        if m2: return datetime.date(int(m2.group(3)), int(m2.group(2)), int(m2.group(1)))
        nums = re.findall(r'\d+', s); words = re.findall(r'[a-z]+', s)
        if not nums: return datetime.date(1970,1,1)
        day = int(nums[0])
        year = int(nums[-1]) if len(nums)>1 else datetime.date.today().year
        if year < 100: year += 2000
        month = next((meses[w[:3]] for w in words if w[:3] in meses), 1)
        return datetime.date(year, month, day if day<=31 else year)
    except: return datetime.date(1970, 1, 1)

def format_fecha(d):
    meses_nombres = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
    return f"{d.day} de {meses_nombres[d.month]} de {d.year}" if d.year != 1970 else "Fecha Desconocida"

def cargar_pendientes(usuario):
    archivo = os.path.join(RUTA_ONEDRIVE, f"bandeja_{usuario.replace(' ', '_')}.json")
    if os.path.exists(archivo):
        try:
            with open(archivo, "r", encoding="utf-8") as f: return json.load(f)
        except: return []
    return []

def guardar_pendientes(usuario, pendientes):
    os.makedirs(RUTA_ONEDRIVE, exist_ok=True)
    archivo = os.path.join(RUTA_ONEDRIVE, f"bandeja_{usuario.replace(' ', '_')}.json")
    try:
        with open(archivo, "w", encoding="utf-8") as f: json.dump(pendientes, f, ensure_ascii=False, indent=4)
    except: pass

def wk_to_date(wk_string):
    try:
        s = str(wk_string).strip().upper()
        year_match = re.search(r'(202\d)', s)
        y = int(year_match.group(1)) if year_match else None
        nums = re.findall(r'\d+', s)
        wk_num = int(nums[0])
        if not y:
            y = 2025 if wk_num >= 50 else 2026
        return datetime.date.fromisocalendar(y, wk_num, 1)
    except: return None

def calcular_mes_minero(wk_string):
    if pd.isna(wk_string) or str(wk_string).strip() == "": return "Sin Asignar"
    d = wk_to_date(wk_string)
    if not d: return "Sin Asignar"
    meses_full = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    m_idx = d.month - 1 if d.day <= 15 else (d.month if d.month < 12 else 0)
    y = d.year
    if d.month == 12 and d.day > 15: y += 1
    return f"{meses_full[m_idx]} {y}"

def get_current_wk():
    hoy = datetime.date.today()
    return f"WK{hoy.isocalendar()[1]:02d}_{hoy.year}"

def formatear_wk(wk_str):
    if pd.isna(wk_str) or str(wk_str).strip() == "": return ""
    s = str(wk_str).strip().upper()
    nums = re.findall(r'\d+', s)
    if not nums: return s
    wk = int(nums[0])
    if len(nums) > 1 and int(nums[-1]) > 2000:
        return f"WK{wk:02d}_{nums[-1]}"
    return f"WK{wk:02d}"

def get_semanas_mes_minero(mes_nombre):
    if mes_nombre == "Todas" or mes_nombre == "Sin Asignar": return "Todas"
    try:
        parts = mes_nombre.split(" ")
        m_str, y_num = parts[0], int(parts[1])
        meses_map_full = {"Enero": 1, "Febrero": 2, "Marzo": 3, "Abril": 4, "Mayo": 5, "Junio": 6, "Julio": 7, "Agosto": 8, "Septiembre": 9, "Octubre": 10, "Noviembre": 11, "Diciembre": 12}
        m_num = meses_map_full[m_str]
        if m_num == 1: min_d = datetime.date(y_num - 1, 12, 16)
        else: min_d = datetime.date(y_num, m_num - 1, 16)
        max_d = datetime.date(y_num, m_num, 15)
        return f"WK{min_d.isocalendar()[1]:02d} a WK{max_d.isocalendar()[1]:02d}"
    except: return ""

def safe_date_str(x):
    try: return x[:10] if isinstance(x, str) else x.strftime("%Y-%m-%d")
    except: return ""

# =============================================================================
# 5. MOTOR PLANIFICACIÓN
# =============================================================================
DATOS_PLAN_BASE = [
    {"TAG": "70-GC-013", "S_Programada": "WK51_2025", "Tipo": "P2", "Estado": "Hecho", "S_Realizada": "2025-12-15", "Observacion": ""},
    {"TAG": "70-GC-013", "S_Programada": "WK02_2026", "Tipo": "INSP", "Estado": "Hecho", "S_Realizada": "2026-01-05", "Observacion": ""},
    {"TAG": "70-GC-013", "S_Programada": "WK04_2026", "Tipo": "P1", "Estado": "Hecho", "S_Realizada": "2026-01-19", "Observacion": ""},
    {"TAG": "70-GC-013", "S_Programada": "WK07_2026", "Tipo": "P1", "Estado": "Hecho", "S_Realizada": "2026-02-10", "Observacion": ""},
    {"TAG": "70-GC-013", "S_Programada": "WK11_2026", "Tipo": "INSP", "Estado": "Pendiente", "S_Realizada": "", "Observacion": ""},
    {"TAG": "70-GC-014", "S_Programada": "WK52_2025", "Tipo": "INSP", "Estado": "Hecho", "S_Realizada": "2025-12-22", "Observacion": ""},
    {"TAG": "70-GC-014", "S_Programada": "WK02_2026", "Tipo": "P2", "Estado": "Hecho", "S_Realizada": "2026-01-05", "Observacion": ""},
    {"TAG": "70-GC-014", "S_Programada": "WK04_2026", "Tipo": "INSP", "Estado": "F/S", "S_Realizada": "", "Observacion": ""}, 
    {"TAG": "70-GC-014", "S_Programada": "WK09_2026", "Tipo": "INSP", "Estado": "Hecho", "S_Realizada": "2026-02-23", "Observacion": ""},
    {"TAG": "70-GC-014", "S_Programada": "WK10_2026", "Tipo": "INSP", "Estado": "Pendiente", "S_Realizada": "", "Observacion": ""},
    {"TAG": "50-GC-001", "S_Programada": "WK01_2026", "Tipo": "P2", "Estado": "Hecho", "S_Realizada": "2025-12-29", "Observacion": ""},
    {"TAG": "50-GC-001", "S_Programada": "WK04_2026", "Tipo": "P1", "Estado": "Hecho", "S_Realizada": "2026-01-21", "Observacion": ""},
    {"TAG": "50-GC-001", "S_Programada": "WK09_2026", "Tipo": "P1", "Estado": "Hecho", "S_Realizada": "2026-02-23", "Observacion": ""},
    {"TAG": "50-GC-001", "S_Programada": "WK10_2026", "Tipo": "P3", "Estado": "Pendiente", "S_Realizada": "", "Observacion": ""},
    {"TAG": "50-GC-002", "S_Programada": "WK01_2026", "Tipo": "INSP", "Estado": "Hecho", "S_Realizada": "2025-12-29", "Observacion": ""},
    {"TAG": "50-GC-002", "S_Programada": "WK02_2026", "Tipo": "P2", "Estado": "F/S", "S_Realizada": "", "Observacion": ""},
    {"TAG": "50-GC-002", "S_Programada": "WK04_2026", "Tipo": "INSP", "Estado": "Hecho", "S_Realizada": "2026-01-19", "Observacion": ""},
    {"TAG": "50-GC-002", "S_Programada": "WK09_2026", "Tipo": "INSP", "Estado": "Pendiente", "S_Realizada": "", "Observacion": ""},
    {"TAG": "50-GC-003", "S_Programada": "WK01_2026", "Tipo": "P2", "Estado": "Hecho", "S_Realizada": "2025-12-29", "Observacion": ""},
    {"TAG": "50-GC-003", "S_Programada": "WK07_2026", "Tipo": "P1", "Estado": "F/S", "S_Realizada": "", "Observacion": ""},
    {"TAG": "50-GC-003", "S_Programada": "WK11_2026", "Tipo": "P1", "Estado": "Pendiente", "S_Realizada": "", "Observacion": ""},
    {"TAG": "55-GC-015", "S_Programada": "WK01_2026", "Tipo": "P2", "Estado": "Hecho", "S_Realizada": "2025-12-29", "Observacion": ""},
    {"TAG": "55-GC-015", "S_Programada": "WK06_2026", "Tipo": "P1", "Estado": "Hecho", "S_Realizada": "2026-02-04", "Observacion": ""},
    {"TAG": "55-GC-015", "S_Programada": "WK08_2026", "Tipo": "INSP", "Estado": "Hecho", "S_Realizada": "2026-02-16", "Observacion": ""},
    {"TAG": "65-GC-011", "S_Programada": "WK01_2026", "Tipo": "P3", "Estado": "Hecho", "S_Realizada": "2025-12-29", "Observacion": ""},
    {"TAG": "65-GC-011", "S_Programada": "WK05_2026", "Tipo": "P1", "Estado": "Hecho", "S_Realizada": "2026-01-28", "Observacion": ""},
    {"TAG": "65-GC-011", "S_Programada": "WK11_2026", "Tipo": "INSP", "Estado": "Hecho", "S_Realizada": "2026-03-09", "Observacion": ""},
    {"TAG": "35-GC-006", "S_Programada": "WK01_2026", "Tipo": "P3", "Estado": "Hecho", "S_Realizada": "2025-12-29", "Observacion": ""},
    {"TAG": "35-GC-006", "S_Programada": "WK02_2026", "Tipo": "P1", "Estado": "F/S", "S_Realizada": "", "Observacion": ""},
    {"TAG": "35-GC-006", "S_Programada": "WK08_2026", "Tipo": "INSP", "Estado": "Hecho", "S_Realizada": "2026-02-16", "Observacion": ""}
]

def generar_plan_futuro():
    PATRONES = {
        "P4_I_P1": ["P4", "INSP", "P1", "INSP", "P2", "INSP", "P1", "INSP", "P3", "INSP"],
        "I_P3_I_P1": ["INSP", "P3", "INSP", "P1", "INSP", "P2", "INSP", "P1", "INSP", "P4"],
        "P3_I_P1": ["P3", "INSP", "P1", "INSP", "P2", "INSP", "P1", "INSP", "P3", "INSP"],
        "I_P3_I_P1_alt": ["INSP", "P3", "INSP", "P1", "INSP", "P2", "INSP", "P1", "INSP", "P3"],
        "I_P4_I_P1": ["INSP", "P4", "INSP", "P1", "INSP", "P2", "INSP", "P1", "INSP", "P3"],
        "I_I_I_P2": ["INSP", "INSP", "INSP", "P2", "INSP", "INSP", "INSP", "INSP", "INSP", "P4"],
        "P1_I_P2": ["P1", "INSP", "P2", "INSP", "P1", "INSP", "P1", "INSP", "P4", "INSP"],
        "I_P1_I_P1": ["INSP", "P1", "INSP", "P1", "INSP", "P2", "INSP", "P1", "INSP", "P1"],
        "I_P2_I_I": ["INSP", "P2", "INSP", "INSP", "INSP", "INSP", "INSP", "P2", "INSP", "INSP"],
        "P1_P1_P2": ["P1", "P1", "P2", "P1", "P1", "P4", "P1", "P1", "P2", "P1"],
        "P2_P1_P1": ["P2", "P1", "P1", "P1", "P1", "P1", "P4", "P1", "P1", "P2"],
        "P1_P2_P1": ["P1", "P2", "P1", "P1", "P4", "P1", "P1", "P2", "P1", "P1"],
        "I_P4_I_I": ["INSP", "P4", "INSP", "INSP", "P1", "INSP", "INSP", "P2", "INSP", "INSP"]
    }
    MAP = {
        "20-GC-001": "P4_I_P1", "20-GC-002": "I_P3_I_P1", "20-GC-003": "P3_I_P1", "20-GC-004": "I_P3_I_P1_alt",
        "35-GC-006": "P4_I_P1", "35-GC-007": "I_P4_I_P1", "35-GC-008": "I_I_I_P2", "50-CD-001": "I_I_I_P2", 
        "50-CD-002": "P4_I_P1", "50-GC-001": "P1_I_P2", "50-GC-002": "I_P1_I_P1", "50-GC-003": "I_P2_I_I", 
        "50-GC-004": "I_P2_I_I", "55-GC-015": "P1_P1_P2", "65-CD-011": "P2_P1_P1", "65-CD-012": "P1_P2_P1", 
        "65-GC-009": "I_P4_I_I", "65-GC-011": "P4_I_P1", "70-GC-013": "P4_I_P1", "70-GC-014": "P4_I_P1", "Taller": "P4_I_P1"
    }
    WKS = ["WK15_2026", "WK19_2026", "WK24_2026", "WK28_2026", "WK32_2026", "WK37_2026", "WK41_2026", "WK45_2026", "WK50_2026", "WK02_2027"]
    
    datos = []
    for tag, p in MAP.items():
        for wk, tipo in zip(WKS, PATRONES[p]):
            datos.append({"TAG": tag, "S_Programada": wk, "Tipo": tipo, "Estado": "Pendiente", "S_Realizada": "", "Observacion": ""})
    return pd.DataFrame(datos)

@st.cache_data(ttl=60, show_spinner=False)
def cargar_cmms():
    headers = ["TAG", "S_Programada", "Tipo", "Estado", "S_Realizada", "Observacion"]
    try:
        sheet = get_sheet("plan_cmms")
        if sheet:
            data = sheet.get_all_values()
            if len(data) > 0:
                df = pd.DataFrame(data[1:], columns=data[0]) if len(data) > 1 else pd.DataFrame(columns=data[0])
                
                if df.empty or "S_Programada" not in df.columns:
                    df_base = pd.DataFrame(DATOS_PLAN_BASE, columns=headers)
                    sheet.clear()
                    sheet.append_rows([headers] + df_base.values.tolist())
                    st.cache_data.clear()
                    return df_base
                    
                if "S_Programada" in df.columns:
                    def migrar_fechas(row):
                        val = str(row.get('S_Realizada', ''))
                        if 'WK' in val.upper():
                            d = wk_to_date(row['S_Programada'])
                            return d.strftime("%Y-%m-%d") if d else ""
                        return val
                    df['S_Realizada'] = df.apply(migrar_fechas, axis=1)
                    
                    def clean_db_values(val):
                        if pd.isnull(val): return ''
                        return re.sub(r'^[^\w\s\/-]+', '', str(val)).strip()
                        
                    if "Tipo" in df.columns: df["Tipo"] = df["Tipo"].apply(clean_db_values)
                    if "Estado" in df.columns: df["Estado"] = df["Estado"].apply(clean_db_values)
                    
                    return df
            else:
                df_base = pd.DataFrame(DATOS_PLAN_BASE, columns=headers)
                sheet.append_rows([headers] + df_base.values.tolist())
                st.cache_data.clear()
                return df_base
    except Exception as e: print(f"Error cargando Planificación: {e}")
    return pd.DataFrame(DATOS_PLAN_BASE, columns=headers)

def guardar_cmms(df):
    sheet = get_sheet("plan_cmms")
    if sheet:
        df_clean = df.copy().fillna("").astype(str).replace(["nan", "NaN", "NaT", "None", "<NA>"], "")
        sheet.clear()
        sheet.append_rows([df_clean.columns.values.tolist()] + df_clean.values.tolist())
        st.cache_data.clear()

def seleccionar_equipo(tag):
    st.session_state.equipo_seleccionado = tag
    st.session_state.vista_firmas = False
    reg = buscar_ultimo_registro(tag)
    
    if reg:
        st.session_state.input_cliente = reg[1]
        st.session_state.input_tec1 = reg[5]
        st.session_state.input_tec2 = reg[6]
        st.session_state.input_estado_eq = reg[12] if reg[12] else "Operativo"
        st.session_state.input_h_marcha = int(reg[9]) if reg[9] else 0
        st.session_state.input_h_carga = int(reg[10]) if reg[10] else 0
        st.session_state.input_temp = str(reg[2]).replace(',', '.') if reg[2] is not None else "70.0"
        try: st.session_state.input_p_carga = str(reg[7]).split()[0].replace(',', '.')
        except: st.session_state.input_p_carga = "7.0"
        try: st.session_state.input_p_descarga = str(reg[8]).split()[0].replace(',', '.')
        except: st.session_state.input_p_descarga = "7.5"
        
        st.session_state.input_estado = str(reg[3]) if reg[3] else ""
        st.session_state.input_reco = str(reg[11]) if reg[11] else ""
    else: 
        st.session_state.update({'input_estado_eq': "Operativo", 'input_estado': "", 'input_reco': ""})

def volver_catalogo(): 
    st.session_state.equipo_seleccionado = None
    st.session_state.vista_firmas = False
    st.session_state.vista_actual = "catalogo"