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
# =============================================================================
# 3. CONVERSIÓN A PDF, FECHAS EN ESPAÑOL Y BANDEJAS PRIVADAS
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

def obtener_fecha_hoy_esp():
    meses = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
    ahora = pd.Timestamp.now()
    return f"{ahora.day} de {meses[ahora.month]} de {ahora.year}"

def obtener_quincena_actual():
    hoy = datetime.date.today()
    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    if hoy.day < 15:
        mes_plan = meses[hoy.month - 1]
        inicio = f"15 de {meses[hoy.month - 2 if hoy.month > 1 else 11]}"
        fin = f"15 de {mes_plan}"
    else:
        mes_plan = meses[hoy.month] if hoy.month < 12 else "Enero"
        inicio = f"15 de {meses[hoy.month - 1]}"
        fin = f"15 de {mes_plan}"
    return mes_plan, f"{inicio} al {fin}"

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

# =============================================================================
# 3.1 BASE DE DATOS: PLANIFICACIÓN EN MATRIZ (CONECTADA A GOOGLE SHEETS)
# =============================================================================
def generar_planificacion_base():
    meses = ["15c Ene", "15c Feb", "15c Mar", "15c Abr", "15c May", "15c Jun", "15c Jul", "15c Ago", "15c Sep", "15c Oct", "15c Nov", "15c Dic"]
    datos = [
        {"TAG": "70-GC-013", "Equipo": "GA 132", "Área": "Descarga Acido", "15c Ene": "INSP", "15c Feb": "P1\nHecho WK7", "15c Mar": "INSP\nHecho W10", "15c Abr": "P4", "15c May": "INSP", "15c Jun": "P1", "15c Jul": "INSP", "15c Ago": "P2", "15c Sep": "INSP", "15c Oct": "P1", "15c Nov": "INSP", "15c Dic": "P3"},
        {"TAG": "70-GC-014", "Equipo": "GA 132", "Área": "Descarga Acido", "15c Ene": "P2\nLista", "15c Feb": "INSP\nFalta", "15c Mar": "P1\nHecho W10", "15c Abr": "INSP", "15c May": "P3", "15c Jun": "INSP", "15c Jul": "P1", "15c Ago": "INSP", "15c Sep": "P2", "15c Oct": "INSP", "15c Nov": "P1", "15c Dic": "INSP"},
        {"TAG": "50-GC-001", "Equipo": "GA 45", "Área": "Planta SX", "15c Ene": "INSP", "15c Feb": "P1\nHecho WK4", "15c Mar": "INSP\nPdte W10", "15c Abr": "P3", "15c May": "INSP", "15c Jun": "P1", "15c Jul": "INSP", "15c Ago": "P2", "15c Sep": "INSP", "15c Oct": "P1", "15c Nov": "INSP", "15c Dic": "P3"},
        {"TAG": "50-GC-002", "Equipo": "GA 45", "Área": "Planta SX", "15c Ene": "P2\nFalta kit", "15c Feb": "INSP\nHecho WK4", "15c Mar": "P1\nPdte W10", "15c Abr": "INSP", "15c May": "P3", "15c Jun": "INSP", "15c Jul": "P1", "15c Ago": "INSP", "15c Sep": "P2", "15c Oct": "INSP", "15c Nov": "P1", "15c Dic": "INSP"},
        {"TAG": "50-GC-003", "Equipo": "ZT 37", "Área": "Planta SX", "15c Ene": "INSP", "15c Feb": "P1\nF/S WK7", "15c Mar": "INSP\nF/S WK9", "15c Abr": "P4", "15c May": "INSP", "15c Jun": "P1", "15c Jul": "INSP", "15c Ago": "P2", "15c Sep": "INSP", "15c Oct": "P1", "15c Nov": "INSP", "15c Dic": "P3"},
        {"TAG": "50-GC-004", "Equipo": "ZT 37", "Área": "Planta SX", "15c Ene": "P2\nLista", "15c Feb": "INSP", "15c Mar": "INSP\nF/S WK8", "15c Abr": "INSP", "15c May": "P4", "15c Jun": "INSP", "15c Jul": "P1", "15c Ago": "INSP", "15c Sep": "P2", "15c Oct": "INSP", "15c Nov": "P1", "15c Dic": "INSP"},
        {"TAG": "50-CD-001", "Equipo": "CD 80+", "Área": "Planta SX", "15c Ene": "P4\nFalta", "15c Feb": "INSP", "15c Mar": "INSP\nWK8", "15c Abr": "INSP", "15c May": "INSP", "15c Jun": "INSP", "15c Jul": "P2", "15c Ago": "INSP", "15c Sep": "INSP", "15c Oct": "INSP", "15c Nov": "INSP", "15c Dic": "INSP"},
        {"TAG": "50-CD-002", "Equipo": "CD 80+", "Área": "Planta SX", "15c Ene": "P4\nFalta", "15c Feb": "INSP", "15c Mar": "INSP\nWK8", "15c Abr": "INSP", "15c May": "INSP", "15c Jun": "INSP", "15c Jul": "P2", "15c Ago": "INSP", "15c Sep": "INSP", "15c Oct": "INSP", "15c Nov": "INSP", "15c Dic": "INSP"},
        {"TAG": "55-GC-015", "Equipo": "GA 30", "Área": "Planta Borra", "15c Ene": "INSP", "15c Feb": "P1\nHecho WK6", "15c Mar": "INSP\nWK11", "15c Abr": "P4", "15c May": "INSP", "15c Jun": "P1", "15c Jul": "INSP", "15c Ago": "P2", "15c Sep": "INSP", "15c Oct": "P1", "15c Nov": "INSP", "15c Dic": "P3"},
        {"TAG": "65-GC-011", "Equipo": "GA 250", "Área": "Patio Estanques", "15c Ene": "INSP", "15c Feb": "P1\nHecho WK5", "15c Mar": "INSP\nWK11", "15c Abr": "P1", "15c May": "INSP", "15c Jun": "P2", "15c Jul": "INSP", "15c Ago": "P1", "15c Sep": "INSP", "15c Oct": "P1", "15c Nov": "INSP", "15c Dic": "P4"},
        {"TAG": "65-GC-009", "Equipo": "GA 250", "Área": "Patio Estanques", "15c Ene": "P1\nFalta Kit", "15c Feb": "INSP", "15c Mar": "P4\nWK8", "15c Abr": "INSP", "15c May": "P1", "15c Jun": "INSP", "15c Jul": "P1", "15c Ago": "INSP", "15c Sep": "P2", "15c Oct": "INSP", "15c Nov": "P1", "15c Dic": "INSP"},
        {"TAG": "65-CD-011", "Equipo": "CD 630", "Área": "Patio Estanques", "15c Ene": "INSP", "15c Feb": "P2\nFalta Kit", "15c Mar": "INSP\nWK8", "15c Abr": "INSP", "15c May": "P2", "15c Jun": "INSP", "15c Jul": "INSP", "15c Ago": "P2", "15c Sep": "INSP", "15c Oct": "INSP", "15c Nov": "P2", "15c Dic": "INSP"},
        {"TAG": "65-CD-012", "Equipo": "CD 630", "Área": "Patio Estanques", "15c Ene": "INSP", "15c Feb": "P2\nFalta Kit", "15c Mar": "INSP\nWK8", "15c Abr": "INSP", "15c May": "P2", "15c Jun": "INSP", "15c Jul": "INSP", "15c Ago": "P2", "15c Sep": "INSP", "15c Oct": "INSP", "15c Nov": "P2", "15c Dic": "INSP"},
        {"TAG": "35-GC-006", "Equipo": "GA 250", "Área": "Chancado Sec.", "15c Ene": "P1\nFalta kit", "15c Feb": "P1\nF/S", "15c Mar": "P2\nF/S WK11", "15c Abr": "P1", "15c May": "P1", "15c Jun": "P2", "15c Jul": "P1", "15c Ago": "P1", "15c Sep": "P4", "15c Oct": "P1", "15c Nov": "P1", "15c Dic": "P2"},
        {"TAG": "35-GC-007", "Equipo": "GA 250", "Área": "Chancado Sec.", "15c Ene": "P3\nHecho", "15c Feb": "P1\nHecho W6", "15c Mar": "P1\nWK11", "15c Abr": "P2", "15c May": "P1", "15c Jun": "P1", "15c Jul": "P2", "15c Ago": "P1", "15c Sep": "P1", "15c Oct": "P4", "15c Nov": "P1", "15c Dic": "P1"},
        {"TAG": "35-GC-008", "Equipo": "GA 250", "Área": "Chancado Sec.", "15c Ene": "P1\nFalta kit", "15c Feb": "P2\nHecho W6", "15c Mar": "P1\nWK11", "15c Abr": "P1", "15c May": "P2", "15c Jun": "P1", "15c Jul": "P1", "15c Ago": "P4", "15c Sep": "P1", "15c Oct": "P1", "15c Nov": "P2", "15c Dic": "P1"},
        {"TAG": "20-GC-004", "Equipo": "GA 37", "Área": "Truck Shop", "15c Ene": "INSP", "15c Feb": "P1\nFalta WK5", "15c Mar": "P1\nWK10", "15c Abr": "INSP", "15c May": "P4", "15c Jun": "INSP", "15c Jul": "INSP", "15c Ago": "P1", "15c Sep": "INSP", "15c Oct": "INSP", "15c Nov": "P2", "15c Dic": "INSP"},
        {"TAG": "20-GC-001", "Equipo": "GA 75", "Área": "Truck Shop", "15c Ene": "INSP", "15c Feb": "P1\nHecho WK4", "15c Mar": "INSP\nWK10", "15c Abr": "P4", "15c May": "INSP", "15c Jun": "P1", "15c Jul": "INSP", "15c Ago": "P2", "15c Sep": "INSP", "15c Oct": "P1", "15c Nov": "INSP", "15c Dic": "P3"},
        {"TAG": "20-GC-002", "Equipo": "GA 75", "Área": "Truck Shop", "15c Ene": "INSP", "15c Feb": "P1\nFalta WK4", "15c Mar": "INSP\nWK10", "15c Abr": "P4", "15c May": "INSP", "15c Jun": "P1", "15c Jul": "INSP", "15c Ago": "P2", "15c Sep": "INSP", "15c Oct": "P1", "15c Nov": "INSP", "15c Dic": "P3"},
        {"TAG": "20-GC-003", "Equipo": "GA 90", "Área": "Truck Shop", "15c Ene": "INSP", "15c Feb": "P1\nHecho W7", "15c Mar": "INSP\nWK10", "15c Abr": "P4", "15c May": "INSP", "15c Jun": "P1", "15c Jul": "INSP", "15c Ago": "P2", "15c Sep": "INSP", "15c Oct": "P1", "15c Nov": "INSP", "15c Dic": "P3"},
        {"TAG": "Taller", "Equipo": "GA 18", "Área": "Taller", "15c Ene": "INSP", "15c Feb": "P2\nHecho WK5", "15c Mar": "INSP", "15c Abr": "INSP", "15c May": "INSP", "15c Jun": "INSP", "15c Jul": "INSP", "15c Ago": "INSP", "15c Sep": "INSP", "15c Oct": "INSP", "15c Nov": "INSP", "15c Dic": "INSP"}
    ]
    return pd.DataFrame(datos)

@st.cache_data(ttl=60)
def cargar_planificacion():
    try:
        sheet = get_sheet("planificacion")
        if sheet:
            data = sheet.get_all_values()
            if len(data) > 1:
                df = pd.DataFrame(data[1:], columns=data[0])
                return df
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

# =============================================================================
# ESTRATEGIA VISUAL DE COLORES (MATRIZ Y TICKETS)
# =============================================================================
def estilo_dinamico_celdas(val):
    if pd.isna(val) or val == "": return ''
    v = str(val).upper()
    base_css = 'white-space: pre-wrap; line-height: 1.4; border-radius: 6px; padding: 6px; text-align: center; '
    
    if 'F/S' in v or 'FUERA' in v: return base_css + 'background-color: rgba(255, 23, 68, 0.25); color: #ff1744; font-weight: bold; border-left: 4px solid #ff1744;'
    if 'HECHO' in v or 'LISTO' in v or 'OK' in v: return base_css + 'background-color: rgba(0, 230, 118, 0.25); color: #00e676; font-weight: bold; border-left: 4px solid #00e676;'
    if any(x in v for x in ['FALTA', 'PENDIENTE', 'WK', 'PEND', 'LUNES', 'MARTES', 'MIÉRCOLES', 'MIERCOLES', 'JUEVES']): 
        return base_css + 'background-color: rgba(255, 193, 7, 0.25); color: #FFC107; font-weight: bold; border-left: 4px solid #FFC107;'
    
    if 'P1' in v: return base_css + 'background-color: rgba(0, 191, 255, 0.15); color: #00BFFF; font-weight: bold;'
    if 'P2' in v: return base_css + 'background-color: rgba(255, 152, 0, 0.15); color: #FF9800; font-weight: bold;'
    if 'P3' in v: return base_css + 'background-color: rgba(156, 39, 176, 0.15); color: #9C27B0; font-weight: bold;'
    if 'P4' in v: return base_css + 'background-color: rgba(244, 67, 54, 0.15); color: #F44336; font-weight: bold;'
    if 'INSP' in v or v == 'I': return base_css + 'color: #8c9eb5; font-style: italic;'
    return base_css

def estilo_simple_editor(val):
    if pd.isna(val) or val == "": return ''
    v = str(val).upper()
    if 'F/S' in v or 'FUERA' in v: return 'background-color: #ff1744; color: white;'
    if 'HECHO' in v or 'LISTO' in v or 'OK' in v: return 'background-color: #00e676; color: black;'
    if any(x in v for x in ['FALTA', 'PENDIENTE', 'WK', 'PEND', 'LUNES', 'MARTES', 'MIÉRCOLES', 'MIERCOLES', 'JUEVES']): return 'background-color: #FFC107; color: black;'
    if 'P1' in v: return 'background-color: #003b5c; color: #00BFFF;' 
    if 'P2' in v: return 'background-color: #5c3700; color: #FF9800;'
    if 'P3' in v: return 'background-color: #430c4d; color: #e166ff;'
    if 'P4' in v: return 'background-color: #5c0e0e; color: #ff6e6e;'
    if 'INSP' in v or v == 'I': return 'color: #8c9eb5;'
    return ''

def estilo_pautas_puras(val):
    """Estilo exclusivo de "Badges" para la columna de Intervención en los tickets."""
    v = str(val).upper()
    if 'P1' == v: return 'background-color: #00BFFF; color: white; font-weight: bold; text-align: center; border-radius: 4px;'
    if 'P2' == v: return 'background-color: #FF9800; color: white; font-weight: bold; text-align: center; border-radius: 4px;'
    if 'P3' == v: return 'background-color: #9C27B0; color: white; font-weight: bold; text-align: center; border-radius: 4px;'
    if 'P4' == v: return 'background-color: #F44336; color: white; font-weight: bold; text-align: center; border-radius: 4px;'
    if 'INSP' in v or 'I' == v: return 'background-color: transparent; color: #8c9eb5; font-weight: bold; text-align: center; border: 1px dashed #8c9eb5; border-radius: 4px;'
    return ''

# =============================================================================
# 4. INICIALIZACIÓN DE VARIABLES DE SESIÓN
# =============================================================================
ESPECIFICACIONES = obtener_especificaciones(DEFAULT_SPECS)
default_states = {
    'logged_in': False, 'usuario_actual': "", 'equipo_seleccionado': None, 'vista_actual': "catalogo",
    'input_cliente': "Lorena Rojas", 'input_tec1': "Ignacio Morales", 'input_tec2': "emian Sanchez",
    'input_h_marcha': 0, 'input_h_carga': 0, 'input_temp': "70.0",
    'input_p_carga': "7.0", 'input_p_descarga': "7.5", 'input_estado': "",
    'input_reco': "", 'input_estado_eq': "Operativo",
    'vista_firmas': False
}
for key, value in default_states.items():
    if key not in st.session_state: st.session_state[key] = value

if 'informes_pendientes' not in st.session_state:
    st.session_state.informes_pendientes = []

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
    st.session_state.equipo_seleccionado = None; st.session_state.vista_firmas = False; st.session_state.vista_actual = "catalogo"

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
                        st.session_state.informes_pendientes = cargar_pendientes(u_in)
                        st.rerun()
                    else: st.error("❌ Credenciales inválidas.")

# =============================================================================
# 6. PANTALLA PRINCIPAL: APLICACIÓN AUTENTICADA
# =============================================================================
else:
    with st.sidebar:
        st.markdown("<h2 style='text-align: center; border-bottom:none; margin-top: -20px;'><span style='color:#007CA6;'>Atlas Copco</span> <span style='color:#FF6600;'>Spence</span></h2>", unsafe_allow_html=True)
        st.markdown(f"**Usuario Activo:**<br>{st.session_state.usuario_actual.title()}", unsafe_allow_html=True)
        st.markdown("---")
        
        if st.button("🏭 Catálogo de Activos", use_container_width=True, type="primary" if st.session_state.vista_actual == "catalogo" else "secondary"):
            st.session_state.vista_actual = "catalogo"; st.session_state.vista_firmas = False; st.session_state.equipo_seleccionado = None; st.rerun()
            
        if st.button("📅 Planificación Hidrometalurgia", use_container_width=True, type="primary" if st.session_state.vista_actual == "planificacion" else "secondary"):
            st.session_state.vista_actual = "planificacion"; st.session_state.vista_firmas = False; st.session_state.equipo_seleccionado = None; st.rerun()
            
        if len(st.session_state.informes_pendientes) > 0:
            st.markdown("---")
            st.warning(f"📝 Tienes {len(st.session_state.informes_pendientes)} reportes esperando firmas.")
            if st.button("✍️ Ir a Pizarra de Firmas", use_container_width=True, type="primary" if st.session_state.vista_actual == "firmas" else "secondary"): 
                st.session_state.vista_firmas = True; st.session_state.vista_actual = "firmas"; st.session_state.equipo_seleccionado = None; st.rerun()
        st.markdown("---")
        if st.button("🚪 Cerrar Sesión", use_container_width=True): st.session_state.logged_in = False; st.rerun()

    # --- 6.0 VISTA MATRIZ Y GESTIÓN DUAL (ANTI-ERRORES) ---
    if st.session_state.vista_actual == "planificacion":
        df_plan = cargar_planificacion()
        if "Área" not in df_plan.columns or "TAG" not in df_plan.columns: df_plan = generar_planificacion_base()
        df_plan = df_plan.fillna("")
        
        mes_plan, rango_fechas = obtener_quincena_actual()
        mes_col_actual = f"15c {mes_plan[:3]}"
        
        st.markdown(f"""
            <div style="margin-top: 1rem; margin-bottom: 1rem; background: linear-gradient(90deg, rgba(0,124,166,0.1) 0%, rgba(0,124,166,0.2) 50%, rgba(0,124,166,0.1) 100%); padding: 20px; border-radius: 15px; border-left: 5px solid var(--ac-blue);">
                <h2 style="color: white; margin: 0;">📅 Planificación Operativa</h2>
                <p style="color: #8c9eb5; margin: 0; font-weight: 600;">Ciclo en curso: {rango_fechas}</p>
            </div>
        """, unsafe_allow_html=True)
        
        col_w1, col_w2 = st.columns([1, 4])
        with col_w1:
            semana_actual = st.text_input("📆 Semana en curso:", value="WK10", help="Esta semana se registrará automáticamente en la matriz.")
        
        st.markdown("<br>", unsafe_allow_html=True)
        tab_faltantes, tab_kanban, tab_matriz = st.tabs(["⚠️ Listado de Faltantes (Tickets)", "🗓️ Tablero Turno (4x3)", "📊 Matriz Anual Completa"])

        # ==========================================
        # PESTAÑA 1: FALTANTES DE LA QUINCENA (TICKETS INTERACTIVOS + COLUMNA DE INTERVENCIÓN)
        # ==========================================
        with tab_faltantes:
            st.markdown("### ⚠️ Equipos Faltantes de la Quincena")
            st.info("Marca con un ticket (✔️) la casilla de la izquierda para los equipos que ya realizaste y dale a Guardar. Se anotarán como 'Hecho' con la semana actual.")
            
            if mes_col_actual in df_plan.columns:
                df_quincena_act = df_plan[df_plan[mes_col_actual].str.strip() != ""]
                df_faltantes = df_quincena_act[~df_quincena_act[mes_col_actual].str.upper().str.contains('HECHO|OK|LISTO')].copy()
                
                if not df_faltantes.empty:
                    # EXTRAER LA PAUTA PARA LA NUEVA COLUMNA
                    import re
                    def extraer_pauta(txt):
                        match = re.search(r'(P[1-4]|INSP|I)', str(txt).upper())
                        return match.group(1) if match else "INSP"
                        
                    df_faltantes["Intervención"] = df_faltantes[mes_col_actual].apply(extraer_pauta)
                    df_faltantes.insert(0, "✔️ Terminado", False)
                    
                    df_mostrar_falta = df_faltantes[['✔️ Terminado', 'TAG', 'Equipo', 'Área', 'Intervención', mes_col_actual]]
                    
                    # APLICAR LOS COLORES ASIGNADOS SOLO A LA COLUMNA INTERVENCIÓN
                    try: df_falta_estilo = df_mostrar_falta.style.map(estilo_pautas_puras, subset=['Intervención']).map(estilo_simple_editor, subset=[mes_col_actual])
                    except AttributeError: df_falta_estilo = df_mostrar_falta.style.applymap(estilo_pautas_puras, subset=['Intervención']).applymap(estilo_simple_editor, subset=[mes_col_actual])
                    
                    configuracion_columnas = {
                        "✔️ Terminado": st.column_config.CheckboxColumn("¿Listo?", default=False),
                        "TAG": st.column_config.TextColumn("TAG", disabled=True),
                        "Equipo": st.column_config.TextColumn("Equipo", disabled=True),
                        "Área": st.column_config.TextColumn("Área", disabled=True),
                        "Intervención": st.column_config.TextColumn("Intervención", disabled=True),
                        mes_col_actual: st.column_config.TextColumn("Comentario Original", disabled=True)
                    }
                    
                    edited_faltantes = st.data_editor(
                        df_falta_estilo,
                        hide_index=True,
                        use_container_width=True,
                        column_config=configuracion_columnas,
                        height=500
                    )
                    
                    if st.button("💾 Guardar Equipos Terminados", type="primary"):
                        terminados = edited_faltantes[edited_faltantes["✔️ Terminado"] == True]
                        if len(terminados) > 0:
                            for _, row in terminados.iterrows():
                                tag_completado = row["TAG"]
                                pauta_limpia = row["Intervención"]
                                idx = df_plan.index[df_plan['TAG'] == tag_completado].tolist()[0]
                                
                                # Escribimos Hecho + WK automáticamente usando la pauta limpia
                                df_plan.at[idx, mes_col_actual] = f"{pauta_limpia}\nHecho {semana_actual}"

                            guardar_planificacion(df_plan)
                            st.success(f"✅ ¡Excelente! {len(terminados)} equipos actualizados a 'Hecho {semana_actual}'.")
                            st.rerun()
                        else:
                            st.warning("No marcaste ningún equipo con el ticket. Haz clic en el cuadradito vacío primero.")
                else:
                    st.success("🎉 ¡Impresionante! No hay ningún equipo pendiente para esta quincena.")

        # ==========================================
        # PESTAÑA 2: TABLERO 4x3 CON SCROLL INTERNO
        # ==========================================
        with tab_kanban:
            st.markdown("""
                <style>
                .kanban-col { background-color: #1a212b; border: 1px solid #2b3543; border-radius: 8px; padding: 15px; height: 500px; overflow-y: auto; position: relative; }
                .kanban-col::-webkit-scrollbar { width: 6px; }
                .kanban-col::-webkit-scrollbar-track { background: transparent; }
                .kanban-col::-webkit-scrollbar-thumb { background-color: #455065; border-radius: 10px; }
                .kanban-col::-webkit-scrollbar-thumb:hover { background-color: #00BFFF; }
                .kanban-header { color: white; text-align: center; border-bottom: 3px solid; padding-bottom: 10px; margin-bottom: 15px; font-weight: bold; position: sticky; top: -15px; background-color: #1a212b; z-index: 10; padding-top: 5px; }
                .kanban-card { background-color: #2b3543; border-left: 4px solid #007CA6; border-radius: 6px; padding: 12px; margin-bottom: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
                .kanban-card-title { color: white; font-weight: 800; font-size: 1.1rem; margin:0 0 5px 0; display: flex; justify-content: space-between; align-items: center;}
                .kanban-card-sub { color: #8c9eb5; font-size: 0.8rem; margin:0; }
                </style>
            """, unsafe_allow_html=True)

            if mes_col_actual not in df_plan.columns:
                st.error(f"No hay una columna llamada {mes_col_actual} en la matriz.")
            else:
                df_quincena = df_plan[df_plan[mes_col_actual].str.strip() != ""]
                
                lunes, martes, miercoles, jueves, completados = [], [], [], [], []
                pendientes_lista = []
                
                for _, row in df_quincena.iterrows():
                    texto = str(row[mes_col_actual]).upper()
                    import re
                    match = re.search(r'(P[1-4]|INSP|I)', texto)
                    pauta_txt = match.group(1) if match else "INSP"
                    
                    item = {"tag": row["TAG"], "eq": row["Equipo"], "area": row["Área"], "txt": pauta_txt}
                    
                    if any(x in texto for x in ['HECHO', 'OK', 'LISTO']): completados.append(item)
                    elif "LUNES" in texto: lunes.append(item)
                    elif "MARTES" in texto: martes.append(item)
                    elif "MIÉRCOLES" in texto or "MIERCOLES" in texto: miercoles.append(item)
                    elif "JUEVES" in texto: jueves.append(item)
                    else: pendientes_lista.append(item)

                k_cols = st.columns(5)
                
                def render_kanban_col(col_obj, title, items, color_border):
                    with col_obj:
                        st.markdown(f'<div class="kanban-col"><div class="kanban-header" style="border-bottom-color: {color_border};">{title} ({len(items)})</div>', unsafe_allow_html=True)
                        for it in items:
                            st.markdown(f"""
                            <div class="kanban-card" style="border-left-color: {color_border};">
                                <div class="kanban-card-title"><span>{it['tag']}</span> <span style="font-size:0.8rem; background:#1e2530; color:{color_border}; padding:3px 8px; border-radius:4px; border: 1px solid {color_border};">{it['txt']}</span></div>
                                <p class="kanban-card-sub">{it['eq']} • {it['area']}</p>
                            </div>
                            """, unsafe_allow_html=True)
                        st.markdown('</div>', unsafe_allow_html=True)

                render_kanban_col(k_cols[0], "Día 1 (Lunes)", lunes, "#00BFFF")
                render_kanban_col(k_cols[1], "Día 2 (Martes)", martes, "#00BFFF")
                render_kanban_col(k_cols[2], "Día 3 (Miércoles)", miercoles, "#00BFFF")
                render_kanban_col(k_cols[3], "Día 4 (Jueves)", jueves, "#F44336")
                render_kanban_col(k_cols[4], "✅ Completados", completados, "#00e676")

                st.markdown("---")
                st.markdown("### ⚙️ Mover Tarjetas de Día")
                with st.form("form_asignacion_kanban"):
                    c_f1, c_f2, c_f3 = st.columns([2, 1, 1])
                    todos_disponibles = [f"{it['tag']} ({it['txt']}) - {it['area']}" for it in pendientes_lista + lunes + martes + miercoles + jueves]
                    if len(todos_disponibles) == 0: todos_disponibles = ["No hay tareas para asignar"]
                    
                    tag_sel_raw = c_f1.selectbox("1. Elige el Equipo (Faltantes o Asignados):", ["-- Selecciona un equipo --"] + todos_disponibles)
                    dia_asignar = c_f2.selectbox("2. Mover a:", ["Lunes", "Martes", "Miércoles", "Jueves", "Devolver a Faltantes"])
                    
                    c_f3.markdown("<div style='margin-top:28px;'></div>", unsafe_allow_html=True)
                    if c_f3.form_submit_button("🚀 Actualizar Tablero", use_container_width=True):
                        if tag_sel_raw != "-- Selecciona un equipo --" and tag_sel_raw != "No hay tareas para asignar":
                            tag_asignar = tag_sel_raw.split(" ")[0] 
                            idx = df_plan.index[df_plan['TAG'] == tag_asignar].tolist()[0]
                            celda_actual = str(df_plan.at[idx, mes_col_actual])
                            
                            import re
                            match_p = re.search(r'(P[1-4]|INSP|I)', celda_actual.upper())
                            pauta_limpia = match_p.group(1) if match_p else "INSP"
                            
                            if dia_asignar == "Devolver a Faltantes": nuevo_texto = f"{pauta_limpia}\nFalta"
                            else: nuevo_texto = f"{pauta_limpia}\n{dia_asignar} {semana_actual}"
                                
                            df_plan.at[idx, mes_col_actual] = nuevo_texto
                            guardar_planificacion(df_plan)
                            st.success(f"✅ {tag_asignar} movido a {dia_asignar} {semana_actual}.")
                            st.rerun()

        # ==========================================
        # PESTAÑA 3: LA MATRIZ ANUAL (MACROMANEJO)
        # ==========================================
        with tab_matriz:
            col_fil1, col_fil2, col_fil3 = st.columns([1, 1, 1.5])
            with col_fil1:
                areas_disp = ["Todas"] + sorted(list(df_plan["Área"].unique()))
                filtro_area = st.selectbox("🏢 Filtrar por Área:", areas_disp, key="filtro_area_matriz")
            with col_fil2:
                modo_edicion_matriz = st.toggle("✏️ Edición de Matriz Completa")
            with col_fil3:
                st.markdown("<div style='margin-top:30px;'></div>", unsafe_allow_html=True)
                if modo_edicion_matriz: st.info("Edita cualquier celda del año completo.")
                
            df_mostrar = df_plan.copy()
            if filtro_area != "Todas": df_mostrar = df_mostrar[df_mostrar["Área"] == filtro_area]
                
            columnas_15cenas = [col for col in df_plan.columns if "15c" in col]

            if modo_edicion_matriz:
                try: df_estilizado_edit = df_mostrar.style.map(estilo_simple_editor, subset=columnas_15cenas)
                except AttributeError: df_estilizado_edit = df_mostrar.style.applymap(estilo_simple_editor, subset=columnas_15cenas)
                config_cols = {col: st.column_config.TextColumn(width="medium") for col in columnas_15cenas}
                df_editado = st.data_editor(df_estilizado_edit, use_container_width=True, hide_index=True, height=700, column_config=config_cols)
                if st.button("💾 Guardar Matriz en Nube", type="primary", use_container_width=True):
                    df_final_guardar = df_plan.copy()
                    df_editado_str = df_editado.astype(str)
                    df_final_guardar.update(df_editado_str)
                    guardar_planificacion(df_final_guardar)
                    st.success("✅ ¡Base de Datos actualizada con éxito!")
                    st.rerun()
            else:
                try: df_estilizado_view = df_mostrar.style.map(estilo_dinamico_celdas, subset=columnas_15cenas)
                except AttributeError: df_estilizado_view = df_mostrar.style.applymap(estilo_dinamico_celdas, subset=columnas_15cenas)
                st.dataframe(df_estilizado_view, use_container_width=True, hide_index=True, height=700)

    # --- 6.1 VISTA DE FIRMAS (FIRMA MANUAL LIMPIA) ---
    elif st.session_state.vista_firmas or st.session_state.vista_actual == "firmas":
        c_v1, c_v2 = st.columns([1,4])
        with c_v1: 
            if st.button("⬅️ Volver", use_container_width=True): volver_catalogo(); st.rerun()
        with c_v2: st.markdown("<h1 style='margin-top:-15px;'>✍️ Pizarra de Firmas Digital</h1>", unsafe_allow_html=True)
        st.markdown("---"); st.markdown(f"### 📑 Revisión de Informes ({len(st.session_state.informes_pendientes)})")
        
        for i, inf in enumerate(st.session_state.informes_pendientes):
            c_exp, c_del = st.columns([12, 1])
            with c_exp:
                with st.expander(f"📄 Ver documento preliminar: {inf['tag']} ({inf['tipo_plan']})"):
                    if inf.get('ruta_prev_pdf') and os.path.exists(inf['ruta_prev_pdf']):
                        try: pdf_viewer(inf['ruta_prev_pdf'], width=700, height=600)
                        except Exception as e: st.error(f"No se pudo desplegar el visor: {e}")
                        st.markdown("<br>", unsafe_allow_html=True)
                    else: st.warning("⚠️ La vista preliminar no está disponible.")
            with c_del:
                st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)
                if st.button("❌", key=f"del_inf_{i}", help="Quitar este informe de la bandeja"):
                    st.session_state.informes_pendientes.pop(i)
                    guardar_pendientes(st.session_state.usuario_actual, st.session_state.informes_pendientes) 
                    if len(st.session_state.informes_pendientes) == 0: volver_catalogo()
                    st.rerun()
                    
        st.markdown("---")
        
        c_tec, c_cli = st.columns(2)
        with c_tec:
            st.markdown("### 🧑‍🔧 Firma del Técnico")
            canvas_tec = st_canvas(stroke_width=4, stroke_color="#000", background_color="#fff", height=200, width=400, drawing_mode="freedraw", key="canvas_tecnico")
        with c_cli:
            st.markdown("### 👷 Firma del Cliente")
            canvas_cli = st_canvas(stroke_width=4, stroke_color="#000", background_color="#fff", height=200, width=400, drawing_mode="freedraw", key="canvas_cliente")
        
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🚀 Aprobar, Firmar y Subir a la Nube", type="primary", use_container_width=True):
            
            tec_ok = canvas_tec.image_data is not None and canvas_tec.json_data is not None and len(canvas_tec.json_data.get("objects", [])) > 0
            cli_ok = canvas_cli.image_data is not None and canvas_cli.json_data is not None and len(canvas_cli.json_data.get("objects", [])) > 0
            
            if tec_ok and cli_ok:
                def procesar_imagen_firma(img_data):
                    img = Image.fromarray(img_data.astype('uint8'), 'RGBA'); img_io = io.BytesIO(); img.save(img_io, format='PNG'); img_io.seek(0); return img_io
                
                io_tec = procesar_imagen_firma(canvas_tec.image_data)
                io_cli = procesar_imagen_firma(canvas_cli.image_data)
                
                informes_finales = []
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
                        if exito: 
                            st.success("✅ ¡PERFECTO! Los documentos oficiales se firmaron, convirtieron a PDF y ya están camino a tu OneDrive.")
                            st.session_state.informes_pendientes = []
                            guardar_pendientes(st.session_state.usuario_actual, []) 
                            st.balloons()
                        else: st.error(f"Error de red: {mensaje_correo}")
                    except Exception as e: st.error(f"Error sistémico procesando las firmas: {e}")
            else: 
                st.warning("⚠️ Asegúrate de que ambas pizarras contengan una firma visible antes de generar los PDFs finales.")

    # --- 6.2 VISTA CATÁLOGO (DASHBOARD CINETICO Y PREMIUM) ---
    elif st.session_state.vista_actual == "catalogo" and st.session_state.equipo_seleccionado is None:
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
                    color_borde = "#00e676"; badge_html = "<div style='background: rgba(0,230,118,0.15); color: #00e676; border: 1px solid #00e676; padding: 4px 10px; border-radius: 20px; font-size: 0.75rem; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; display: inline-block;'>OPERATIVO</div>"
                else:
                    color_borde = "#ff1744"; badge_html = "<div style='background: rgba(255,23,68,0.15); color: #ff1744; border: 1px solid #ff1744; padding: 4px 10px; border-radius: 20px; font-size: 0.75rem; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; display: inline-block;'>FUERA DE SERVICIO</div>"
                
                with columnas[contador % 4]:
                    with st.container(border=True):
                        st.markdown(f"<div style='border-top: 4px solid {color_borde}; padding-top: 10px; text-align: center; margin-top:-10px;'>{badge_html}</div>", unsafe_allow_html=True)
                        st.button(f"{tag}", key=f"btn_{tag}", on_click=seleccionar_equipo, args=(tag,), use_container_width=True)
                        st.markdown(f"<p style='color: #8c9eb5; margin-top: 5px; font-size: 0.85rem; text-align: center;'><strong style='color:#007CA6;'>{modelo}</strong> &bull; {area.title()}</p>", unsafe_allow_html=True)
                contador += 1

    # --- 6.3 VISTA FORMULARIO Y GENERACIÓN ---
    elif st.session_state.equipo_seleccionado is not None:
        tag_sel = st.session_state.equipo_seleccionado; mod_d, ser_d, area_d, ubi_d = inventario_equipos[tag_sel]
        c_btn, c_tit = st.columns([1, 4])
        with c_btn: st.button("⬅️ Volver", on_click=volver_catalogo, use_container_width=True)
        with c_tit: st.markdown(f"<h1 style='margin-top:-15px;'>⚙️ Ficha de Serviço: <span style='color:#007CA6;'>{tag_sel}</span></h1>", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True); tab1, tab2, tab3, tab4 = st.tabs(["📋 1. Reporte y Diagnóstico", "📚 2. Ficha Técnica", "🔍 3. Bitácora de Observaciones", "👤 4. Gestión de Área"])
        with tab1:
            st.markdown("### Datos de la Intervención"); tipo_plan = st.selectbox("🛠️ Tipo de Plan / Orden:", ["Inspección", "PM03"] if "CD" in tag_sel else ["Inspección", "P1", "P2", "P3", "PM03"]); c1, c2, c3, c4 = st.columns(4); modelo = c1.text_input("Modelo", mod_d, disabled=True); numero_serie = c2.text_input("N° Serie", ser_d, disabled=True); area = c3.text_input("Área", area_d, disabled=True); ubicacion = c4.text_input("Ubicación", ubi_d, disabled=True); c5, c6, c7, c8 = st.columns([1, 1, 1, 1.3])
            
            fecha = c5.text_input("Fecha Ejecución", obtener_fecha_hoy_esp())
            
            tec1 = c6.text_input("Técnico 1", key="input_tec1"); tec2 = c7.text_input("Técnico 2", key="input_tec2")
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
                st.session_state.informes_pendientes.append({"tag": tag_sel, "area": area_d, "tec1": tec1, "cli": cli_cont, "tipo_plan": tipo_plan, "file_plantilla": file_plantilla, "context": context, "tupla_db": tupla_db, "ruta_docx": ruta, "nombre_archivo_base": nombre_archivo, "ruta_prev_pdf": ruta_prev_pdf})
                guardar_pendientes(st.session_state.usuario_actual, st.session_state.informes_pendientes) 
                st.success("✅ Datos guardados. Agrega otro equipo o ve a la bandeja para firmar."); st.session_state.equipo_seleccionado = None; st.rerun()
                    
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