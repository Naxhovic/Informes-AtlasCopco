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
# 0.2 ESTILOS PREMIUM (EXTREMO ANTI-STREAMLIT Y GITHUB)
# =============================================================================
st.set_page_config(page_title="Atlas Spence | Gestión de Reportes", layout="wide", page_icon="⚙️", initial_sidebar_state="expanded")

def aplicar_estilos_premium():
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;600;800&display=swap');
        :root { --ac-blue: #007CA6; --ac-dark: #005675; --bhp-orange: #FF6600; }
        html, body, p, h1, h2, h3, h4, h5, h6, span, div { font-family: 'Montserrat', sans-serif; }
        
        /* 🔥 ANIQUILACIÓN EXTREMA: MENÚS, STREAMLIT Y GITHUB 🔥 */
        header {visibility: hidden !important; display: none !important;}
        [data-testid="stHeader"] {display: none !important;}
        [data-testid="stToolbar"] {display: none !important;}
        [data-testid="stDecoration"] {display: none !important;}
        #MainMenu {display: none !important;} 
        footer {display: none !important;} 
        .stAppDeployButton {display: none !important;}
        
        /* 🚫 BLOQUEO DIRECTO A CUALQUIER LINK DE GITHUB Y MARCAS DE AGUA 🚫 */
        a[href*="github.com"] { display: none !important; visibility: hidden !important; opacity: 0 !important; pointer-events: none !important; }
        [data-testid="viewerBadge"] {display: none !important;}
        div[class^="viewerBadge_container"] {display: none !important;}
        
        /* BOTÓN DE MENÚ LATERAL CELESTE VIBRANTE */
        [data-testid="collapsedControl"] {
            display: flex !important; visibility: visible !important; opacity: 1 !important;
            z-index: 999999 !important; background-color: #00BFFF !important; 
            border-radius: 8px !important; box-shadow: 0 4px 15px rgba(0, 191, 255, 0.4) !important;
            margin-top: 15px !important; margin-left: 15px !important; transition: all 0.3s ease !important;
        }
        [data-testid="collapsedControl"]:hover { background-color: var(--ac-blue) !important; transform: scale(1.05) !important; }
        [data-testid="collapsedControl"] svg { fill: white !important; stroke: white !important; }
        
        div.stButton > button:first-child { background: linear-gradient(135deg, var(--ac-blue) 0%, var(--ac-dark) 100%); color: white; border-radius: 8px; border: none; font-weight: 600; padding: 0.6rem 1.2rem; transition: all 0.3s ease; box-shadow: 0 4px 15px rgba(0, 124, 166, 0.4); }
        div.stButton > button:first-child:hover { transform: translateY(-2px); box-shadow: 0 8px 20px rgba(0, 124, 166, 0.6); }
        
        [data-testid="stVerticalBlockBorderWrapper"] { background: linear-gradient(145deg, #1a212b, #151a22) !important; border-radius: 12px !important; border: 1px solid #2b3543 !important; transition: transform 0.3s ease, box-shadow 0.3s ease, border-color 0.3s ease !important; }
        [data-testid="stVerticalBlockBorderWrapper"]:hover { transform: translateY(-6px) !important; box-shadow: 0 10px 25px rgba(0, 124, 166, 0.25) !important; border-color: var(--ac-blue) !important; }
        
        .stTextInput>div>div>input, .stNumberInput>div>div>input, .stSelectbox>div>div>select { border-radius: 6px !important; border: 1px solid #2b3543 !important; background-color: #1e2530 !important; color: white !important; }
        .stTextInput>div>div>input:focus, .stNumberInput>div>div>input:focus, .stSelectbox>div>div>select:focus { border-color: var(--bhp-orange) !important; box-shadow: 0 0 8px rgba(255, 102, 0, 0.3) !important; }
        .stTabs [data-baseweb="tab-list"] { border-bottom: 2px solid #2b3543; }
        .stTabs [aria-selected="true"] { color: var(--bhp-orange) !important; border-bottom: 3px solid var(--bhp-orange) !important; }
        </style>
    """, unsafe_allow_html=True)
aplicar_estilos_premium()

# =============================================================================
# 1. DATOS MAESTROS E INVENTARIO
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
    "20-GC-001": ["GA 75", "AII482673", "truck shop", "Mina"], "20-GC-002": ["GA 75", "AII482674", "truck shop", "Mina"], "20-GC-003": ["GA 90", "AIF095178", "truck shop", "Mina"], "20-GC-004": ["GA 37", "AII390776", "truck shop", "Mina"],
    "35-GC-006": ["GA 250", "AIF095420", "chancado secundario", "Área Seca"], "35-GC-007": ["GA 250", "AIF095421", "chancado secundario", "Área Seca"], "35-GC-008": ["GA 250", "AIF095302", "chancado secundario", "Área Seca"],
    "50-GC-001": ["GA 45", "API542705", "planta SX", "Área Húmeda"], "50-GC-002": ["GA 45", "API542706", "planta SX", "Área Húmeda"], "50-GC-003": ["ZT 37", "API791692", "planta SX", "Área Húmeda"], "50-GC-004": ["ZT 37", "API791693", "planta SX", "Área Húmeda"], "50-CD-001": ["CD 80+", "API095825", "planta SX", "Área Húmeda"], "50-CD-002": ["CD 80+", "API095826", "planta SX", "Área Húmeda"],
    "55-GC-015": ["GA 30", "API501440", "planta borra", "Área Húmeda"],
    "65-GC-009": ["GA 250", "APF253608", "patio de estanques", "Área Húmeda"], "65-GC-011": ["GA 250", "APF253581", "patio de estanques", "Área Húmeda"], "65-CD-011": ["CD 630", "WXF300015", "patio de estanques", "Área Húmeda"], "65-CD-012": ["CD 630", "WXF300016", "patio de estanques", "Área Húmeda"],
    "70-GC-013": ["GA 132", "AIF095296", "descarga de acido", "Área Húmeda"], "70-GC-014": ["GA 132", "AIF095297", "descarga de acido", "Área Húmeda"],
    "Taller": ["GA 18", "API335343", "Taller", "Taller Central"]
}

# =============================================================================
# 2. CONEXIÓN A GOOGLE SHEETS (COMPATIBLE CON RENDER Y STREAMLIT CLOUD)
# =============================================================================
@st.cache_resource
def get_gspread_client():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    try:
        # Intenta usar la variable de entorno de Render
        creds_dict = json.loads(os.environ["gcp_json"])
    except:
        # Si falla, usa los secretos de Streamlit (por si acaso)
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

# =============================================================================
# 3. FUNCIONES DE BASE DE DATOS (ESTADOS PERSISTENTES)
# =============================================================================
@st.cache_data(ttl=120)
def obtener_estados_actuales():
    estados = {}
    try:
        sheet = get_sheet("estados_equipos") # Lee directamente la pestaña dedicada a los estados
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
            try:
                # Si el equipo ya existe, lo actualizamos
                celda = sheet.find(tag)
                sheet.update_cell(celda.row, 2, nuevo_estado)
            except Exception:
                # Si no existe en la hoja, lo agregamos al final
                sheet.append_row([tag, nuevo_estado])
            st.cache_data.clear() # Limpia la caché para que el catálogo se actualice al instante
    except Exception as e: pass

@st.cache_data(ttl=120)
def obtener_datos_equipo(tag):
    datos = {}; sheet = get_sheet("datos_equipo")
    if sheet:
        data = sheet.get_all_values()
        for r in data:
            if len(r) >= 3 and r[0] == tag: datos[r[1]] = r[2]
    return datos

@st.cache_data(ttl=120)
def obtener_observaciones(tag):
    sheet = get_sheet("observaciones")
    if sheet:
        data = sheet.get_all_values()
        obs = [{"id": r[0], "fecha": r[2], "usuario": r[3], "texto": r[4]} for r in data if len(r) >= 6 and r[1] == tag and r[5] == "ACTIVO"]
        if obs: return pd.DataFrame(obs).iloc[::-1]
    return pd.DataFrame(columns=["id", "fecha", "usuario", "texto"])

@st.cache_data(ttl=120)
def obtener_contactos():
    sheet = get_sheet("contactos")
    if sheet:
        data = sheet.get_all_values()
        contactos = [r[0] for r in data if len(r) > 1 and r[1] == "ACTIVO"]
        if contactos: return sorted(list(set(contactos)))
    return ["Lorena Rojas"]

@st.cache_data(ttl=120)
def buscar_ultimo_registro(tag):
    sheet = get_sheet("intervenciones")
    if sheet:
        data = sheet.get_all_values()
        for row in reversed(data):
            if len(row) >= 20 and row[0] == tag: return (row[5], row[6], row[9], row[14], row[15], row[7], row[8], row[10], row[11], row[12], row[13], row[16], row[17])
    return None

def guardar_dato_equipo(tag, clave, valor):
    sheet = get_sheet("datos_equipo")
    if sheet: sheet.append_row([tag, clave, valor]); st.cache_data.clear()

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

def guardar_especificacion_db(modelo, clave, valor):
    sheet = get_sheet("especificaciones")
    if sheet: sheet.append_row([modelo, clave, valor]); st.cache_data.clear()
    # =============================================================================
# 4. FUNCIONES AUXILIARES Y PLANIFICACIÓN
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
    return None

def obtener_fecha_hoy_esp():
    meses = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
    ahora = pd.Timestamp.now()
    return f"{ahora.day} de {meses[ahora.month]} de {ahora.year}"

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

def obtener_quincena_actual():
    hoy = datetime.date.today(); meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    if hoy.day < 15: return meses[hoy.month - 1], f"15 de {meses[hoy.month - 2 if hoy.month > 1 else 11]} al 15 de {meses[hoy.month - 1]}"
    else: return meses[hoy.month] if hoy.month < 12 else "Enero", f"15 de {meses[hoy.month - 1]} al 15 de {meses[hoy.month] if hoy.month < 12 else 'Enero'}"

def obtener_mes_actual_abrev(): return ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"][datetime.date.today().month - 1]

def generar_planificacion_base():
    datos = [
        {"TAG": "70-GC-013", "Equipo": "GA 132", "Área": "Descarga Acido", "15c Ene": "INSP", "15c Feb": "P1\n20/02 (WK8)", "15c Mar": "INSP", "15c Abr": "P4", "15c May": "INSP", "15c Jun": "P1", "15c Jul": "INSP", "15c Ago": "P2", "15c Sep": "INSP", "15c Oct": "P1", "15c Nov": "INSP", "15c Dic": "P3"},
        {"TAG": "35-GC-006", "Equipo": "GA 250", "Área": "Chancado Sec.", "15c Ene": "P1\nFalta", "15c Feb": "P1\nF/S", "15c Mar": "P2\nF/S", "15c Abr": "P1", "15c May": "P1", "15c Jun": "P2", "15c Jul": "P1", "15c Ago": "P1", "15c Sep": "P4", "15c Oct": "P1", "15c Nov": "P1", "15c Dic": "P2"},
        {"TAG": "20-GC-001", "Equipo": "GA 75", "Área": "Truck Shop", "15c Ene": "INSP", "15c Feb": "P1", "15c Mar": "INSP", "15c Abr": "P4", "15c May": "INSP", "15c Jun": "P1", "15c Jul": "INSP", "15c Ago": "P2", "15c Sep": "INSP", "15c Oct": "P1", "15c Nov": "INSP", "15c Dic": "P3"}
    ]
    return pd.DataFrame(datos)

@st.cache_data(ttl=60)
def cargar_planificacion():
    sheet = get_sheet("planificacion")
    if sheet:
        data = sheet.get_all_values()
        if len(data) > 1:
            df = pd.DataFrame(data[1:], columns=data[0])
            if "15c Ene" in df.columns: return df
    return generar_planificacion_base()

def guardar_planificacion(df):
    sheet = get_sheet("planificacion")
    if sheet:
        sheet.clear() 
        sheet.append_rows([df.columns.values.tolist()] + df.values.tolist()); st.cache_data.clear() 

def estilo_dinamico_celdas(val):
    if pd.isna(val) or val == "": return ''
    v = str(val).upper()
    base = 'white-space: pre-wrap; line-height: 1.4; border-radius: 6px; padding: 6px; text-align: center; '
    if 'F/S' in v or 'FUERA' in v: return base + 'background-color: #471015; color: #ff8a93; font-weight: bold; border-left: 4px solid #ef4444;'
    if any(x in v for x in ['FALTA', 'PENDIENTE', 'WK', 'PEND']): return base + 'background-color: #423205; color: #fde047; font-weight: bold; border-left: 4px solid #eab308;'
    import re
    if re.search(r'(\d{2}/\d{2}|WK\d+|HECHO|OK)', v): return base + 'background-color: #063f22; color: #6ee7b7; font-weight: bold; border-left: 4px solid #10b981;'
    if 'P1' in v: return base + 'background-color: #0c2d48; color: #66c2ff; font-weight: bold;'
    if 'P2' in v: return base + 'background-color: #4a2c00; color: #ffb04c; font-weight: bold;'
    if 'P3' in v: return base + 'background-color: #301047; color: #d78aff; font-weight: bold;'
    if 'P4' in v: return base + 'background-color: #471015; color: #ff8a93; font-weight: bold;'
    return base + 'color: #8c9eb5; font-style: italic;'

def estilo_simple_editor(val):
    v = str(val).upper()
    if 'F/S' in v: return 'background-color: #471015; color: #ff8a93;'
    import re
    if re.search(r'(\d{2}/\d{2}|WK\d+)', v) and not 'FALTA' in v: return 'background-color: #063f22; color: #6ee7b7;'
    if 'FALTA' in v: return 'background-color: #423205; color: #fde047;'
    if 'P1' in v: return 'background-color: #0c2d48; color: #66c2ff;' 
    if 'P4' in v: return 'background-color: #471015; color: #ff8a93;'
    return 'color: #8c9eb5;'

# =============================================================================
# 5. INICIALIZACIÓN DE ESTADOS
# =============================================================================
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

default_states = {
    'logged_in': False, 'usuario_actual': "", 'equipo_seleccionado': None, 'vista_actual': "catalogo",
    'input_cliente': "Lorena Rojas", 'input_tec1': "Ignacio Morales", 'input_tec2': "emian Sanchez",
    'input_h_marcha': 0, 'input_h_carga': 0, 'input_temp': "70.0",
    'input_p_carga': "7.0", 'input_p_descarga': "7.5", 'input_estado': "",
    'input_reco': "", 'input_estado_eq': "Operativo", 'vista_firmas': False
}
for key, value in default_states.items():
    if key not in st.session_state: st.session_state[key] = value

if 'informes_pendientes' not in st.session_state: st.session_state.informes_pendientes = []

# =============================================================================
# 6. INTERFAZ: LOGIN
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
# 7. INTERFAZ PRINCIPAL
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

    # --- 7.1 VISTA PLANIFICACIÓN ---
    if st.session_state.vista_actual == "planificacion":
        df_plan = cargar_planificacion(); df_plan = df_plan.fillna("")
        mes_plan, rango_fechas = obtener_quincena_actual(); mes_col_actual = f"15c {mes_plan[:3]}"
        
        st.markdown(f"""
            <div style="margin-top: 1rem; margin-bottom: 1rem; background: linear-gradient(90deg, rgba(0,124,166,0.1) 0%, rgba(0,124,166,0.2) 50%, rgba(0,124,166,0.1) 100%); padding: 20px; border-radius: 15px; border-left: 5px solid var(--ac-blue);">
                <h2 style="color: white; margin: 0;">📅 Planificación Operativa</h2>
                <p style="color: #8c9eb5; margin: 0; font-weight: 600;">Ciclo en curso: {rango_fechas}</p>
            </div>
        """, unsafe_allow_html=True)
        tab_faltantes, tab_calendario, tab_matriz = st.tabs(["⚠️ Faltantes (Tickets)", "📆 Mapa Histórico", "📊 Matriz Anual"])

        with tab_faltantes:
            st.markdown("### ⚠️ Equipos Pendientes")
            c_fec1, c_fec2 = st.columns([1, 4])
            with c_fec1: fecha_rapida = st.date_input("Fecha de ejecución a registrar:", datetime.date.today(), key="fecha_faltantes")
            if mes_col_actual in df_plan.columns:
                df_quincena_act = df_plan[df_plan[mes_col_actual].str.strip() != ""]
                import re
                def es_pendiente(val):
                    v = str(val).upper()
                    if 'FALTA' in v or 'PEND' in v: return True
                    if not re.search(r'(\d{2}/\d{2}|WK\d+|HECHO|OK)', v): return True
                    return False
                
                df_faltantes = df_quincena_act[df_quincena_act[mes_col_actual].apply(es_pendiente)].copy()
                if not df_faltantes.empty:
                    def extraer_pauta(txt):
                        match = re.search(r'(P[1-4]|INSP|I)', str(txt).upper())
                        return match.group(1) if match else "INSP"
                        
                    df_faltantes["Intervención"] = df_faltantes[mes_col_actual].apply(extraer_pauta)
                    df_faltantes.insert(0, "✔️ Terminado", False)
                    df_mostrar_falta = df_faltantes[['✔️ Terminado', 'TAG', 'Equipo', 'Área', 'Intervención', mes_col_actual]]
                    try: df_falta_estilo = df_mostrar_falta.style.applymap(lambda x: 'color:#66c2ff', subset=['Intervención'])
                    except: df_falta_estilo = df_mostrar_falta
                    config_cols = {"✔️ Terminado": st.column_config.CheckboxColumn("¿Listo?", default=False), "TAG": st.column_config.TextColumn(disabled=True), "Equipo": st.column_config.TextColumn(disabled=True), "Área": st.column_config.TextColumn(disabled=True), "Intervención": st.column_config.TextColumn(disabled=True), mes_col_actual: st.column_config.TextColumn("Original", disabled=True)}
                    edited_faltantes = st.data_editor(df_falta_estilo, hide_index=True, use_container_width=True, column_config=config_cols, height=500)
                    if st.button("💾 Guardar Equipos Terminados", type="primary"):
                        terminados = edited_faltantes[edited_faltantes["✔️ Terminado"] == True]
                        if len(terminados) > 0:
                            str_fecha = f"{fecha_rapida.strftime('%d/%m')} (WK{fecha_rapida.isocalendar()[1]})"
                            for _, row in terminados.iterrows():
                                df_plan.at[df_plan.index[df_plan['TAG'] == row["TAG"]].tolist()[0], mes_col_actual] = f"{row['Intervención']}\n{str_fecha}"
                            guardar_planificacion(df_plan); st.success(f"✅ Guardado."); st.rerun()
                else: st.success("🎉 No hay ningún equipo pendiente.")

        with tab_calendario:
            st.markdown("### 📆 Mapa Histórico del Mes")
            import calendar; hoy = datetime.date.today(); cal = calendar.Calendar(calendar.MONDAY); semanas_mes = cal.monthdatescalendar(hoy.year, hoy.month)
            tareas_por_fecha = {}
            for col in df_plan.columns:
                if "15c" in col:
                    for idx, row in df_plan.iterrows():
                        val = str(row[col]).upper()
                        import re; matches = re.findall(r'(\d{2}/\d{2})', val)
                        for m in matches:
                            try:
                                d, m_num = map(int, m.split('/')); fecha_tarea = datetime.date(hoy.year, m_num, d)
                                if fecha_tarea not in tareas_por_fecha: tareas_por_fecha[fecha_tarea] = []
                                p_txt = re.search(r'(P[1-4]|INSP|I)', val).group(1) if re.search(r'(P[1-4]|INSP|I)', val) else "INSP"
                                tareas_por_fecha[fecha_tarea].append((row['TAG'], p_txt))
                            except: pass

            html_cal = '<div style="display:grid; grid-template-columns: repeat(7, 1fr); gap: 10px; margin-top:20px;">'
            for d in ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]: html_cal += f'<div style="text-align:center; color:#8c9eb5; font-weight:bold; font-size:0.9rem;">{d}</div>'
            for semana in semanas_mes:
                for dia in semana:
                    is_current_month = dia.month == hoy.month; bg_color = "#1a212b" if is_current_month else "#11151c"; border_color = "#00BFFF" if dia == hoy else "#2b3543"
                    html_cal += f'<div style="background:{bg_color}; border: 1px solid {border_color}; border-radius: 8px; padding: 5px; min-height: 120px;">'
                    html_cal += f'<div style="text-align:right; color:white; font-size:0.9rem; margin-bottom:8px;">{dia.day}</div>'
                    if dia in tareas_por_fecha:
                        for tag, pt in tareas_por_fecha[dia]: html_cal += f'<div style="background:#063f22; color:#6ee7b7; padding:4px; margin-bottom:4px; border-radius:4px; font-size:0.75rem;"><b>{tag}</b> {pt}</div>'
                    html_cal += '</div>'
            html_cal += '</div>'
            st.markdown(html_cal, unsafe_allow_html=True)

        with tab_matriz:
            col_fil1, col_fil2, col_fil3 = st.columns([1, 1, 1.5])
            with col_fil1: filtro_area = st.selectbox("🏢 Filtrar por Área:", ["Todas"] + sorted(list(df_plan["Área"].unique())))
            with col_fil2: modo_edicion_matriz = st.toggle("✏️ Edición de Matriz Completa")
            df_mostrar = df_plan.copy() if filtro_area == "Todas" else df_plan[df_plan["Área"] == filtro_area]
            columnas_15cenas = [col for col in df_plan.columns if "15c" in col]
            if modo_edicion_matriz:
                try: df_estilizado_edit = df_mostrar.style.map(estilo_simple_editor, subset=columnas_15cenas)
                except AttributeError: df_estilizado_edit = df_mostrar.style.applymap(estilo_simple_editor, subset=columnas_15cenas)
                df_editado = st.data_editor(df_estilizado_edit, use_container_width=True, hide_index=True, height=700)
                if st.button("💾 Guardar Matriz", type="primary", use_container_width=True):
                    df_final_guardar = df_plan.copy(); df_final_guardar.update(df_editado.astype(str)); guardar_planificacion(df_final_guardar); st.success("✅ Guardado!"); st.rerun()
            else:
                try: st.dataframe(df_mostrar.style.map(estilo_dinamico_celdas, subset=columnas_15cenas), use_container_width=True, hide_index=True, height=700)
                except AttributeError: st.dataframe(df_mostrar.style.applymap(estilo_dinamico_celdas, subset=columnas_15cenas), use_container_width=True, hide_index=True, height=700)

    # --- 7.2 VISTA DE FIRMAS (AGRUPADA POR MACRO-ÁREA EJ: MINA, ÁREA SECA) ---
    elif st.session_state.vista_firmas or st.session_state.vista_actual == "firmas":
        c_v1, c_v2 = st.columns([1,4])
        with c_v1: 
            if st.button("⬅️ Volver", use_container_width=True): volver_catalogo(); st.rerun()
        with c_v2: st.markdown("<h1 style='margin-top:-15px;'>✍️ Pizarra de Firmas por Área</h1>", unsafe_allow_html=True)
        st.markdown("---")
        
        if len(st.session_state.informes_pendientes) == 0:
            st.info("🎉 ¡Excelente! No tienes ningún informe pendiente por firmar.")
        else:
            # MAGIA DE AGRUPACIÓN POR LA MACRO ÁREA (Ubicación: "Mina", "Área Seca", etc.)
            areas_agrupadas = {}
            for inf in st.session_state.informes_pendientes:
                # Usamos la "ubicacion" que guardamos desde el diccionario de equipos
                macro_area = inf.get('ubicacion', 'General').title()
                if macro_area not in areas_agrupadas: areas_agrupadas[macro_area] = []
                areas_agrupadas[macro_area].append(inf)

            # Iteramos creando un bloque de firmas independiente para cada MACRO ÁREA
            for macro_area, informes_area in areas_agrupadas.items():
                st.markdown(f"### 🏢 Informes de {macro_area} ({len(informes_area)} pendientes)")
                
                with st.container(border=True):
                    # 1. Expanders para ver todos los reportes de esta área
                    for inf in informes_area:
                        c_exp, c_del = st.columns([12, 1])
                        with c_exp:
                            with st.expander(f"📄 Ver borrador: {inf['tag']} ({inf['tipo_plan']} - {inf['area'].title()})"):
                                if inf.get('ruta_prev_pdf') and os.path.exists(inf['ruta_prev_pdf']):
                                    try: pdf_viewer(inf['ruta_prev_pdf'], width=700, height=600)
                                    except Exception as e: st.error(f"Error visor: {e}")
                                else: st.warning("⚠️ Vista preliminar no disponible.")
                        with c_del:
                            st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)
                            if st.button("❌", key=f"del_{inf['tag']}_{inf['fecha'].replace(' ','_')}", help="Quitar este informe"):
                                st.session_state.informes_pendientes.remove(inf)
                                guardar_pendientes(st.session_state.usuario_actual, st.session_state.informes_pendientes) 
                                st.rerun()
                    
                    st.markdown("---")
                    
                    # 2. Lienzos de firma exclusivos para ESTA área (Usamos key dinámico para evitar conflictos)
                    c_tec, c_cli = st.columns(2)
                    with c_tec:
                        st.markdown(f"#### 🧑‍🔧 Firma Técnico")
                        canvas_tec = st_canvas(stroke_width=4, stroke_color="#000", background_color="#fff", height=180, width=400, drawing_mode="freedraw", key=f"tec_{macro_area}")
                    with c_cli:
                        st.markdown(f"#### 👷 Firma Cliente ({macro_area})")
                        canvas_cli = st_canvas(stroke_width=4, stroke_color="#000", background_color="#fff", height=180, width=400, drawing_mode="freedraw", key=f"cli_{macro_area}")
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    
                    # 3. Botón para enviar SÓLO los reportes de esta Área
                    if st.button(f"🚀 Aprobar, Firmar y Subir Informes de {macro_area}", type="primary", use_container_width=True, key=f"btn_subir_{macro_area}"):
                        tec_ok = canvas_tec.image_data is not None and canvas_tec.json_data is not None and len(canvas_tec.json_data.get("objects", [])) > 0
                        cli_ok = canvas_cli.image_data is not None and canvas_cli.json_data is not None and len(canvas_cli.json_data.get("objects", [])) > 0
                        
                        if tec_ok and cli_ok:
                            def procesar_imagen_firma(img_data):
                                img = Image.fromarray(img_data.astype('uint8'), 'RGBA'); img_io = io.BytesIO(); img.save(img_io, format='PNG'); img_io.seek(0); return img_io
                            
                            io_tec = procesar_imagen_firma(canvas_tec.image_data)
                            io_cli = procesar_imagen_firma(canvas_cli.image_data)
                            
                            informes_finales = []
                            with st.spinner(f"Inyectando firmas en los {len(informes_area)} reportes de {macro_area}..."):
                                try:
                                    for inf in informes_area:
                                        doc = DocxTemplate(inf['file_plantilla']); context = inf['context']
                                        context['firma_tecnico'] = InlineImage(doc, io_tec, width=Mm(40)); context['firma_cliente'] = InlineImage(doc, io_cli, width=Mm(40)); doc.render(context); doc.save(inf['ruta_docx']); ruta_pdf_gen = convertir_a_pdf(inf['ruta_docx'])
                                        if ruta_pdf_gen: ruta_final = ruta_pdf_gen; nombre_final = inf['nombre_archivo_base'].replace(".docx", ".pdf")
                                        else: ruta_final = inf['ruta_docx']; nombre_final = inf['nombre_archivo_base']
                                        tupla_lista = list(inf['tupla_db']); tupla_lista[18] = ruta_final; guardar_registro(tuple(tupla_lista))
                                        informes_finales.append({"tag": inf['tag'], "tipo": inf['tipo_plan'], "ruta": ruta_final, "nombre_archivo": f"{macro_area}@@{inf['tag']}@@{nombre_final}"})
                                    
                                    exito, mensaje_correo = enviar_carrito_por_correo(MI_CORREO_CORPORATIVO, informes_finales)
                                    if exito: 
                                        st.success(f"✅ ¡Listos y enviados los reportes de {macro_area}!")
                                        # Eliminamos de la bandeja global SÓLO los informes de esta macro área que ya firmamos
                                        st.session_state.informes_pendientes = [i for i in st.session_state.informes_pendientes if i.get('ubicacion', 'General').title() != macro_area]
                                        guardar_pendientes(st.session_state.usuario_actual, st.session_state.informes_pendientes) 
                                        time.sleep(2)
                                        st.rerun()
                                    else: st.error(f"Error de red: {mensaje_correo}")
                                except Exception as e: st.error(f"Error procesando los PDFs: {e}")
                        else: st.warning(f"⚠️ Asegúrate de firmar ambas pizarras para procesar los documentos de {macro_area}.")
                st.markdown("<br><br>", unsafe_allow_html=True)

    # --- 7.3 VISTA CATÁLOGO ---
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
            if busqueda in tag.lower() or busqueda in area.lower() or busqueda in modelo.lower() or busqueda in ubicacion.lower():
                estado = estados_db.get(tag, "Operativo")
                if estado == "Operativo":
                    color_borde = "#00e676"; badge_html = "<div style='background: rgba(0,230,118,0.15); color: #00e676; border: 1px solid #00e676; padding: 4px 10px; border-radius: 20px; font-size: 0.75rem; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; display: inline-block;'>OPERATIVO</div>"
                else:
                    color_borde = "#ff1744"; badge_html = "<div style='background: rgba(255,23,68,0.15); color: #ff1744; border: 1px solid #ff1744; padding: 4px 10px; border-radius: 20px; font-size: 0.75rem; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; display: inline-block;'>FUERA DE SERVICIO</div>"
                
                with columnas[contador % 4]:
                    with st.container(border=True):
                        st.markdown(f"<div style='border-top: 4px solid {color_borde}; padding-top: 10px; text-align: center; margin-top:-10px;'>{badge_html}</div>", unsafe_allow_html=True)
                        st.button(f"{tag}", key=f"btn_{tag}", on_click=seleccionar_equipo, args=(tag,), use_container_width=True)
                        st.markdown(f"<p style='color: #8c9eb5; margin-top: 5px; font-size: 0.85rem; text-align: center;'><strong style='color:#007CA6;'>{modelo}</strong> &bull; {area.title()}<br><small style='color: #556b82;'>{ubicacion.title()}</small></p>", unsafe_allow_html=True)
                contador += 1

    # --- 7.4 VISTA FORMULARIO Y GENERACIÓN ---
    elif st.session_state.equipo_seleccionado is not None:
        tag_sel = st.session_state.equipo_seleccionado; mod_d, ser_d, area_d, ubi_d = inventario_equipos[tag_sel]
        c_btn, c_tit = st.columns([1, 4])
        with c_btn: st.button("⬅️ Volver", on_click=volver_catalogo, use_container_width=True)
        with c_tit: st.markdown(f"<h1 style='margin-top:-15px;'>⚙️ Ficha de Serviço: <span style='color:#007CA6;'>{tag_sel}</span></h1>", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True); tab1, tab2, tab3, tab4 = st.tabs(["📋 1. Reporte y Diagnóstico", "📚 2. Ficha Técnica", "🔍 3. Bitácora de Observaciones", "👤 4. Gestión de Área"])
        with tab1:
            st.markdown("### Datos de la Intervención"); tipo_plan = st.selectbox("🛠️ Tipo de Plan / Orden:", ["Inspección", "PM03"] if "CD" in tag_sel else ["Inspección", "P1", "P2", "P3", "PM03"]); c1, c2, c3, c4 = st.columns(4); modelo = c1.text_input("Modelo", mod_d, disabled=True); numero_serie = c2.text_input("N° Serie", ser_d, disabled=True); area = c3.text_input("Área Específica", area_d, disabled=True); ubicacion = c4.text_input("Macro-Área", ubi_d, disabled=True); c5, c6, c7, c8 = st.columns([1, 1, 1, 1.3])
            
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
                
                # LA LÍNEA MÁGICA PARA QUE EL ESTADO NUNCA SE BORRE EN RENDER
                actualizar_estado_equipo_en_nube(tag_sel, est_eq)
                
                if "CD" in tag_sel: file_plantilla = "plantilla/secadorfueradeservicio.docx" if est_eq == "Fuera de servicio" else "plantilla/inspeccionsecador.docx"
                else: file_plantilla = "plantilla/fueradeservicio.docx" if est_eq == "Fuera de servicio" else f"plantilla/{tipo_plan.lower()}.docx" if tipo_plan in ["P1", "P2", "P3"] else "plantilla/inspeccion.docx"
                context = {"tipo_intervencion": tipo_plan, "modelo": mod_d, "tag": tag_sel, "area": area_d, "ubicacion": ubi_d, "cliente_contacto": cli_cont, "p_carga": f"{p_c_clean} {unidad_p}", "p_descarga": f"{p_d_clean} {unidad_p}", "temp_salida": t_salida_clean, "horas_marcha": int(h_m), "horas_carga": int(h_c), "tecnico_1": tec1, "tecnico_2": tec2, "estado_equipo": est_eq, "estado_entrega": est_ent, "recomendaciones": reco, "serie": ser_d, "tipo_orden": tipo_plan.upper(), "fecha": fecha, "equipo_modelo": mod_d}; nombre_archivo = f"Informe_{tipo_plan}_{tag_sel}_{fecha.replace(' ','_')}.docx"; ruta = os.path.join(RUTA_ONEDRIVE, nombre_archivo); temp_db = float(t_salida_clean) if t_salida_clean.replace('.', '', 1).isdigit() else 0.0; tupla_db = (tag_sel, mod_d, ser_d, area_d, ubi_d, fecha, cli_cont, tec1, tec2, temp_db, f"{p_c_clean} {unidad_p}", f"{p_d_clean} {unidad_p}", h_m, h_c, est_ent, tipo_plan, reco, est_eq, "", st.session_state.usuario_actual)
                with st.spinner("Creando borrador del documento para vista preliminar..."):
                    doc_prev = DocxTemplate(file_plantilla); ctx_prev = context.copy(); ctx_prev['firma_tecnico'] = ""; ctx_prev['firma_cliente'] = ""; doc_prev.render(ctx_prev); os.makedirs(RUTA_ONEDRIVE, exist_ok=True); ruta_prev_docx = os.path.join(RUTA_ONEDRIVE, f"PREVIEW_{nombre_archivo}"); doc_prev.save(ruta_prev_docx); ruta_prev_pdf = convertir_a_pdf(ruta_prev_docx)
                
                # AQUÍ AÑADIMOS LA MACRO ÁREA (ubicacion) PARA QUE LA BANDEJA PUEDA AGRUPARLOS
                st.session_state.informes_pendientes.append({"tag": tag_sel, "area": area_d, "ubicacion": ubi_d, "tec1": tec1, "cli": cli_cont, "tipo_plan": tipo_plan, "file_plantilla": file_plantilla, "context": context, "tupla_db": tupla_db, "ruta_docx": ruta, "nombre_archivo_base": nombre_archivo, "ruta_prev_pdf": ruta_prev_pdf})
                guardar_pendientes(st.session_state.usuario_actual, st.session_state.informes_pendientes) 
                
                st.success(f"✅ Datos guardados. El equipo se anotó como '{est_eq}' en tu Base de Datos y el informe se fue a la Bandeja de {ubi_d.title()}."); st.session_state.equipo_seleccionado = None; st.rerun()
                    
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