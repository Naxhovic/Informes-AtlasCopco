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
# 0.2 ESTILOS PREMIUM
# =============================================================================
def aplicar_estilos_premium():
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;600;800&display=swap');
        :root { --ac-blue: #007CA6; --ac-dark: #005675; --bhp-orange: #FF6600; }
        html, body, p, h1, h2, h3, h4, h5, h6, span, div { font-family: 'Montserrat', sans-serif; }
        
        [data-testid="stStatusWidget"] { visibility: hidden !important; display: none !important; }
        header { background: transparent !important; }
        [data-testid="stToolbar"] { visibility: hidden !important; display: none !important; } 
        [data-testid="stDecoration"] { display: none !important; }
        [data-testid="collapsedControl"] {
            display: flex !important; visibility: visible !important; opacity: 1 !important;
            z-index: 999999 !important; background-color: var(--ac-blue) !important; 
            border-radius: 8px !important; box-shadow: 0 4px 15px rgba(0, 124, 166, 0.4) !important;
            margin-top: 15px !important; margin-left: 15px !important; transition: all 0.3s ease !important;
        }
        [data-testid="collapsedControl"]:hover { background-color: var(--bhp-orange) !important; transform: scale(1.05) !important; }
        [data-testid="collapsedControl"] svg { fill: white !important; stroke: white !important; }
        a[href*="github.com"] { display: none !important; visibility: hidden !important; opacity: 0 !important; pointer-events: none !important; }
        [data-testid="viewerBadge"] {display: none !important;}
        div[class^="viewerBadge_container"] {display: none !important;}
        footer {display: none !important;} 
        
        div.stButton > button:first-child { background: linear-gradient(135deg, var(--ac-blue) 0%, var(--ac-dark) 100%); color: white; border-radius: 8px; border: none; font-weight: 600; padding: 0.6rem 1.2rem; transition: all 0.3s ease; box-shadow: 0 4px 15px rgba(0, 124, 166, 0.4); }
        div.stButton > button:first-child:hover { transform: translateY(-2px); box-shadow: 0 8px 20px rgba(0, 124, 166, 0.6); }
        [data-testid="stVerticalBlockBorderWrapper"] { background: linear-gradient(145deg, #1a212b, #151a22) !important; border-radius: 12px !important; border: 1px solid #2b3543 !important; transition: transform 0.3s ease, box-shadow 0.3s ease, border-color 0.3s ease !important; }
        [data-testid="stVerticalBlockBorderWrapper"]:hover { transform: translateY(-6px) !important; box-shadow: 0 10px 25px rgba(0, 124, 166, 0.25) !important; border-color: var(--ac-blue) !important; }
        .stTextInput>div>div>input, .stNumberInput>div>div>input, .stSelectbox>div>div>select, .stDateInput>div>div>input { border-radius: 6px !important; border: 1px solid #2b3543 !important; background-color: #1e2530 !important; color: white !important; }
        .stTextInput>div>div>input:focus, .stNumberInput>div>div>input:focus, .stSelectbox>div>div>select:focus, .stDateInput>div>div>input:focus { border-color: var(--bhp-orange) !important; box-shadow: 0 0 8px rgba(255, 102, 0, 0.3) !important; }
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

# --- FIN DE LA PARTE 1 ---
# =============================================================================
# 4. FUNCIONES AUXILIARES Y CEREBRO MATEMÁTICO MINERO
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
        wk_num = int(re.sub(r'\D', '', str(wk_string)))
        if wk_num >= 50: return datetime.date.fromisocalendar(2025, wk_num, 1)
        return datetime.date.fromisocalendar(2026, wk_num, 1)
    except: return None

def calcular_mes_minero(wk_string):
    if pd.isna(wk_string) or str(wk_string).strip() == "": return "Sin Asignar"
    d = wk_to_date(wk_string)
    if not d: return "Sin Asignar"
    meses_full = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    if d.day <= 15: return meses_full[d.month - 1]
    else: return meses_full[d.month if d.month < 12 else 0]

def get_current_wk():
    hoy = datetime.date.today()
    wk_num = hoy.isocalendar()[1]
    return f"WK{wk_num:02d}"

def formatear_wk(wk_str):
    if pd.isna(wk_str) or str(wk_str).strip() == "": return ""
    nums = re.findall(r'\d+', str(wk_str))
    if nums: return f"WK{int(nums[0]):02d}"
    return str(wk_str).upper()

def get_semanas_mes_minero(mes_nombre):
    if mes_nombre == "Todas" or mes_nombre == "Sin Asignar": return "Todas"
    meses_map_full = {"Enero": 1, "Febrero": 2, "Marzo": 3, "Abril": 4, "Mayo": 5, "Junio": 6, "Julio": 7, "Agosto": 8, "Septiembre": 9, "Octubre": 10, "Noviembre": 11, "Diciembre": 12}
    if mes_nombre not in meses_map_full: return ""
    m_num = meses_map_full[mes_nombre]
    y_num = 2025 if m_num == 12 else 2026
    if m_num == 1: min_d = datetime.date(y_num - 1, 12, 16)
    else: min_d = datetime.date(y_num, m_num - 1, 16)
    max_d = datetime.date(y_num, m_num, 15)
    return f"WK{min_d.isocalendar()[1]:02d} a WK{max_d.isocalendar()[1]:02d}"

# =============================================================================
# 5. MOTOR CMMS CON DATOS REALES (LIMPIOS)
# =============================================================================
@st.cache_data(ttl=60, show_spinner=False)
def cargar_cmms():
    headers = ["TAG", "S_Programada", "Tipo", "Estado", "S_Realizada", "Observacion"]
    datos_reales = [
        {"TAG": "70-GC-013", "S_Programada": "WK51", "Tipo": "P2", "Estado": "✅ Hecho", "S_Realizada": "2025-12-15", "Observacion": ""},
        {"TAG": "70-GC-013", "S_Programada": "WK02", "Tipo": "INSP", "Estado": "✅ Hecho", "S_Realizada": "2026-01-05", "Observacion": ""},
        {"TAG": "70-GC-013", "S_Programada": "WK04", "Tipo": "P1", "Estado": "✅ Hecho", "S_Realizada": "2026-01-19", "Observacion": ""},
        {"TAG": "70-GC-013", "S_Programada": "WK07", "Tipo": "P1", "Estado": "✅ Hecho", "S_Realizada": "2026-02-10", "Observacion": ""},
        {"TAG": "70-GC-013", "S_Programada": "WK11", "Tipo": "INSP", "Estado": "⏳ Pendiente", "S_Realizada": "", "Observacion": ""},
        {"TAG": "70-GC-014", "S_Programada": "WK52", "Tipo": "INSP", "Estado": "✅ Hecho", "S_Realizada": "2025-12-22", "Observacion": ""},
        {"TAG": "70-GC-014", "S_Programada": "WK02", "Tipo": "P2", "Estado": "✅ Hecho", "S_Realizada": "2026-01-05", "Observacion": ""},
        {"TAG": "70-GC-014", "S_Programada": "WK04", "Tipo": "INSP", "Estado": "🚨 F/S", "S_Realizada": "", "Observacion": ""}, 
        {"TAG": "70-GC-014", "S_Programada": "WK09", "Tipo": "INSP", "Estado": "✅ Hecho", "S_Realizada": "2026-02-23", "Observacion": ""},
        {"TAG": "70-GC-014", "S_Programada": "WK10", "Tipo": "INSP", "Estado": "⏳ Pendiente", "S_Realizada": "", "Observacion": ""},
        {"TAG": "50-GC-001", "S_Programada": "WK01", "Tipo": "P2", "Estado": "✅ Hecho", "S_Realizada": "2025-12-29", "Observacion": ""},
        {"TAG": "50-GC-001", "S_Programada": "WK04", "Tipo": "P1", "Estado": "✅ Hecho", "S_Realizada": "2026-01-21", "Observacion": ""},
        {"TAG": "50-GC-001", "S_Programada": "WK09", "Tipo": "P1", "Estado": "✅ Hecho", "S_Realizada": "2026-02-23", "Observacion": ""},
        {"TAG": "50-GC-001", "S_Programada": "WK10", "Tipo": "P3", "Estado": "⏳ Pendiente", "S_Realizada": "", "Observacion": ""},
        {"TAG": "50-GC-002", "S_Programada": "WK01", "Tipo": "INSP", "Estado": "✅ Hecho", "S_Realizada": "2025-12-29", "Observacion": ""},
        {"TAG": "50-GC-002", "S_Programada": "WK02", "Tipo": "P2", "Estado": "🚨 F/S", "S_Realizada": "", "Observacion": ""},
        {"TAG": "50-GC-002", "S_Programada": "WK04", "Tipo": "INSP", "Estado": "✅ Hecho", "S_Realizada": "2026-01-19", "Observacion": ""},
        {"TAG": "50-GC-002", "S_Programada": "WK09", "Tipo": "INSP", "Estado": "⏳ Pendiente", "S_Realizada": "", "Observacion": ""},
        {"TAG": "50-GC-003", "S_Programada": "WK01", "Tipo": "P2", "Estado": "✅ Hecho", "S_Realizada": "2025-12-29", "Observacion": ""},
        {"TAG": "50-GC-003", "S_Programada": "WK07", "Tipo": "P1", "Estado": "🚨 F/S", "S_Realizada": "", "Observacion": ""},
        {"TAG": "50-GC-003", "S_Programada": "WK11", "Tipo": "P1", "Estado": "⏳ Pendiente", "S_Realizada": "", "Observacion": ""},
        {"TAG": "55-GC-015", "S_Programada": "WK01", "Tipo": "P2", "Estado": "✅ Hecho", "S_Realizada": "2025-12-29", "Observacion": ""},
        {"TAG": "55-GC-015", "S_Programada": "WK06", "Tipo": "P1", "Estado": "✅ Hecho", "S_Realizada": "2026-02-04", "Observacion": ""},
        {"TAG": "55-GC-015", "S_Programada": "WK08", "Tipo": "INSP", "Estado": "✅ Hecho", "S_Realizada": "2026-02-16", "Observacion": ""},
        {"TAG": "65-GC-011", "S_Programada": "WK01", "Tipo": "P3", "Estado": "✅ Hecho", "S_Realizada": "2025-12-29", "Observacion": ""},
        {"TAG": "65-GC-011", "S_Programada": "WK05", "Tipo": "P1", "Estado": "✅ Hecho", "S_Realizada": "2026-01-28", "Observacion": ""},
        {"TAG": "65-GC-011", "S_Programada": "WK11", "Tipo": "INSP", "Estado": "✅ Hecho", "S_Realizada": "2026-03-09", "Observacion": ""},
        {"TAG": "35-GC-006", "S_Programada": "WK01", "Tipo": "P3", "Estado": "✅ Hecho", "S_Realizada": "2025-12-29", "Observacion": ""},
        {"TAG": "35-GC-006", "S_Programada": "WK02", "Tipo": "P1", "Estado": "🚨 F/S", "S_Realizada": "", "Observacion": ""},
        {"TAG": "35-GC-006", "S_Programada": "WK08", "Tipo": "INSP", "Estado": "✅ Hecho", "S_Realizada": "2026-02-16", "Observacion": ""}
    ]

    try:
        sheet = get_sheet("plan_cmms")
        if sheet:
            data = sheet.get_all_values()
            if len(data) > 0:
                df = pd.DataFrame(data[1:], columns=data[0]) if len(data) > 1 else pd.DataFrame(columns=data[0])
                if "S_Programada" in df.columns:
                    def migrar_fechas(row):
                        val = str(row.get('S_Realizada', ''))
                        if 'WK' in val.upper():
                            d = wk_to_date(row['S_Programada'])
                            return d.strftime("%Y-%m-%d") if d else ""
                        return val
                    df['S_Realizada'] = df.apply(migrar_fechas, axis=1)
                    df['Estado'] = df['Estado'].replace({'Hecho': '✅ Hecho', 'Pendiente': '⏳ Pendiente', 'F/S': '🚨 F/S', 'N/A': '⚪ N/A'})
                    return df
                sheet.clear(); df_base = pd.DataFrame(datos_reales, columns=headers)
                sheet.append_rows([headers] + df_base.values.tolist()); st.cache_data.clear(); return df_base
            else:
                df_base = pd.DataFrame(datos_reales, columns=headers)
                sheet.append_rows([headers] + df_base.values.tolist()); st.cache_data.clear(); return df_base
    except Exception as e: print(f"Error cargando CMMS: {e}")
    return pd.DataFrame(datos_reales, columns=headers)

def guardar_cmms(df):
    sheet = get_sheet("plan_cmms")
    if sheet:
        df_clean = df.copy()
        df_clean = df_clean.fillna("").astype(str)
        df_clean = df_clean.replace(["nan", "NaN", "NaT", "None", "<NA>"], "")
        sheet.clear()
        sheet.append_rows([df_clean.columns.values.tolist()] + df_clean.values.tolist())
        st.cache_data.clear()

def seleccionar_equipo(tag):
    st.session_state.equipo_seleccionado = tag; st.session_state.vista_firmas = False
    reg = buscar_ultimo_registro(tag)
    if reg:
        st.session_state.input_cliente = reg[1]; st.session_state.input_tec1 = reg[5]; st.session_state.input_tec2 = reg[6]
        st.session_state.input_estado = ""; st.session_state.input_reco = ""
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
    'input_reco': "", 'input_estado_eq': "Operativo", 'vista_firmas': False,
    'firma_tec_json': None, 'firma_tec_img': None
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
                        st.session_state.update({'logged_in': True, 'usuario_actual': u_in}); st.session_state.informes_pendientes = cargar_pendientes(u_in); st.rerun()
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
        
        if st.button("📊 Planificación", use_container_width=True, type="primary" if st.session_state.vista_actual == "planificacion" else "secondary"):
            st.session_state.vista_actual = "planificacion"; st.session_state.vista_firmas = False; st.session_state.equipo_seleccionado = None; st.rerun()
        
        if st.button("📜 Últimas Intervenciones", use_container_width=True, type="primary" if st.session_state.vista_actual == "historial" else "secondary"):
            st.session_state.vista_actual = "historial"; st.session_state.vista_firmas = False; st.session_state.equipo_seleccionado = None; st.rerun()
            
        if len(st.session_state.informes_pendientes) > 0:
            st.markdown("---")
            st.warning(f"📝 Tienes {len(st.session_state.informes_pendientes)} reportes esperando firmas.")
            if st.button("✍️ Ir a Pizarra de Firmas", use_container_width=True, type="primary" if st.session_state.vista_actual == "firmas" else "secondary"): 
                st.session_state.vista_firmas = True; st.session_state.vista_actual = "firmas"; st.session_state.equipo_seleccionado = None; st.rerun()
        st.markdown("---")
        if st.button("🚪 Cerrar Sesión", use_container_width=True): st.session_state.logged_in = False; st.rerun()

    # --- 7.0 VISTA: ÚLTIMAS INTERVENCIONES ---
    if st.session_state.vista_actual == "historial":
        st.markdown("""
            <div style="margin-top: 1rem; margin-bottom: 2.5rem; text-align: center; background: linear-gradient(90deg, rgba(255,102,0,0) 0%, rgba(255,102,0,0.15) 50%, rgba(255,102,0,0) 100%); padding: 20px; border-radius: 15px;">
                <h1 style="color: #FF6600; font-size: 3.5em; font-weight: 800; margin: 0; letter-spacing: -1px; text-transform: uppercase;">Muro de Intervenciones</h1>
                <p style="color: #8c9eb5; font-size: 1.2em; font-weight: 300; margin-top: -10px;">Registro Histórico en Tiempo Real</p>
            </div>
        """, unsafe_allow_html=True)
        
        def parse_fecha_historial(fecha_str):
            try:
                s = str(fecha_str).lower().strip()
                meses = {"enero":1, "ene":1, "febrero":2, "feb":2, "marzo":3, "mar":3, "abril":4, "abr":4, "mayo":5, "may":5, "junio":6, "jun":6, "julio":7, "jul":7, "agosto":8, "ago":8, "septiembre":9, "sep":9, "sept":9, "octubre":10, "oct":10, "noviembre":11, "nov":11, "diciembre":12, "dic":12}
                m = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)
                if m: return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
                m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})", s)
                if m: return datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
                nums = re.findall(r'\d+', s)
                words = re.findall(r'[a-z]+', s)
                if not nums: return datetime.date(1970,1,1)
                day = int(nums[0])
                year = int(nums[-1]) if len(nums)>1 else datetime.date.today().year
                if year < 100: year += 2000
                month = 1
                for w in words:
                    if w in meses:
                        month = meses[w]; break
                if day > 31: day, year = year, day
                return datetime.date(year, month, day)
            except: return datetime.date(1970, 1, 1)

        def format_fecha_historial(d):
            meses_nombres = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
            if d.year == 1970: return "Fecha Desconocida"
            return f"{d.day} de {meses_nombres[d.month]} de {d.year}"

        historial_global = obtener_historial_global()
        if not historial_global:
            st.info("Aún no hay reportes firmados y almacenados en la base de datos central.")
        else:
            historial_unico = []
            vistos = set()
            for item in historial_global:
                d_obj = parse_fecha_historial(item['fecha'])
                identificador = (item['tag'], d_obj)
                if identificador not in vistos:
                    vistos.add(identificador)
                    item['fecha_obj'] = d_obj
                    historial_unico.append(item)

            historial_agrupado = {}
            for item in historial_unico:
                f_obj = item['fecha_obj']
                if f_obj not in historial_agrupado: historial_agrupado[f_obj] = []
                historial_agrupado[f_obj].append(item)

            fechas_ordenadas = sorted(list(historial_agrupado.keys()), reverse=True)

            for d_obj in fechas_ordenadas:
                fecha_str = format_fecha_historial(d_obj)
                intervenciones = historial_agrupado[d_obj]
                
                st.markdown(f"<h3 style='color: white; border-bottom: 2px solid #2b3543; padding-bottom: 5px; margin-top: 15px;'>🗓️ {fecha_str}</h3>", unsafe_allow_html=True)
                columnas_muro = st.columns(3) 
                
                for idx, item in enumerate(intervenciones):
                    b_color = "#00e676" if item['estado'] == "Operativo" else "#ff1744"
                    bg_color = "rgba(0, 230, 118, 0.1)" if item['estado'] == "Operativo" else "rgba(255, 23, 68, 0.1)"
                    icono = "✅" if item['estado'] == "Operativo" else "🚨"
                    
                    cond_safe = str(item.get('condicion', '')).replace('\n', ' ').replace("'", '"')
                    reco_safe = str(item.get('reco', '')).replace('\n', ' ').replace("'", '"')
                    
                    cond_html = f"<hr style='margin: 8px 0; border-color: #2b3543;'><p style='margin: 5px 0 0 0; color: #aeb9cc; font-size: 0.8em; line-height: 1.3;'>📝 <b>Condición Final:</b> {cond_safe}</p>" if cond_safe.strip() else ""
                    reco_html = f"<p style='margin: 5px 0 0 0; color: #aeb9cc; font-size: 0.8em; line-height: 1.3;'>💡 <b>Nota:</b> {reco_safe}</p>" if reco_safe.strip() else ""

                    with columnas_muro[idx % 3]:
                        with st.container(border=True):
                            html_card = (
                                f"<div style='border-left: 5px solid {b_color}; padding-left: 12px; height: 100%; display: flex; flex-direction: column; justify-content: space-between;'>"
                                f"<div>"
                                f"<div style='display: flex; justify-content: space-between; align-items: flex-start;'>"
                                f"<h3 style='margin: 0; color: #007CA6; font-size: 1.4em;'>{item['tag']}</h3>"
                                f"<span style='background: #2b3543; color: white; padding: 3px 8px; border-radius: 12px; font-size: 0.75em; font-weight: bold;'>🛠️ {item['tipo']}</span>"
                                f"</div>"
                                f"<p style='margin: 2px 0 10px 0; color: #aeb9cc; font-size: 0.9em;'>{item['modelo']} &bull; {item['area'].title()}</p>"
                                f"<div style='background: #151a22; padding: 8px; border-radius: 8px; margin-bottom: 5px;'>"
                                f"<p style='margin: 0; color: #8c9eb5; font-size: 0.85em;'>🧑‍🔧 <b>Técnico:</b> {item['tecnico']}</p>"
                                f"{cond_html}"
                                f"{reco_html}"
                                f"</div>"
                                f"</div>"
                                f"<div style='margin-top: 10px; background: {bg_color}; border: 1px solid {b_color}; color: {b_color}; padding: 5px; border-radius: 6px; text-align: center; font-weight: bold; font-size: 0.85em;'>"
                                f"{icono} {item['estado']}"
                                f"</div>"
                                f"</div>"
                            )
                            st.markdown(html_card, unsafe_allow_html=True)

    # --- 7.1 VISTA PLANIFICACIÓN ---
    elif st.session_state.vista_actual == "planificacion":
        df_cmms = cargar_cmms()
        semana_actual = get_current_wk()
        df_cmms['S_Programada'] = df_cmms['S_Programada'].apply(formatear_wk)
        df_cmms['Mes_Calc'] = df_cmms['S_Programada'].apply(calcular_mes_minero)
        mes_de_hoy_full = calcular_mes_minero(semana_actual)
        
        if 'filtro_mes_activo' not in st.session_state:
            st.session_state.filtro_mes_activo = mes_de_hoy_full
            
        mes_visualizado = st.session_state.filtro_mes_activo if st.session_state.filtro_mes_activo != "Todas" else mes_de_hoy_full
        rango_semanas_header = get_semanas_mes_minero(mes_visualizado)
        
        st.markdown(f"""
            <div style="margin-top: 1rem; margin-bottom: 1rem; background: linear-gradient(90deg, rgba(0,124,166,0.1) 0%, rgba(0,124,166,0.2) 50%, rgba(0,124,166,0.1) 100%); padding: 20px; border-radius: 15px; border-left: 5px solid var(--ac-blue);">
                <h2 style="color: white; margin: 0;">📅 Panel de Control</h2>
                <p style="color: #8c9eb5; margin: 0; font-weight: 600;">Semanas del Mes: {rango_semanas_header} &nbsp;|&nbsp; Planificación Activa: {mes_visualizado}</p>
            </div>
        """, unsafe_allow_html=True)
        
        df_kpi = df_cmms[(df_cmms["Mes_Calc"] == mes_visualizado) & (df_cmms["Tipo"] != "N/A")]
        total_tareas = len(df_kpi)
        hechas = len(df_kpi[df_kpi["Estado"] == "✅ Hecho"])
        fs = len(df_kpi[df_kpi["Estado"] == "🚨 F/S"])
        pendientes = len(df_kpi[df_kpi["Estado"] == "⏳ Pendiente"])
        
        total_evaluable = hechas + pendientes
        cumplimiento = int((hechas / total_evaluable * 100)) if total_evaluable > 0 else (100 if hechas > 0 else 0)
        
        c_kpi1, c_kpi2, c_kpi3, c_kpi4 = st.columns(4)
        c_kpi1.metric(label="📈 Cumplimiento Mes", value=f"{cumplimiento}%")
        c_kpi2.metric(label="🎯 Tareas Programadas", value=total_tareas)
        c_kpi3.metric(label="✅ Tareas Completadas", value=hechas)
        c_kpi4.metric(label="🚨 Equipos F/S", value=fs)
        
        st.markdown("---")
        
        tab_gestion, tab_calendario, tab_matriz = st.tabs(["📋 Tablero", "📆 Calendario", "📊 Matriz de Mantenimiento"])
        
        with tab_gestion:
            c_f1, c_f2 = st.columns([1, 3])
            orden_meses_full = ["Todas", "Diciembre", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre"]
            
            filtro_mes = c_f1.selectbox("Filtrar por Mes:", orden_meses_full, key="filtro_mes_activo")
            
            min_date_val, max_date_val = None, None
            if filtro_mes != "Todas":
                meses_map_full = {"Enero": 1, "Febrero": 2, "Marzo": 3, "Abril": 4, "Mayo": 5, "Junio": 6, "Julio": 7, "Agosto": 8, "Septiembre": 9, "Octubre": 10, "Noviembre": 11, "Diciembre": 12}
                if filtro_mes in meses_map_full:
                    m_num = meses_map_full[filtro_mes]
                    y_num = 2025 if m_num == 12 else 2026
                    if m_num == 1:
                        min_date_val = datetime.date(y_num - 1, 12, 16)
                    else:
                        min_date_val = datetime.date(y_num, m_num - 1, 16)
                    max_date_val = datetime.date(y_num, m_num, 15)

            df_mostrar = df_cmms.copy() if filtro_mes == "Todas" else df_cmms[df_cmms["Mes_Calc"] == filtro_mes].copy()
            
            tags_presentes = df_mostrar['TAG'].tolist()
            todos_los_tags = list(inventario_equipos.keys())
            tags_faltantes = [t for t in todos_los_tags if t not in tags_presentes]
            
            if tags_faltantes:
                filas_vacias = pd.DataFrame([{"TAG": t, "S_Programada": "", "Tipo": "N/A", "Estado": "⚪ N/A", "S_Realizada": None, "Observacion": "", "Mes_Calc": filtro_mes if filtro_mes != "Todas" else "Sin Asignar"} for t in tags_faltantes])
                df_mostrar = pd.concat([df_mostrar, filas_vacias], ignore_index=True)
            df_mostrar = df_mostrar.sort_values(by="TAG").reset_index(drop=True)

            df_editado = pd.DataFrame()
            
            def safe_get_wk(x):
                if pd.isnull(x) or str(x).strip() in ["", "None", "NaT"]: return ""
                try:
                    if isinstance(x, str): x = datetime.datetime.strptime(x[:10], "%Y-%m-%d").date()
                    return f"WK{x.isocalendar()[1]:02d}"
                except: return ""

            def safe_date_str(x):
                if pd.isnull(x) or str(x).strip() in ["", "None", "NaT"]: return ""
                try:
                    if isinstance(x, str): return x[:10]
                    return x.strftime("%Y-%m-%d")
                except: return ""

            if not df_mostrar.empty:
                def string_to_date(val):
                    try: return datetime.datetime.strptime(str(val).strip(), "%Y-%m-%d").date()
                    except: return None
                
                df_mostrar['S_Realizada'] = df_mostrar['S_Realizada'].apply(string_to_date)
                df_mostrar['Día Programado'] = df_mostrar['S_Programada'].apply(wk_to_date)
                df_mostrar.insert(0, "🗑️ Quitar", False)
                
                if "kanban_table" in st.session_state:
                    edits = st.session_state["kanban_table"].get("edited_rows", {})
                    for idx_str, changes in edits.items():
                        if "Día Programado" in changes:
                            val = changes["Día Programado"]
                            if val is not None:
                                try:
                                    if isinstance(val, str): new_date = datetime.datetime.strptime(val[:10], "%Y-%m-%d").date()
                                    else: new_date = val
                                    wk_calculada = f"WK{new_date.isocalendar()[1]:02d}"
                                    df_mostrar.at[int(idx_str), 'S_Programada'] = wk_calculada
                                except: pass

                columnas_ordenadas = ["🗑️ Quitar", "TAG", "Día Programado", "Tipo", "Estado", "S_Realizada", "Observacion", "Mes_Calc", "S_Programada"]
                df_mostrar = df_mostrar[columnas_ordenadas]
                
                config_columnas = {
                    "🗑️ Quitar": st.column_config.CheckboxColumn("Quitar", default=False),
                    "TAG": st.column_config.TextColumn("Equipo", disabled=True),
                    "Mes_Calc": None, 
                    "S_Programada": None, 
                    "Día Programado": st.column_config.DateColumn("📆 Prog. para (Día y WK)", format="DD/MM/YYYY - [WK]WW", min_value=min_date_val, max_value=max_date_val, disabled=False),
                    "Tipo": st.column_config.SelectboxColumn("Intervención", options=["N/A", "INSP", "P1", "P2", "P3", "P4", "PM03"], disabled=False),
                    "Estado": st.column_config.SelectboxColumn("Estado Actual", options=["⚪ N/A", "⏳ Pendiente", "✅ Hecho", "🚨 F/S"], required=True),
                    "S_Realizada": st.column_config.DateColumn("Día Ejecución (Día y WK) 📅", format="DD/MM/YYYY - [WK]WW", disabled=False),
                    "Observacion": st.column_config.TextColumn("Comentarios")
                }
                def color_estado(val):
                    if val == '✅ Hecho': return 'background-color: #063f22; color: #6ee7b7; font-weight: bold;'
                    if val == '⏳ Pendiente': return 'background-color: #423205; color: #fde047; font-weight: bold;'
                    if val == '🚨 F/S': return 'background-color: #471015; color: #ff8a93; font-weight: bold;'
                    if val == '⚪ N/A': return 'color: #556b82; font-style: italic;'
                    return ''
                try: df_estilizado = df_mostrar.style.map(color_estado, subset=['Estado'])
                except AttributeError: df_estilizado = df_mostrar.style.applymap(color_estado, subset=['Estado'])
                
                df_editado = st.data_editor(df_estilizado, key="kanban_table", hide_index=True, use_container_width=True, column_config=config_columnas, height=750)
                
                if st.button("💾 Guardar Avances y Limpiar Tabla", type="primary"):
                    def get_final_wk(row):
                        d = row['Día Programado']
                        if pd.notnull(d) and str(d).strip() not in ["", "None", "NaT"]:
                            if isinstance(d, str): d = datetime.datetime.strptime(d[:10], "%Y-%m-%d").date()
                            return f"WK{d.isocalendar()[1]:02d}"
                        return row['S_Programada']
                        
                    df_editado['S_Programada'] = df_editado.apply(get_final_wk, axis=1)
                    df_editado['S_Realizada'] = df_editado['S_Realizada'].apply(safe_date_str)
                    
                    filas_validas = df_editado[
                        (df_editado["🗑️ Quitar"] == False) & (df_editado["Tipo"] != "N/A") & 
                        (df_editado["Estado"] != "⚪ N/A") & (df_editado["S_Programada"] != "")
                    ].copy()
                    
                    filas_validas.loc[(filas_validas['Estado'] == '✅ Hecho') & (filas_validas['S_Realizada'] == ""), 'S_Realizada'] = datetime.date.today().strftime("%Y-%m-%d")
                    
                    if filtro_mes == "Todas": df_cmms_final = filas_validas
                    else:
                        df_cmms_rest = df_cmms[df_cmms["Mes_Calc"] != filtro_mes]
                        df_cmms_final = pd.concat([df_cmms_rest, filas_validas], ignore_index=True)
                        
                    for col in ['Mes_Calc', '🗑️ Quitar', 'Día Programado']:
                        if col in df_cmms_final.columns: df_cmms_final = df_cmms_final.drop(columns=[col])
                    
                    guardar_cmms(df_cmms_final); st.success(f"✅ ¡Guardado!"); time.sleep(1.5); st.rerun()

            st.markdown("<br>", unsafe_allow_html=True)
            with st.expander("➕ Inyectar Tarea Extra", expanded=False):
                with st.form("form_nueva_tarea"):
                    c1, c2, c3 = st.columns(3)
                    n_tag = c1.selectbox("Equipo:", sorted(list(inventario_equipos.keys())))
                    n_tipo = c2.selectbox("Tipo de Tarea:", ["INSP", "P1", "P2", "P3", "P4", "PM03"])
                    
                    default_d = datetime.date.today()
                    if min_date_val and max_date_val:
                        if not (min_date_val <= default_d <= max_date_val): default_d = min_date_val
                        
                    n_fecha_prog = c3.date_input("📆 Día a Programar:", value=default_d, min_value=min_date_val, max_value=max_date_val)
                    n_obs = st.text_input("Observación inicial (Opcional):")
                    
                    if st.form_submit_button("🚀 Inyectar Tarea y Guardar Todo", type="primary", use_container_width=True):
                        df_cmms_guardar = df_cmms.copy()
                        if not df_editado.empty:
                            df_editado_clean = df_editado.copy()
                            def get_final_wk_clean(row):
                                d = row['Día Programado']
                                if pd.notnull(d) and str(d).strip() not in ["", "None", "NaT"]:
                                    if isinstance(d, str): d = datetime.datetime.strptime(d[:10], "%Y-%m-%d").date()
                                    return f"WK{d.isocalendar()[1]:02d}"
                                return row['S_Programada']
                            df_editado_clean['S_Programada'] = df_editado_clean.apply(get_final_wk_clean, axis=1)
                            df_editado_clean['S_Realizada'] = df_editado_clean['S_Realizada'].apply(safe_date_str)
                            filas_validas_f = df_editado_clean[
                                (df_editado_clean["🗑️ Quitar"] == False) & (df_editado_clean["Tipo"] != "N/A") & 
                                (df_editado_clean["Estado"] != "⚪ N/A") & (df_editado_clean["S_Programada"] != "")
                            ].copy()
                            filas_validas_f.loc[(filas_validas_f['Estado'] == '✅ Hecho') & (filas_validas_f['S_Realizada'] == ""), 'S_Realizada'] = datetime.date.today().strftime("%Y-%m-%d")
                            if filtro_mes == "Todas": df_cmms_guardar = filas_validas_f
                            else:
                                df_cmms_rest_f = df_cmms[df_cmms["Mes_Calc"] != filtro_mes]
                                df_cmms_guardar = pd.concat([df_cmms_rest_f, filas_validas_f], ignore_index=True)

                        n_sem_format = f"WK{n_fecha_prog.isocalendar()[1]:02d}"
                        nueva_fila = pd.DataFrame([{"TAG": n_tag, "S_Programada": n_sem_format, "Tipo": n_tipo, "Estado": "⏳ Pendiente", "S_Realizada": "", "Observacion": n_obs}])
                        for col in ['Mes_Calc', '🗑️ Quitar', 'Día Programado']:
                            if col in df_cmms_guardar.columns: df_cmms_guardar = df_cmms_guardar.drop(columns=[col])
                        
                        df_cmms_final_extra = pd.concat([df_cmms_guardar, nueva_fila], ignore_index=True)
                        guardar_cmms(df_cmms_final_extra); st.success("✅ Guardado."); time.sleep(1.5); st.rerun()

        with tab_calendario:
            opciones_meses_calendario = ["Diciembre 2025"] + [f"{m} 2026" for m in ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]]
            c_cal_tit, c_cal_sel = st.columns([2, 1])
            with c_cal_tit: st.markdown("### 📆 Calendario")
            with c_cal_sel:
                hoy_cal = datetime.date.today()
                mes_str = f"Diciembre 2025" if hoy_cal.year == 2025 and hoy_cal.month == 12 else f"{opciones_meses_calendario[hoy_cal.month]}" if hoy_cal.year == 2026 else "Enero 2026"
                mes_sel = st.selectbox("📅 Mes a visualizar:", opciones_meses_calendario, index=opciones_meses_calendario.index(mes_str) if mes_str in opciones_meses_calendario else 1)
                
            cal_year = 2025 if "2025" in mes_sel else 2026
            meses_nombres_cal = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
            cal_month = meses_nombres_cal.index(mes_sel.split(" ")[0]) + 1
                
            cal = calendar.Calendar(calendar.MONDAY)
            semanas_mes = cal.monthdatescalendar(cal_year, cal_month) 
            
            tareas_por_fecha = {}
            for _, row in df_cmms.iterrows():
                d_prog = wk_to_date(row['S_Programada'])
                d_target = None
                if row['Estado'] == '✅ Hecho' and str(row['S_Realizada']).strip() != "":
                    try: d_target = datetime.datetime.strptime(str(row['S_Realizada']).strip(), "%Y-%m-%d").date()
                    except: d_target = d_prog
                else: d_target = d_prog
                if d_target:
                    if d_target not in tareas_por_fecha: tareas_por_fecha[d_target] = []
                    tareas_por_fecha[d_target].append({"tag": row['TAG'], "tipo": row['Tipo'], "est": row['Estado']})
            
            html_cal = '<div style="display:grid; grid-template-columns: 65px repeat(7, 1fr); gap: 10px; margin-top:10px;">'
            html_cal += '<div style="text-align:center; color:#FF6600; font-weight:900; font-size:0.8rem; margin-top: 10px;">REF</div>'
            for d in ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]: 
                html_cal += f'<div style="text-align:center; color:#8c9eb5; font-weight:bold; font-size:0.9rem;">{d}</div>'
                
            for semana in semanas_mes:
                wk_num = semana[0].isocalendar()[1]
                if cal_year == 2025 and semana[0].month == 12: wk_num = semana[0].isocalendar()[1] 
                
                html_cal += f'<div style="display:flex; align-items:center; justify-content:center; background:#2b3543; border-radius:8px; border-left: 4px solid #FF6600; color:white; font-weight:bold; font-size:0.85rem; box-shadow: 0 4px 6px rgba(0,0,0,0.3); height: 120px;">WK{wk_num:02d}</div>'
                
                for dia in semana:
                    is_current_month = dia.month == cal_month
                    bg_color = "#1a212b" if is_current_month else "#11151c"
                    border_color = "#00BFFF" if dia == hoy_cal else "#2b3543"
                    html_cal += f'<div style="background:{bg_color}; border: 1px solid {border_color}; border-radius: 8px; padding: 5px; min-height: 120px;">'
                    html_cal += f'<div style="text-align:right; color:white; font-size:0.9rem; margin-bottom:8px;">{dia.day}</div>'
                    if dia in tareas_por_fecha:
                        for t in tareas_por_fecha[dia]:
                            c_bg = "transparent"; c_tx = "#8c9eb5"; b_style = "1px dashed #455065" 
                            if t['est'] == '✅ Hecho': c_bg, c_tx, b_style = "#063f22", "#6ee7b7", "1px solid #10b981"
                            elif t['est'] == '🚨 F/S': c_bg, c_tx, b_style = "#471015", "#ff8a93", "1px solid #ef4444"
                            elif t['tipo'] == 'P1': c_bg, c_tx, b_style = "#0c2d48", "#66c2ff", "1px solid #1a5c94"
                            elif t['tipo'] == 'P2': c_bg, c_tx, b_style = "#4a2c00", "#ffb04c", "1px solid #8c5300"
                            elif t['tipo'] == 'P3': c_bg, c_tx, b_style = "#301047", "#d78aff", "1px solid #622291"
                            elif t['tipo'] == 'P4': c_bg, c_tx, b_style = "#471015", "#ff8a93", "1px solid #8e202a"
                            html_cal += f'<div style="background:{c_bg}; color:{c_tx}; padding:4px; margin-bottom:4px; border-radius:4px; font-size:0.75rem; border: {b_style};"><b>{t["tag"]}</b> - {t["tipo"]}</div>'
                    html_cal += '</div>'
            html_cal += '</div>'
            st.markdown(html_cal, unsafe_allow_html=True)

        with tab_matriz:
            df_pivot_base = df_cmms[df_cmms['Tipo'] != 'N/A'].copy()
            df_pivot_base['Contenido'] = df_pivot_base['Tipo'] + "\n" + df_pivot_base['Estado'].apply(lambda x: str(x).split(" ")[1] if " " in str(x) else str(x))
            
            c_mat1, c_mat2 = st.columns([1.5, 2])
            with c_mat1: 
                vista_matriz = st.radio("Modo de Visualización:", ["🔍 Por Mes (Zoom In)", "📆 Anual (Semanas WK)", "📅 Anual (Por Meses)"], horizontal=True)
            
            def map_mes_full(q):
                if q == "Diciembre": return "dic-25"
                meses = {"Enero":"ene-26", "Febrero":"feb-26", "Marzo":"mar-26", "Abril":"abr-26", "Mayo":"may-26", "Junio":"jun-26", "Julio":"jul-26", "Agosto":"ago-26", "Septiembre":"sept-26", "Octubre":"oct-26", "Noviembre":"nov-26"}
                return meses.get(q, q)

            df_pivot_base['Mes_Vista'] = df_pivot_base['Mes_Calc'].apply(map_mes_full)
            
            if vista_matriz == "📅 Anual (Por Meses)":
                col_pivot = 'Mes_Vista'
                cols_todas = ["dic-25", "ene-26", "feb-26", "mar-26", "abr-26", "may-26", "jun-26", "jul-26", "ago-26", "sept-26", "oct-26", "nov-26"]
            else:
                col_pivot = 'S_Programada'
                semanas_brutas = ["WK51", "WK52"] + [f"WK{i:02d}" for i in range(1, 53)]
                cols_todas = list(dict.fromkeys(semanas_brutas))
                
            df_pivot = df_pivot_base.groupby(['TAG', col_pivot])['Contenido'].apply(lambda x: '\n---\n'.join(x)).unstack().fillna("")
            
            lista_info = []
            for t in df_pivot.index:
                if t in inventario_equipos: eq, _, area, _ = inventario_equipos[t]; lista_info.append({"TAG": t, "Equipo": eq, "Área": area.title()})
                else: lista_info.append({"TAG": t, "Equipo": "-", "Área": "-"})
            
            df_info = pd.DataFrame(lista_info).set_index("TAG")
            df_matriz = pd.concat([df_info, df_pivot], axis=1).reset_index()
            
            cols_base = ['TAG', 'Equipo', 'Área']
            
            for c in cols_todas:
                if c not in df_matriz.columns: df_matriz[c] = ""
                
            df_matriz = df_matriz[cols_base + cols_todas]
            
            cols_finales = cols_base.copy()
            if vista_matriz == "🔍 Por Mes (Zoom In)":
                wk_a_quincena = {wk: calcular_mes_minero(wk) for wk in cols_todas}
                with c_mat2:
                    orden_meses_zoom = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
                    q_unicas = list(set(wk_a_quincena.values()))
                    q_unicas.sort(key=lambda x: orden_meses_zoom.index(x) if x in orden_meses_zoom else 99)
                    quin_seleccionada = st.selectbox("Selecciona el Mes a enfocar:", q_unicas, index=q_unicas.index(mes_visualizado) if mes_visualizado in q_unicas else 0)
                wks_mostrar = [wk for wk, q in wk_a_quincena.items() if q == quin_seleccionada]
                cols_finales.extend(wks_mostrar)
            else:
                cols_finales.extend(cols_todas)
                
            df_matriz_final = df_matriz[cols_finales]
            df_matriz_congelada = df_matriz_final.set_index(['TAG', 'Equipo', 'Área'])
            
            def estilo_matriz_colores(val):
                v = str(val).upper()
                if not v or v == "NAN": return ''
                base = 'white-space: pre-wrap; line-height: 1.4; border-radius: 6px; padding: 6px; text-align: center; font-size: 0.85em; '
                if 'HECHO' in v: return base + 'background-color: #063f22; color: #6ee7b7; font-weight: bold; border-left: 4px solid #10b981;'
                if 'F/S' in v: return base + 'background-color: #471015; color: #ff8a93; font-weight: bold; border-left: 4px solid #ef4444;'
                if 'PENDIENTE' in v: 
                    if 'P1' in v: return base + 'background-color: #0c2d48; color: #66c2ff; font-weight: bold; border-left: 4px solid #eab308;'
                    if 'P2' in v: return base + 'background-color: #4a2c00; color: #ffb04c; font-weight: bold; border-left: 4px solid #eab308;'
                    if 'P3' in v: return base + 'background-color: #301047; color: #d78aff; font-weight: bold; border-left: 4px solid #eab308;'
                    if 'P4' in v: return base + 'background-color: #471015; color: #ff8a93; font-weight: bold; border-left: 4px solid #eab308;'
                    return base + 'background-color: #423205; color: #fde047; font-weight: bold; border-left: 4px solid #eab308;'
                return base + 'color: #8c9eb5; font-style: italic;'
                
            columnas_pintar = [c for c in cols_finales if c not in cols_base]
            if len(columnas_pintar) > 0:
                try: st.dataframe(df_matriz_congelada.style.map(estilo_matriz_colores, subset=columnas_pintar), use_container_width=True, height=600)
                except AttributeError: st.dataframe(df_matriz_congelada.style.applymap(estilo_matriz_colores, subset=columnas_pintar), use_container_width=True, height=600)
            else:
                st.dataframe(df_matriz_congelada, use_container_width=True, height=600)

    # --- 7.2 VISTA DE FIRMAS, EDICIÓN Y DESCARGAS ---
    elif st.session_state.vista_firmas or st.session_state.vista_actual == "firmas":
        c_v1, c_v2 = st.columns([1,4])
        with c_v1: 
            if st.button("⬅️ Volver", use_container_width=True): volver_catalogo(); st.rerun()
        with c_v2: st.markdown("<h1 style='margin-top:-15px;'>✍️ Pizarra de Firmas y Revisión</h1>", unsafe_allow_html=True)
        st.markdown("---")
        
        if len(st.session_state.informes_pendientes) == 0: st.info("🎉 ¡Excelente! No tienes ningún informe pendiente por firmar.")
        else:
            # 🔥 FIRMA TÉCNICO CENTRADA Y PREMIUM
            st.markdown("<h3 style='text-align: center; color: #007CA6;'>🧑‍🔧 Configuración de Mi Firma Fija (Técnico)</h3>", unsafe_allow_html=True)
            _, col_canvas, _ = st.columns([1, 2, 1])
            with col_canvas:
                with st.container(border=True):
                    st.markdown("<p style='text-align: center; color: #8c9eb5; margin-bottom: 10px;'>Dibuja tu firma una sola vez aquí. Se aplicará automáticamente a todos los informes que apruebes.</p>", unsafe_allow_html=True)
                    canvas_tec_global = st_canvas(stroke_width=4, stroke_color="#000", background_color="#fff", height=180, width=400, drawing_mode="freedraw", key="canvas_tec_global", initial_drawing=st.session_state.firma_tec_json if st.session_state.firma_tec_json else None)
                    c_btn1, c_btn2 = st.columns(2)
                    with c_btn1:
                        if st.button("💾 Guardar Mi Firma", use_container_width=True):
                            if canvas_tec_global.json_data is not None and len(canvas_tec_global.json_data.get("objects", [])) > 0:
                                st.session_state.firma_tec_json = canvas_tec_global.json_data
                                st.session_state.firma_tec_img = canvas_tec_global.image_data
                                st.success("✅ Firma guardada correctamente.")
                                time.sleep(1); st.rerun()
                            else: st.warning("⚠️ Dibuja tu firma.")
                    with c_btn2:
                        if st.button("🔄 Reiniciar Firma", use_container_width=True):
                            st.session_state.firma_tec_json = None
                            st.session_state.firma_tec_img = None
                            st.rerun()

            st.markdown("<br><hr style='border-color: #2b3543;'>", unsafe_allow_html=True)

            areas_agrupadas = {}
            for inf in st.session_state.informes_pendientes:
                macro_area = inventario_equipos[inf['tag']][3].title() if inf['tag'] in inventario_equipos else "General"
                if macro_area not in areas_agrupadas: areas_agrupadas[macro_area] = []
                areas_agrupadas[macro_area].append(inf)

            for macro_area, informes_area in areas_agrupadas.items():
                st.markdown(f"### 🏢 Reportes de {macro_area} ({len(informes_area)} pendientes)")
                with st.container(border=True):
                    for idx, inf in enumerate(informes_area):
                        c_exp, c_del = st.columns([11, 1])
                        with c_exp:
                            with st.expander(f"📝 Revisar Documento: {inf['tag']} ({inf['tipo_plan']})"):
                                tab_ver, tab_editar = st.tabs(["📄 Ver y Descargar Borrador", "✏️ Corregir Datos Faltantes / Erróneos"])
                                
                                with tab_ver:
                                    if inf.get('ruta_prev_pdf') and os.path.exists(inf['ruta_prev_pdf']):
                                        c_dl1, c_dl2 = st.columns(2)
                                        with c_dl1:
                                            with open(inf['ruta_prev_pdf'], "rb") as f:
                                                st.download_button("⬇️ Descargar Borrador (PDF)", f, file_name=f"Borrador_{inf['nombre_archivo_base'].replace('.docx', '.pdf')}", mime="application/pdf", key=f"dl_pdf_{inf['tag']}_{idx}")
                                        with c_dl2:
                                            ruta_docx_prev = os.path.join(RUTA_ONEDRIVE, f"PREVIEW_{inf['nombre_archivo_base']}")
                                            if os.path.exists(ruta_docx_prev):
                                                with open(ruta_docx_prev, "rb") as f:
                                                    st.download_button("⬇️ Descargar Borrador (Word Editable)", f, file_name=f"Borrador_{inf['nombre_archivo_base']}", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"dl_word_{inf['tag']}_{idx}")
                                        
                                        st.markdown("<br>", unsafe_allow_html=True)
                                        try: pdf_viewer(inf['ruta_prev_pdf'], width=950, height=900)
                                        except Exception as e: st.error(f"Error visor: {e}")
                                    else: st.warning("⚠️ Vista preliminar no disponible.")
                                    
                                with tab_editar:
                                    st.info("Si olvidaste algún dato o te equivocaste, corrígelo aquí abajo y presiona guardar. El PDF se actualizará automáticamente.")
                                    with st.form(f"edit_form_{inf['tag']}_{idx}"):
                                        c1, c2, c3 = st.columns(3)
                                        new_h_m = c1.number_input("Horas Marcha Totales", value=int(inf['context'].get('horas_marcha', 0)), step=1)
                                        new_h_c = c2.number_input("Horas en Carga", value=int(inf['context'].get('horas_carga', 0)), step=1)
                                        new_t_s = c3.text_input("Temp Salida (°C)", value=str(inf['context'].get('temp_salida', '0')))
                                        
                                        c4, c5 = st.columns(2)
                                        new_p_c = c4.text_input("P. Carga (con unidad)", value=str(inf['context'].get('p_carga', '')))
                                        new_p_d = c5.text_input("P. Descarga (con unidad)", value=str(inf['context'].get('p_descarga', '')))
                                        
                                        new_est_ent = st.text_area("Descripción Condición Final", value=str(inf['context'].get('estado_entrega', '')))
                                        new_reco = st.text_area("Recomendaciones / Acciones Pendientes", value=str(inf['context'].get('recomendaciones', '')))
                                        
                                        if st.form_submit_button("💾 Guardar Corrección y Regenerar PDF", type="primary"):
                                            inf['context']['horas_marcha'] = new_h_m
                                            inf['context']['horas_carga'] = new_h_c
                                            inf['context']['temp_salida'] = new_t_s
                                            inf['context']['p_carga'] = new_p_c
                                            inf['context']['p_descarga'] = new_p_d
                                            inf['context']['estado_entrega'] = new_est_ent
                                            inf['context']['recomendaciones'] = new_reco
                                            
                                            t_list = list(inf['tupla_db'])
                                            try: t_list[9] = float(new_t_s.replace(',', '.'))
                                            except: t_list[9] = 0.0
                                            t_list[10] = new_p_c
                                            t_list[11] = new_p_d
                                            t_list[12] = new_h_m
                                            t_list[13] = new_h_c
                                            t_list[14] = new_est_ent
                                            t_list[16] = new_reco
                                            inf['tupla_db'] = tuple(t_list)
                                            
                                            ruta_prev_docx = os.path.join(RUTA_ONEDRIVE, f"PREVIEW_{inf['nombre_archivo_base']}")
                                            doc_prev = DocxTemplate(inf['file_plantilla'])
                                            ctx_prev = inf['context'].copy()
                                            ctx_prev['firma_tecnico'] = ""
                                            ctx_prev['firma_cliente'] = ""
                                            doc_prev.render(ctx_prev)
                                            doc_prev.save(ruta_prev_docx)
                                            ruta_prev_pdf = convertir_a_pdf(ruta_prev_docx)
                                            if ruta_prev_pdf: inf['ruta_prev_pdf'] = ruta_prev_pdf
                                            
                                            guardar_pendientes(st.session_state.usuario_actual, st.session_state.informes_pendientes)
                                            st.success("✅ ¡Documento corregido exitosamente! Ve a la pestaña 'Ver y Descargar Borrador'.")
                                            time.sleep(1.5)
                                            st.rerun()

                        with c_del:
                            st.markdown("<div style='margin-top: 35px;'></div>", unsafe_allow_html=True)
                            if st.button("❌", key=f"del_{inf['tag']}_{idx}", help="Quitar este informe de la bandeja"):
                                st.session_state.informes_pendientes.remove(inf)
                                guardar_pendientes(st.session_state.usuario_actual, st.session_state.informes_pendientes) 
                                if len(st.session_state.informes_pendientes) == 0: volver_catalogo()
                                st.rerun()
                    
                    st.markdown("---")
                    
                    nombres_clientes = " y ".join(list(set([inf['cli'] for inf in informes_area if inf.get('cli')])))
                    if not nombres_clientes: nombres_clientes = "Cliente a cargo"
                    
                    st.markdown(f"<h3 style='text-align: center; color: #007CA6;'>Firma de Aprobación Final</h3>", unsafe_allow_html=True)
                    st.markdown(f"<p style='text-align: center; color: #8c9eb5; margin-top: -10px;'><b>Aprobador:</b> {nombres_clientes}</p>", unsafe_allow_html=True)
                    
                    _, col_firma, _ = st.columns([1, 2, 1])
                    with col_firma:
                        with st.container(border=True):
                            canvas_cli = st_canvas(stroke_width=4, stroke_color="#000", background_color="#f0f2f6", height=200, width=450, drawing_mode="freedraw", key=f"cli_{macro_area}")
                            st.markdown("<p style='text-align: center; font-size: 0.8em; color: gray;'>Coloque su firma en el recuadro superior</p>", unsafe_allow_html=True)
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button(f"🚀 Aprobar, Firmar y Subir Informes de {macro_area}", type="primary", use_container_width=True, key=f"btn_subir_{macro_area}"):
                        tec_ok = st.session_state.firma_tec_img is not None
                        cli_ok = canvas_cli.image_data is not None and canvas_cli.json_data is not None and len(canvas_cli.json_data.get("objects", [])) > 0
                        
                        if not tec_ok: st.warning("⚠️ Debes guardar primero tu Firma de Técnico en el panel superior de esta pantalla.")
                        elif not cli_ok: st.warning(f"⚠️ Falta la Firma de Aprobación de {nombres_clientes}.")
                        else:
                            def procesar_imagen_firma(img_data): img = Image.fromarray(img_data.astype('uint8'), 'RGBA'); img_io = io.BytesIO(); img.save(img_io, format='PNG'); img_io.seek(0); return img_io
                            
                            informes_finales = []
                            with st.spinner(f"Generando documentos sellados para {macro_area}..."):
                                try:
                                    for inf in informes_area:
                                        io_tec_local = procesar_imagen_firma(st.session_state.firma_tec_img)
                                        io_cli_local = procesar_imagen_firma(canvas_cli.image_data)
                                        
                                        doc = DocxTemplate(inf['file_plantilla']); context = inf['context']
                                        context['firma_tecnico'] = InlineImage(doc, io_tec_local, width=Mm(40)); context['firma_cliente'] = InlineImage(doc, io_cli_local, width=Mm(40)); doc.render(context); doc.save(inf['ruta_docx']); ruta_pdf_gen = convertir_a_pdf(inf['ruta_docx'])
                                        
                                        if ruta_pdf_gen: ruta_final = ruta_pdf_gen; nombre_final = inf['nombre_archivo_base'].replace(".docx", ".pdf")
                                        else: ruta_final = inf['ruta_docx']; nombre_final = inf['nombre_archivo_base']
                                        
                                        tupla_lista = list(inf['tupla_db']); tupla_lista[18] = ruta_final; guardar_registro(tuple(tupla_lista))
                                        informes_finales.append({"tag": inf['tag'], "tipo": inf['tipo_plan'], "ruta": ruta_final, "nombre_archivo": f"{macro_area}@@{inf['tag']}@@{nombre_final}"})
                                    
                                    exito, mensaje_correo = enviar_carrito_por_correo(MI_CORREO_CORPORATIVO, informes_finales)
                                    if exito: 
                                        st.success(f"✅ ¡Listos y enviados los reportes de {macro_area}!")
                                        for inf_enviado in informes_area:
                                            if inf_enviado in st.session_state.informes_pendientes: st.session_state.informes_pendientes.remove(inf_enviado)
                                        guardar_pendientes(st.session_state.usuario_actual, st.session_state.informes_pendientes) 
                                        time.sleep(2)
                                        if len(st.session_state.informes_pendientes) == 0: volver_catalogo()
                                        st.rerun()
                                    else: st.error(f"Error de red: {mensaje_correo}")
                                except Exception as e: st.error(f"Error procesando los PDFs: {e}")
                st.markdown("<br><br>", unsafe_allow_html=True)

    # --- 7.3 VISTA CATÁLOGO Y 7.4 FICHAS TÉCNICAS ---
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
        st.markdown("<br>", unsafe_allow_html=True)
        
        # 🔥 CATÁLOGO AGRUPADO POR ÁREAS
        equipos_filtrados = {}
        for tag, (modelo, serie, area, ubicacion) in inventario_equipos.items():
            es_secador = "CD" in modelo.upper()
            if filtro_tipo == "Compresores" and es_secador: continue
            if filtro_tipo == "Secadores" and not es_secador: continue
            if busqueda in tag.lower() or busqueda in area.lower() or busqueda in modelo.lower() or busqueda in ubicacion.lower():
                area_format = ubicacion.title()
                if area_format not in equipos_filtrados: equipos_filtrados[area_format] = []
                equipos_filtrados[area_format].append((tag, modelo, area))

        for area, equipos in sorted(equipos_filtrados.items()):
            st.markdown(f"<h4 style='color: #8c9eb5; margin-top: 25px; margin-bottom: 15px; border-bottom: 1px solid #2b3543; padding-bottom: 5px;'>📍 Área: {area}</h4>", unsafe_allow_html=True)
            columnas = st.columns(4)
            for idx, (tag, modelo, area_eq) in enumerate(equipos):
                estado = estados_db.get(tag, "Operativo")
                if estado == "Operativo": color_borde = "#00e676"; badge_html = "<div style='background: rgba(0,230,118,0.15); color: #00e676; border: 1px solid #00e676; padding: 4px 10px; border-radius: 20px; font-size: 0.75rem; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; display: inline-block;'>OPERATIVO</div>"
                else: color_borde = "#ff1744"; badge_html = "<div style='background: rgba(255,23,68,0.15); color: #ff1744; border: 1px solid #ff1744; padding: 4px 10px; border-radius: 20px; font-size: 0.75rem; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; display: inline-block;'>FUERA DE SERVICIO</div>"
                
                with columnas[idx % 4]:
                    with st.container(border=True):
                        st.markdown(f"<div style='border-top: 4px solid {color_borde}; padding-top: 10px; text-align: center; margin-top:-10px;'>{badge_html}</div>", unsafe_allow_html=True)
                        st.button(f"{tag}", key=f"btn_{tag}", on_click=seleccionar_equipo, args=(tag,), use_container_width=True)
                        st.markdown(f"<p style='color: #8c9eb5; margin-top: 5px; font-size: 0.85rem; text-align: center;'><strong style='color:#007CA6;'>{modelo}</strong> &bull; {area_eq.title()}</p>", unsafe_allow_html=True)

    elif st.session_state.equipo_seleccionado is not None:
        tag_sel = st.session_state.equipo_seleccionado; mod_d, ser_d, area_d, ubi_d = inventario_equipos[tag_sel]
        c_btn, c_tit = st.columns([1, 4])
        with c_btn: st.button("⬅️ Volver", on_click=volver_catalogo, use_container_width=True)
        with c_tit: st.markdown(f"<h1 style='margin-top:-15px;'>⚙️ Ficha de Serviço: <span style='color:#007CA6;'>{tag_sel}</span></h1>", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        
        ESPECIFICACIONES = obtener_especificaciones(DEFAULT_SPECS)
        tab1, tab2, tab3, tab4 = st.tabs(["📋 1. Reporte y Diagnóstico", "📚 2. Ficha Técnica", "🔍 3. Bitácora de Observaciones", "👤 4. Gestión de Área"])
        
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
                actualizar_estado_equipo_en_nube(tag_sel, est_eq)
                if "CD" in tag_sel: file_plantilla = "plantilla/secadorfueradeservicio.docx" if est_eq == "Fuera de servicio" else "plantilla/inspeccionsecador.docx"
                else: file_plantilla = "plantilla/fueradeservicio.docx" if est_eq == "Fuera de servicio" else f"plantilla/{tipo_plan.lower()}.docx" if tipo_plan in ["P1", "P2", "P3"] else "plantilla/inspeccion.docx"
                context = {"tipo_intervencion": tipo_plan, "modelo": mod_d, "tag": tag_sel, "area": area_d, "ubicacion": ubi_d, "cliente_contacto": cli_cont, "p_carga": f"{p_c_clean} {unidad_p}", "p_descarga": f"{p_d_clean} {unidad_p}", "temp_salida": t_salida_clean, "horas_marcha": int(h_m), "horas_carga": int(h_c), "tecnico_1": tec1, "tecnico_2": tec2, "estado_equipo": est_eq, "estado_entrega": est_ent, "recomendaciones": reco, "serie": ser_d, "tipo_orden": tipo_plan.upper(), "fecha": fecha, "equipo_modelo": mod_d}; nombre_archivo = f"Informe_{tipo_plan}_{tag_sel}_{fecha.replace(' ','_')}.docx"; ruta = os.path.join(RUTA_ONEDRIVE, nombre_archivo); temp_db = float(t_salida_clean) if t_salida_clean.replace('.', '', 1).isdigit() else 0.0; tupla_db = (tag_sel, mod_d, ser_d, area_d, ubi_d, fecha, cli_cont, tec1, tec2, temp_db, f"{p_c_clean} {unidad_p}", f"{p_d_clean} {unidad_p}", h_m, h_c, est_ent, tipo_plan, reco, est_eq, "", st.session_state.usuario_actual)
                with st.spinner("Creando borrador del documento para vista preliminar..."):
                    doc_prev = DocxTemplate(file_plantilla); ctx_prev = context.copy(); ctx_prev['firma_tecnico'] = ""; ctx_prev['firma_cliente'] = ""; doc_prev.render(ctx_prev); os.makedirs(RUTA_ONEDRIVE, exist_ok=True); ruta_prev_docx = os.path.join(RUTA_ONEDRIVE, f"PREVIEW_{nombre_archivo}"); doc_prev.save(ruta_prev_docx); ruta_prev_pdf = convertir_a_pdf(ruta_prev_docx)
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