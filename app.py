import streamlit as st

# 🔥 CONFIGURACIÓN DE PÁGINA
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
    cuerpo = f"Estimado/a,\n\nSe adjuntan {len(lista_informes)} reportes de servicio técnico (Firmados) generados en la presente jornada.\n\nEquipos intervenidos:\n"
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
        return True, "✅ Enviado."
    except Exception as e: return False, f"❌ Error: {e}"

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
        [data-testid="collapsedControl"] { display: flex !important; background-color: var(--ac-blue) !important; border-radius: 8px !important; margin-top: 15px !important; margin-left: 15px !important; }
        [data-testid="collapsedControl"] svg { fill: white !important; stroke: white !important; }
        footer {display: none !important;} 
        div.stButton > button:first-child { background: linear-gradient(135deg, var(--ac-blue) 0%, var(--ac-dark) 100%); color: white; border-radius: 8px; border: none; font-weight: 600; }
        [data-testid="stVerticalBlockBorderWrapper"] { background: linear-gradient(145deg, #1a212b, #151a22) !important; border-radius: 12px !important; border: 1px solid #2b3543 !important; }
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
    "GA 18": {"Litros de Aceite": "14.1 L", "Manual": "manuales/manual_ga18.pdf"},
    "GA 30": {"Litros de Aceite": "14.6 L", "Manual": "manuales/manual_ga30.pdf"},
    "GA 37": {"Litros de Aceite": "14.6 L", "Manual": "manuales/manual_ga37.pdf"},
    "GA 45": {"Litros de Aceite": "17.9 L", "Manual": "manuales/manual_ga45.pdf"},
    "GA 75": {"Litros de Aceite": "35.2 L", "Manual": "manuales/manual_ga75.pdf"},
    "GA 90": {"Litros de Aceite": "69 L", "Manual": "manuales/manual_ga90.pdf"},
    "GA 132": {"Litros de Aceite": "93 L", "Manual": "manuales/manual_ga132.pdf"},
    "GA 250": {"Litros de Aceite": "130 L", "Manual": "manuales/manual_ga250.pdf"},
    "ZT 37": {"Litros de Aceite": "23 L", "Manual": "manuales/manual_zt37.pdf"},
    "CD 80+": {"Manual": "manuales/manual_cd80.pdf"},
    "CD 630": {"Manual": "manuales/manual_cd630.pdf"}
}

inventario_equipos = {
    "20-GC-001": ["GA 75", "AII482673", "truck shop", "Mina"], "20-GC-002": ["GA 75", "AII482674", "truck shop", "Mina"], "20-GC-003": ["GA 90", "AIF095178", "truck shop", "Mina"], "20-GC-004": ["GA 37", "AII390776", "truck shop", "Mina"],
    "35-GC-006": ["GA 250", "AIF095420", "chancado secundario", "Área Seca"], "35-GC-007": ["GA 250", "AIF095421", "chancado secundario", "Área Seca"], "35-GC-008": ["GA 250", "AIF095302", "chancado secundario", "Área Seca"],
    "50-GC-001": ["GA 45", "API542705", "planta SX", "Área Húmeda"], "50-GC-002": ["GA 45", "API542706", "planta SX", "Área Húmeda"], "50-GC-003": ["ZT 37", "API791692", "planta SX", "Área Húmeda"], "50-GC-004": ["ZT 37", "API791693", "planta SX", "Área Húmeda"], "50-CD-001": ["CD 80+", "API095825", "planta SX", "Área Húmeda"], "50-CD-002": ["CD 80+", "API095826", "planta SX", "Área Húmeda"],
    "55-GC-015": ["GA 30", "API501440", "planta borra", "Área Húmeda"],
    "65-GC-009": ["GA 250", "APF253608", "patio de estanques", "Área Húmeda"], "65-GC-011": ["GA 250", "APF253581", "patio de estanques", "Área Húmeda"], "65-CD-011": ["CD 630", "WXF300015", "patio de estanques", "Área Húmeda"], "65-CD-012": ["CD 630", "WXF300016", "patio de estanques", "Área Húmeda"],
    "70-GC-013": ["GA 132", "AIF095296", "descarga de acido", "Área Húmeda"], "70-GC-014": ["GA 132", "AIF095297", "descarga de acido", "Área Húmeda"],
    "Taller": ["GA 18", "API335343", "Taller", "Laboratorio"] # 🔥 APLICADO: Laboratorio
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
import streamlit as st
# Configuración inicial (Debe ser la línea 1 de código Streamlit)
st.set_page_config(page_title="Atlas Spence | Gestión de Reportes", layout="wide", page_icon="⚙️", initial_sidebar_state="expanded")

import os, subprocess, smtplib, time, json, uuid, io, re, calendar, datetime
import pandas as pd
import gspread
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from streamlit_drawable_canvas import st_canvas
from PIL import Image
from google.oauth2.service_account import Credentials
from streamlit_pdf_viewer import pdf_viewer

# --- 1. CONFIGURACIÓN GLOBAL Y CORREO ---
RUTA_ONEDRIVE = "Reportes_Temporales" 
MI_CORREO_CORPORATIVO = "ignacio.a.morales@atlascopco.com"
CORREO_REMITENTE = "informeatlas.spence@gmail.com"
PASSWORD_APLICACION = "jbumdljbdpyomnna"

def enviar_carrito_por_correo(destinatario, lista_informes):
    """Envía los reportes firmados adjuntos vía SMTP de Gmail."""
    msg = MIMEMultipart()
    msg['From'], msg['To'] = CORREO_REMITENTE, destinatario
    msg['Subject'] = f"REVISIÓN PREVIA: Reportes Atlas Copco - Firmados - {pd.Timestamp.now().strftime('%d/%m/%Y')}"
    cuerpo = f"Estimado/a,\n\nSe adjuntan {len(lista_informes)} reportes de servicio técnico (Firmados) generados en la presente jornada.\n\nEquipos intervenidos:\n" + "".join([f"- TAG: {i['tag']} | Orden: {i['tipo']}\n" for i in lista_informes]) + "\nSaludos cordiales,\nSistema Integrado InforGem"
    msg.attach(MIMEText(cuerpo, 'plain'))
    
    for item in lista_informes:
        if os.path.exists(item['ruta']):
            with open(item['ruta'], "rb") as f: part = MIMEBase('application', 'octet-stream'); part.set_payload(f.read())
            encoders.encode_base64(part)
            n_seguro = item["nombre_archivo"].translate(str.maketrans("áéíóúÁÉÍÓÚ", "aeiouAEIOU"))
            part.add_header('Content-Disposition', f'attachment; filename="{n_seguro}"')
            msg.attach(part)
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls(); server.login(CORREO_REMITENTE, PASSWORD_APLICACION); server.send_message(msg)
        return True, "✅ Enviado."
    except Exception as e: return False, f"❌ Error: {e}"

# --- 2. INYECCIÓN CSS FRONTEND ---
def aplicar_estilos_premium():
    """Oculta elementos nativos de Streamlit y aplica el UI corporativo."""
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;600;800&display=swap');
        :root { --ac-blue: #007CA6; --ac-dark: #005675; --bhp-orange: #FF6600; }
        html, body, p, h1, h2, h3, h4, h5, h6, span, div { font-family: 'Montserrat', sans-serif; }
        [data-testid="stStatusWidget"], header, [data-testid="stToolbar"], [data-testid="stDecoration"], footer, [data-testid="viewerBadge"], div[class^="viewerBadge_container"], a[href*="github.com"] { display: none !important; visibility: hidden !important; }
        [data-testid="collapsedControl"] { display: flex !important; background-color: var(--ac-blue) !important; border-radius: 8px !important; margin-top: 15px !important; margin-left: 15px !important; transition: all 0.3s ease !important; }
        [data-testid="collapsedControl"]:hover { background-color: var(--bhp-orange) !important; transform: scale(1.05) !important; }
        [data-testid="collapsedControl"] svg { fill: white !important; stroke: white !important; }
        div.stButton > button:first-child { background: linear-gradient(135deg, var(--ac-blue) 0%, var(--ac-dark) 100%); color: white; border-radius: 8px; border: none; font-weight: 600; padding: 0.6rem 1.2rem; box-shadow: 0 4px 15px rgba(0, 124, 166, 0.4); transition: all 0.3s ease;}
        div.stButton > button:first-child:hover { transform: translateY(-2px); box-shadow: 0 8px 20px rgba(0, 124, 166, 0.6); }
        [data-testid="stVerticalBlockBorderWrapper"] { background: linear-gradient(145deg, #1a212b, #151a22) !important; border-radius: 12px !important; border: 1px solid #2b3543 !important; }
        .stTextInput>div>div>input, .stNumberInput>div>div>input, .stSelectbox>div>div>select, .stDateInput>div>div>input { border-radius: 6px !important; border: 1px solid #2b3543 !important; background-color: #1e2530 !important; color: white !important; }
        .stTextInput>div>div>input:focus, .stNumberInput>div>div>input:focus, .stSelectbox>div>div>select:focus, .stDateInput>div>div>input:focus { border-color: var(--bhp-orange) !important; box-shadow: 0 0 8px rgba(255, 102, 0, 0.3) !important; }
        .stTabs [data-baseweb="tab-list"] { border-bottom: 2px solid #2b3543; }
        .stTabs [aria-selected="true"] { color: var(--bhp-orange) !important; border-bottom: 3px solid var(--bhp-orange) !important; }
        </style>
    """, unsafe_allow_html=True)
aplicar_estilos_premium()

# --- 3. DATOS MAESTROS E INVENTARIO ---
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

# --- 4. MOTOR DB (CACHÉ CENTRALIZADO) ---
@st.cache_resource(show_spinner=False)
def get_gspread_client():
    creds_dict = json.loads(os.environ.get("gcp_json", st.secrets.get("gcp_json", "{}")))
    return gspread.authorize(Credentials.from_service_account_info(creds_dict, scopes=['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']))

def get_sheet(sheet_name):
    try:
        client = get_gspread_client(); doc = client.open("BaseDatos")
        try: return doc.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound: return doc.add_worksheet(title=sheet_name, rows="1000", cols="20")
    except: return None

@st.cache_data(ttl=120, show_spinner=False)
def fetch_all_data(sheet_name):
    """Descarga toda la hoja a memoria para evitar llamadas redundantes a la API."""
    sheet = get_sheet(sheet_name)
    return sheet.get_all_values() if sheet else []

def clear_db_cache():
    fetch_all_data.clear(); st.cache_data.clear()

def add_row_db(sheet_name, row_data):
    """Controlador unificado de guardado para todas las tablas."""
    sheet = get_sheet(sheet_name)
    if sheet: sheet.append_row(row_data); clear_db_cache()

# --- Funciones de lectura rápida usando Caché ---
def obtener_estados_actuales():
    estados = {r[0]: r[17] for r in fetch_all_data("intervenciones") if len(r) >= 18}
    estados.update({r[0]: r[1] for r in fetch_all_data("estados_equipos") if len(r) >= 2})
    return estados

def actualizar_estado_equipo_en_nube(tag, nuevo_estado):
    sheet = get_sheet("estados_equipos"); data = sheet.get_all_values() if sheet else []
    fila = next((i + 1 for i, r in enumerate(data) if len(r) > 0 and r[0] == tag), -1)
    if fila != -1: sheet.update_cell(fila, 2, nuevo_estado)
    else: sheet.append_row([tag, nuevo_estado])
    clear_db_cache()

def obtener_datos_equipo(tag):
    return {r[1]: r[2] for r in fetch_all_data("datos_equipo") if len(r) >= 3 and r[0] == tag}

def obtener_observaciones(tag):
    obs = [{"id": r[0], "fecha": r[2], "usuario": r[3], "texto": r[4]} for r in fetch_all_data("observaciones") if len(r) >= 6 and r[1] == tag and r[5] == "ACTIVO"]
    return pd.DataFrame(obs).iloc[::-1] if obs else pd.DataFrame(columns=["id", "fecha", "usuario", "texto"])

def obtener_contactos():
    con = [r[0] for r in fetch_all_data("contactos") if len(r) > 1 and r[1] == "ACTIVO"]
    return sorted(list(set(con))) if con else ["Lorena Rojas"]

def obtener_especificaciones(defaults):
    sp = {k: dict(v) for k, v in defaults.items()}
    for r in fetch_all_data("especificaciones"):
        if len(r) >= 3:
            if r[0] not in sp: sp[r[0]] = {}
            sp[r[0]][r[1]] = r[2]
    return sp

def buscar_ultimo_registro(tag):
    for r in reversed(fetch_all_data("intervenciones")):
        if len(r) >= 20 and r[0] == tag: return (r[5], r[6], r[9], r[14], r[15], r[7], r[8], r[10], r[11], r[12], r[13], r[16], r[17])
    return None

def obtener_todo_el_historial(tag):
    h = [{"fecha": r[5], "tipo_intervencion": r[15], "estado_equipo": r[17], "Cuenta Usuario": r[19], "horas_marcha": r[12], "p_carga": r[10], "temp_salida": r[9]} for r in fetch_all_data("intervenciones") if len(r) >= 20 and r[0] == tag]
    return pd.DataFrame(h).iloc[::-1] if h else pd.DataFrame()

def obtener_historial_global():
    h = []
    for r in reversed(fetch_all_data("intervenciones")):
        if len(r) >= 20 and r[0] != "TAG":
            h.append({"tag": r[0], "modelo": r[1], "area": r[3], "fecha": r[5], "tecnico": r[7], "tipo": r[15], "estado": r[17], "condicion": r[14], "reco": r[16]})
            if len(h) >= 50: break
    return h

def eliminar_observacion(id_obs):
    sheet = get_sheet("observaciones")
    if sheet:
        cell = sheet.find(id_obs)
        if cell: sheet.update_cell(cell.row, 6, "ELIMINADO"); clear_db_cache()

def eliminar_contacto(nombre):
    sheet = get_sheet("contactos")
    if sheet:
        cells = sheet.findall(nombre)
        for cell in cells: sheet.update_cell(cell.row, 2, "ELIMINADO")
        clear_db_cache()

def guardar_registro(data_tuple):
    try: add_row_db("intervenciones", [str(x) for x in data_tuple]); return True
    except: return False
# --- 5. FUNCIONES AUXILIARES DOCUMENTALES Y FECHAS ---
def convertir_a_pdf(ruta_docx):
    ruta_pdf = ruta_docx.replace(".docx", ".pdf"); abs_path = os.path.abspath(ruta_docx)
    try: subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', abs_path, '--outdir', os.path.dirname(abs_path)], capture_output=True); return ruta_pdf if os.path.exists(ruta_pdf) else None
    except: pass
    try: import pythoncom; from docx2pdf import convert; pythoncom.CoInitialize(); convert(abs_path, ruta_pdf); return ruta_pdf if os.path.exists(ruta_pdf) else None
    except: return None

def obtener_fecha_hoy_esp():
    m = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
    return f"{pd.Timestamp.now().day} de {m[pd.Timestamp.now().month]} de {pd.Timestamp.now().year}"

def cargar_pendientes(user):
    p = os.path.join(RUTA_ONEDRIVE, f"bandeja_{user.replace(' ', '_')}.json")
    try: return json.load(open(p, "r", encoding="utf-8")) if os.path.exists(p) else []
    except: return []

def guardar_pendientes(user, pend):
    os.makedirs(RUTA_ONEDRIVE, exist_ok=True)
    try: json.dump(pend, open(os.path.join(RUTA_ONEDRIVE, f"bandeja_{user.replace(' ', '_')}.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=4)
    except: pass

def wk_to_date(wk_str):
    try: 
        w = int(re.sub(r'\D', '', str(wk_str)))
        return datetime.date.fromisocalendar(2025 if w >= 50 else 2026, w, 1)
    except: return None

def calcular_mes_minero(wk_str):
    """Calcula a qué mes pertenece una Semana según el ciclo minero (16 al 15)."""
    d = wk_to_date(wk_str)
    if not d: return "Sin Asignar"
    m_full = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    return m_full[d.month - 1] if d.day <= 15 else m_full[d.month if d.month < 12 else 0]

def formatear_wk(wk_str):
    n = re.findall(r'\d+', str(wk_str))
    return f"WK{int(n[0]):02d}" if n else str(wk_str).upper()

def safe_date_str(x):
    try: return x[:10] if isinstance(x, str) else x.strftime("%Y-%m-%d")
    except: return ""

def parse_fecha(f_str):
    try:
        s = str(f_str).lower().strip(); meses = {"ene":1,"feb":2,"mar":3,"abr":4,"may":5,"jun":6,"jul":7,"ago":8,"sep":9,"oct":10,"nov":11,"dic":12}
        m1 = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", s); m2 = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})", s)
        if m1: return datetime.date(int(m1.group(1)), int(m1.group(2)), int(m1.group(3)))
        if m2: return datetime.date(int(m2.group(3)), int(m2.group(2)), int(m2.group(1)))
        nums = re.findall(r'\d+', s); words = re.findall(r'[a-z]+', s)
        if not nums: return datetime.date(1970,1,1)
        day, year = int(nums[0]), int(nums[-1]) if len(nums)>1 else datetime.date.today().year
        month = next((meses[w[:3]] for w in words if w[:3] in meses), 1)
        return datetime.date(year+2000 if year<100 else year, month, day if day<=31 else year)
    except: return datetime.date(1970, 1, 1)

# --- 6. MOTOR CMMS Y MATRIZ ---
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
    data = fetch_all_data("plan_cmms")
    if len(data) > 1:
        df = pd.DataFrame(data[1:], columns=data[0])
        df['S_Realizada'] = df.apply(lambda r: wk_to_date(r['S_Programada']).strftime("%Y-%m-%d") if 'WK' in str(r.get('S_Realizada', '')).upper() else r.get('S_Realizada', ''), axis=1)
        df['Estado'] = df['Estado'].replace({'Hecho': '✅ Hecho', 'Pendiente': '⏳ Pendiente', 'F/S': '🚨 F/S', 'N/A': '⚪ N/A'})
        return df
    if get_sheet("plan_cmms"):
        get_sheet("plan_cmms").clear(); get_sheet("plan_cmms").append_rows([headers] + pd.DataFrame(datos_reales).values.tolist())
    return pd.DataFrame(datos_reales, columns=headers)

def guardar_cmms(df):
    sheet = get_sheet("plan_cmms")
    if sheet:
        df_clean = df.fillna("").astype(str).replace(["nan", "NaN", "NaT", "None", "<NA>"], "")
        sheet.clear(); sheet.append_rows([df_clean.columns.tolist()] + df_clean.values.tolist())
        clear_db_cache()

def seleccionar_equipo(tag):
    st.session_state.update({'equipo_seleccionado': tag, 'vista_firmas': False, 'input_estado': "", 'input_reco': ""})
    reg = buscar_ultimo_registro(tag)
    if reg:
        st.session_state.update({'input_cliente': reg[1], 'input_tec1': reg[5], 'input_tec2': reg[6], 'input_estado_eq': reg[12] or "Operativo", 'input_h_marcha': int(reg[9]) if reg[9] else 0, 'input_h_carga': int(reg[10]) if reg[10] else 0, 'input_temp': str(reg[2]).replace(',', '.') if reg[2] else "70.0"})
        try: st.session_state.input_p_carga = str(reg[7]).split()[0].replace(',', '.')
        except: st.session_state.input_p_carga = "7.0"
        try: st.session_state.input_p_descarga = str(reg[8]).split()[0].replace(',', '.')
        except: st.session_state.input_p_descarga = "7.5"
    else: st.session_state.input_estado_eq = "Operativo"

def volver_catalogo(): st.session_state.update({'equipo_seleccionado': None, 'vista_firmas': False, 'vista_actual': "catalogo"})

defaults = {'logged_in': False, 'usuario_actual': "", 'equipo_seleccionado': None, 'vista_actual': "catalogo", 'input_cliente': "Lorena Rojas", 'input_tec1': "", 'input_tec2': "", 'input_h_marcha': 0, 'input_h_carga': 0, 'input_temp': "70.0", 'input_p_carga': "7.0", 'input_p_descarga': "7.5", 'input_estado': "", 'input_reco': "", 'input_estado_eq': "Operativo", 'vista_firmas': False, 'firma_tec_json': None, 'firma_tec_img': None, 'informes_pendientes': []}
for k, v in defaults.items(): st.session_state.setdefault(k, v)

# --- 7. VISTA: LOGIN Y RUTEO ---
if not st.session_state.logged_in:
    st.markdown("<br><br><br>", unsafe_allow_html=True); _, col_c, _ = st.columns([1, 1.5, 1])
    with col_c:
        with st.container(border=True):
            st.markdown("<h1 style='text-align: center; border-bottom:none;'>⚙️ <span style='color:#007CA6;'>Atlas</span> <span style='color:#FF6600;'>Spence</span></h1>", unsafe_allow_html=True)
            with st.form("form_login"):
                u = st.text_input("Usuario Corporativo").lower(); p = st.text_input("Contraseña", type="password")
                if st.form_submit_button("Acceder de forma segura", type="primary", use_container_width=True):
                    if u in USUARIOS and USUARIOS[u] == p: st.session_state.update({'logged_in': True, 'usuario_actual': u, 'informes_pendientes': cargar_pendientes(u)}); st.rerun()
                    else: st.error("❌ Credenciales inválidas.")
else:
    with st.sidebar:
        st.markdown("<h2 style='text-align: center; border-bottom:none; margin-top:-20px;'><span style='color:#007CA6;'>Atlas Copco</span> <span style='color:#FF6600;'>Spence</span></h2>", unsafe_allow_html=True)
        st.markdown(f"**Usuario:**<br>{st.session_state.usuario_actual.title()}", unsafe_allow_html=True); st.markdown("---")
        if st.button("🏭 Catálogo de Activos", use_container_width=True, type="primary" if st.session_state.vista_actual == "catalogo" else "secondary"): st.session_state.update({'vista_actual': "catalogo", 'vista_firmas': False, 'equipo_seleccionado': None}); st.rerun()
        if st.button("📊 Planificación", use_container_width=True, type="primary" if st.session_state.vista_actual == "planificacion" else "secondary"): st.session_state.update({'vista_actual': "planificacion", 'vista_firmas': False, 'equipo_seleccionado': None}); st.rerun()
        if st.button("📜 Últimas Intervenciones", use_container_width=True, type="primary" if st.session_state.vista_actual == "historial" else "secondary"): st.session_state.update({'vista_actual': "historial", 'vista_firmas': False, 'equipo_seleccionado': None}); st.rerun()
        if st.session_state.informes_pendientes:
            st.markdown("---"); st.warning(f"📝 {len(st.session_state.informes_pendientes)} reportes pendientes.")
            if st.button("✍️ Pizarra de Firmas", use_container_width=True, type="primary" if st.session_state.vista_actual == "firmas" else "secondary"): st.session_state.update({'vista_firmas': True, 'vista_actual': "firmas", 'equipo_seleccionado': None}); st.rerun()
        st.markdown("---"); 
        if st.button("🚪 Cerrar Sesión", use_container_width=True): st.session_state.logged_in = False; st.rerun()

    # --- 8. VISTAS ESPECÍFICAS ---
    if st.session_state.vista_actual == "historial":
        st.markdown("<div style='text-align:center; padding:20px;'><h1 style='color:#FF6600; font-size:3.5em; margin:0;'>Muro de Intervenciones</h1><p style='color:#8c9eb5; font-size:1.2em;'>Registro Histórico</p></div>", unsafe_allow_html=True)
        hist = obtener_historial_global()
        if not hist: st.info("Aún no hay reportes firmados y almacenados en la base de datos central.")
        else:
            h_u = []; v = set()
            for i in hist:
                d_obj = parse_fecha(i['fecha']); _id = (i['tag'], d_obj)
                if _id not in v: v.add(_id); i['fecha_obj'] = d_obj; h_u.append(i)
            h_agrupado = {}
            for i in h_u: h_agrupado.setdefault(i['fecha_obj'], []).append(i)
            for d_obj in sorted(list(h_agrupado.keys()), reverse=True):
                m_nom = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
                f_str = f"{d_obj.day} de {m_nom[d_obj.month]} de {d_obj.year}" if d_obj.year != 1970 else "Fecha Desconocida"
                st.markdown(f"<h3 style='color: white; border-bottom: 2px solid #2b3543; padding-bottom: 5px; margin-top: 15px;'>🗓️ {f_str}</h3>", unsafe_allow_html=True); cols = st.columns(3) 
                for idx, i in enumerate(h_agrupado[d_obj]):
                    bc = "#00e676" if i['estado'] == "Operativo" else "#ff1744"; bgc = "rgba(0, 230, 118, 0.1)" if i['estado'] == "Operativo" else "rgba(255, 23, 68, 0.1)"; ic = "✅" if i['estado'] == "Operativo" else "🚨"
                    c_s = str(i.get('condicion', '')).replace('\n', ' ').replace("'", '"'); r_s = str(i.get('reco', '')).replace('\n', ' ').replace("'", '"')
                    ch = f"<hr style='margin:8px 0; border-color:#2b3543;'><p style='margin:5px 0 0 0; color:#aeb9cc; font-size:0.8em;'>📝 <b>Condición Final:</b> {c_s}</p>" if c_s.strip() else ""
                    rh = f"<p style='margin:5px 0 0 0; color:#aeb9cc; font-size:0.8em;'>💡 <b>Nota:</b> {r_s}</p>" if r_s.strip() else ""
                    with cols[idx % 3]:
                        with st.container(border=True):
                            st.markdown(f"<div style='border-left: 5px solid {bc}; padding-left: 12px; height: 100%; display: flex; flex-direction: column; justify-content: space-between;'><div><div style='display: flex; justify-content: space-between; align-items: flex-start;'><h3 style='margin: 0; color: #007CA6;'>{i['tag']}</h3><span style='background: #2b3543; color: white; padding: 3px 8px; border-radius: 12px; font-size: 0.75em; font-weight: bold;'>🛠️ {i['tipo']}</span></div><p style='margin: 2px 0 10px 0; color: #aeb9cc; font-size: 0.9em;'>{i['modelo']} &bull; {i['area'].title()}</p><div style='background: #151a22; padding: 8px; border-radius: 8px; margin-bottom: 5px;'><p style='margin: 0; color: #8c9eb5; font-size: 0.85em;'>🧑‍🔧 <b>Técnico:</b> {i['tecnico']}</p>{ch}{rh}</div></div><div style='margin-top: 10px; background: {bgc}; border: 1px solid {bc}; color: {bc}; padding: 5px; border-radius: 6px; text-align: center; font-weight: bold; font-size: 0.85em;'>{ic} {i['estado']}</div></div>", unsafe_allow_html=True)

    elif st.session_state.vista_actual == "planificacion":
        df_cmms = cargar_cmms()
        sem_act = get_current_wk()
        df_cmms['S_Programada'] = df_cmms['S_Programada'].apply(formatear_wk)
        df_cmms['Mes_Calc'] = df_cmms['S_Programada'].apply(calcular_mes_minero)
        mes_hoy = calcular_mes_minero(sem_act)
        
        st.markdown(f"<div style='margin:1rem 0; background: linear-gradient(90deg, rgba(0,124,166,0.1) 0%, rgba(0,124,166,0.2) 50%, rgba(0,124,166,0.1) 100%); padding: 20px; border-radius: 15px; border-left: 5px solid var(--ac-blue);'><h2 style='color: white; margin: 0;'>📅 Panel de Control</h2><p style='color: #8c9eb5; margin: 0; font-weight: 600;'>Semana Actual: {sem_act} &nbsp;|&nbsp; Planificación Activa: {mes_hoy}</p></div>", unsafe_allow_html=True)
        
        d_k = df_cmms[(df_cmms["Mes_Calc"] == mes_hoy) & (df_cmms["Tipo"] != "N/A")]
        t_t, h, fs, pen = len(d_k), len(d_k[d_k["Estado"] == "✅ Hecho"]), len(d_k[d_k["Estado"] == "🚨 F/S"]), len(d_k[d_k["Estado"] == "⏳ Pendiente"])
        te = h + pen; cum = int((h / te * 100)) if te > 0 else (100 if h > 0 else 0)
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("📈 Cumplimiento Mes", f"{cum}%"); c2.metric("🎯 Tareas Programadas", t_t); c3.metric("✅ Tareas Completadas", h); c4.metric("🚨 Equipos F/S", fs)
        st.markdown("---")
        
        tab_g, tab_c, tab_m = st.tabs(["📋 Tablero", "📆 Calendario", "📊 Matriz de Mantenimiento"])
        
        with tab_g:
            c_f1, c_f2 = st.columns([1, 3])
            o_m = ["Todas", "Diciembre", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre"]
            f_m = c_f1.selectbox("Filtrar por Mes:", o_m, index=o_m.index(mes_hoy) if mes_hoy in o_m else 0)
            
            min_d, max_d = None, None
            if f_m != "Todas":
                m_m = {"Enero": 1, "Febrero": 2, "Marzo": 3, "Abril": 4, "Mayo": 5, "Junio": 6, "Julio": 7, "Agosto": 8, "Septiembre": 9, "Octubre": 10, "Noviembre": 11, "Diciembre": 12}
                if f_m in m_m:
                    num = m_m[f_m]; yr = 2025 if num == 12 else 2026
                    min_d = datetime.date(yr - 1, 12, 16) if num == 1 else datetime.date(yr, num - 1, 16)
                    max_d = datetime.date(yr, num, 15)

            df_m = df_cmms.copy() if f_m == "Todas" else df_cmms[df_cmms["Mes_Calc"] == f_m].copy()
            t_f = [t for t in list(inventario_equipos.keys()) if t not in df_m['TAG'].tolist()]
            if t_f: df_m = pd.concat([df_m, pd.DataFrame([{"TAG": t, "S_Programada": "", "Tipo": "N/A", "Estado": "⚪ N/A", "S_Realizada": None, "Observacion": "", "Mes_Calc": f_m if f_m != "Todas" else "Sin Asignar"} for t in t_f])], ignore_index=True)
            df_m = df_m.sort_values(by="TAG").reset_index(drop=True)

            if not df_m.empty:
                df_m['S_Realizada'] = df_m['S_Realizada'].apply(lambda v: datetime.datetime.strptime(str(v).strip(), "%Y-%m-%d").date() if v else None)
                df_m['Día Programado'] = df_m['S_Programada'].apply(wk_to_date); df_m.insert(0, "🗑️ Quitar", False)
                if "k_t" in st.session_state:
                    for i, c in st.session_state["k_t"].get("edited_rows", {}).items():
                        if "Día Programado" in c and c["Día Programado"]:
                            v = c["Día Programado"]; d = datetime.datetime.strptime(v[:10], "%Y-%m-%d").date() if isinstance(v, str) else v
                            df_m.at[int(i), 'S_Programada'] = f"WK{d.isocalendar()[1]:02d}"

                cc = {
                    "🗑️ Quitar": st.column_config.CheckboxColumn("Quitar", default=False), "TAG": st.column_config.TextColumn("Equipo", disabled=True), "Mes_Calc": None, "S_Programada": None, 
                    "Día Programado": st.column_config.DateColumn("📆 Prog. para (Día y WK)", format="DD/MM/YYYY - [WK]WW", min_value=min_d, max_value=max_d, disabled=False),
                    "Tipo": st.column_config.SelectboxColumn("Intervención", options=["N/A", "INSP", "P1", "P2", "P3", "P4", "PM03"], disabled=False),
                    "Estado": st.column_config.SelectboxColumn("Estado Actual", options=["⚪ N/A", "⏳ Pendiente", "✅ Hecho", "🚨 F/S"], required=True),
                    "S_Realizada": st.column_config.DateColumn("Día Ejecución (Día y WK) 📅", format="DD/MM/YYYY - [WK]WW", disabled=False), "Observacion": st.column_config.TextColumn("Comentarios")
                }
                def c_est(v):
                    if v == '✅ Hecho': return 'background-color: #063f22; color: #6ee7b7; font-weight: bold;'
                    if v == '⏳ Pendiente': return 'background-color: #423205; color: #fde047; font-weight: bold;'
                    if v == '🚨 F/S': return 'background-color: #471015; color: #ff8a93; font-weight: bold;'
                    if v == '⚪ N/A': return 'color: #556b82; font-style: italic;'
                    return ''
                df_st = df_m[["🗑️ Quitar", "TAG", "Día Programado", "Tipo", "Estado", "S_Realizada", "Observacion", "Mes_Calc", "S_Programada"]].style.map(c_est, subset=['Estado']) if hasattr(df_m.style, 'map') else df_m[["🗑️ Quitar", "TAG", "Día Programado", "Tipo", "Estado", "S_Realizada", "Observacion", "Mes_Calc", "S_Programada"]].style.applymap(c_est, subset=['Estado'])
                df_e = st.data_editor(df_st, key="k_t", hide_index=True, use_container_width=True, column_config=cc, height=750)
                
                if st.button("💾 Guardar Avances y Limpiar Tabla", type="primary"):
                    def get_fwk(r):
                        d = r['Día Programado']
                        if pd.notnull(d) and str(d).strip() not in ["", "None", "NaT"]: return f"WK{(datetime.datetime.strptime(d[:10], '%Y-%m-%d').date() if isinstance(d, str) else d).isocalendar()[1]:02d}"
                        return r['S_Programada']
                    df_e['S_Programada'], df_e['S_Realizada'] = df_e.apply(get_fwk, axis=1), df_e['S_Realizada'].apply(safe_date_str)
                    fv = df_e[(df_e["🗑️ Quitar"] == False) & (df_e["Tipo"] != "N/A") & (df_e["Estado"] != "⚪ N/A") & (df_e["S_Programada"] != "")].copy()
                    fv.loc[(fv['Estado'] == '✅ Hecho') & (fv['S_Realizada'] == ""), 'S_Realizada'] = datetime.date.today().strftime("%Y-%m-%d")
                    df_f = fv if f_m == "Todas" else pd.concat([df_cmms[df_cmms["Mes_Calc"] != f_m], fv], ignore_index=True)
                    for col in ['Mes_Calc', '🗑️ Quitar', 'Día Programado']:
                        if col in df_f.columns: df_f = df_f.drop(columns=[col])
                    guardar_cmms(df_f); st.success("✅ ¡Guardado!"); time.sleep(1.5); st.rerun()

            with st.expander("➕ Inyectar Tarea Extra", expanded=False):
                with st.form("form_nueva_tarea"):
                    c1, c2, c3 = st.columns(3); n_t = c1.selectbox("Equipo:", sorted(list(inventario_equipos.keys()))); n_tp = c2.selectbox("Tipo de Tarea:", ["INSP", "P1", "P2", "P3", "P4", "PM03"]); d_d = datetime.date.today()
                    if min_d and max_d and not (min_d <= d_d <= max_d): d_d = min_d
                    n_fp = c3.date_input("📆 Día a Programar:", value=d_d, min_value=min_d, max_value=max_d); n_o = st.text_input("Observación inicial (Opcional):")
                    if st.form_submit_button("🚀 Inyectar Tarea y Guardar Todo", type="primary", use_container_width=True):
                        df_g = df_cmms.copy()
                        if not df_e.empty:
                            dec = df_e.copy(); dec['S_Programada'] = dec.apply(get_fwk, axis=1); dec['S_Realizada'] = dec['S_Realizada'].apply(safe_date_str); fvf = dec[(dec["🗑️ Quitar"] == False) & (dec["Tipo"] != "N/A") & (dec["Estado"] != "⚪ N/A") & (dec["S_Programada"] != "")].copy(); fvf.loc[(fvf['Estado'] == '✅ Hecho') & (fvf['S_Realizada'] == ""), 'S_Realizada'] = datetime.date.today().strftime("%Y-%m-%d")
                            df_g = fvf if f_m == "Todas" else pd.concat([df_cmms[df_cmms["Mes_Calc"] != f_m], fvf], ignore_index=True)
                        nf = pd.DataFrame([{"TAG": n_t, "S_Programada": f"WK{n_fp.isocalendar()[1]:02d}", "Tipo": n_tp, "Estado": "⏳ Pendiente", "S_Realizada": "", "Observacion": n_o}])
                        for c in ['Mes_Calc', '🗑️ Quitar', 'Día Programado']:
                            if c in df_g.columns: df_g = df_g.drop(columns=[c])
                        guardar_cmms(pd.concat([df_g, nf], ignore_index=True)); st.success("✅ Guardado."); time.sleep(1.5); st.rerun()

        with tab_c:
            o_mc = ["Diciembre 2025"] + [f"{m} 2026" for m in ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]]; ct, cs = st.columns([2, 1])
            ct.markdown("### 📆 Calendario Interactivo")
            hc = datetime.date.today(); ms = f"Diciembre 2025" if hc.year == 2025 and hc.month == 12 else f"{o_mc[hc.month]}" if hc.year == 2026 else "Enero 2026"
            ms_v = cs.selectbox("📅 Mes a visualizar:", o_mc, index=o_mc.index(ms) if ms in o_mc else 1)
            yr_c = 2025 if "2025" in ms_v else 2026; m_c = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"].index(ms_v.split(" ")[0]) + 1
            sm = calendar.Calendar(calendar.MONDAY).monthdatescalendar(yr_c, m_c); t_f = {}
            for _, r in df_cmms.iterrows():
                dp = wk_to_date(r['S_Programada']); dt = None
                if r['Estado'] == '✅ Hecho' and str(r['S_Realizada']).strip() != "":
                    try: dt = datetime.datetime.strptime(str(r['S_Realizada']).strip(), "%Y-%m-%d").date()
                    except: dt = dp
                else: dt = dp
                if dt: t_f.setdefault(dt, []).append({"tag": r['TAG'], "tipo": r['Tipo'], "est": r['Estado']})
            hl = '<div style="display:grid; grid-template-columns: 65px repeat(7, 1fr); gap: 10px; margin-top:10px;"><div style="text-align:center; color:#FF6600; font-weight:900; font-size:0.8rem; margin-top: 10px;">REF</div>'
            for d in ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]: hl += f'<div style="text-align:center; color:#8c9eb5; font-weight:bold; font-size:0.9rem;">{d}</div>'
            for s in sm:
                wn = s[0].isocalendar()[1] if not (yr_c == 2025 and s[0].month == 12) else s[0].isocalendar()[1]
                hl += f'<div style="display:flex; align-items:center; justify-content:center; background:#2b3543; border-radius:8px; border-left: 4px solid #FF6600; color:white; font-weight:bold; font-size:0.85rem; height: 120px;">WK{wn:02d}</div>'
                for d in s:
                    bg = "#1a212b" if d.month == m_c else "#11151c"; br = "#00BFFF" if d == hc else "#2b3543"
                    hl += f'<div style="background:{bg}; border: 1px solid {br}; border-radius: 8px; padding: 5px; min-height: 120px;"><div style="text-align:right; color:white; font-size:0.9rem; margin-bottom:8px;">{d.day}</div>'
                    if d in t_f:
                        for t in t_f[d]:
                            cb, cx, bs = ("#063f22", "#6ee7b7", "1px solid #10b981") if t['est'] == '✅ Hecho' else ("#471015", "#ff8a93", "1px solid #ef4444") if t['est'] == '🚨 F/S' else ("#0c2d48", "#66c2ff", "1px solid #1a5c94") if t['tipo'] == 'P1' else ("#4a2c00", "#ffb04c", "1px solid #8c5300") if t['tipo'] == 'P2' else ("#301047", "#d78aff", "1px solid #622291") if t['tipo'] == 'P3' else ("#471015", "#ff8a93", "1px solid #8e202a") if t['tipo'] == 'P4' else ("transparent", "#8c9eb5", "1px dashed #455065")
                            hl += f'<div style="background:{cb}; color:{cx}; padding:4px; margin-bottom:4px; border-radius:4px; font-size:0.75rem; border: {bs};"><b>{t["tag"]}</b> - {t["tipo"]}</div>'
                    hl += '</div>'
            st.markdown(hl + '</div>', unsafe_allow_html=True)

        with tab_m:
            df_pb = df_cmms[df_cmms['Tipo'] != 'N/A'].copy(); df_pb['Contenido'] = df_pb['Tipo'] + "\n" + df_pb['Estado'].apply(lambda x: str(x).split(" ")[1] if " " in str(x) else str(x))
            cm1, cm2 = st.columns([1.5, 2]); vm = cm1.radio("Modo:", ["🔍 Por Mes (Zoom In)", "📆 Anual (Semanas WK)", "📅 Anual (Por Meses)"], horizontal=True)
            df_pb['Mes_Vista'] = df_pb['Mes_Calc'].apply(lambda q: "dic-25" if q == "Diciembre" else {"Enero":"ene-26", "Febrero":"feb-26", "Marzo":"mar-26", "Abril":"abr-26", "Mayo":"may-26", "Junio":"jun-26", "Julio":"jul-26", "Agosto":"ago-26", "Septiembre":"sept-26", "Octubre":"oct-26", "Noviembre":"nov-26"}.get(q, q))
            cp, ct = ('Mes_Vista', ["dic-25", "ene-26", "feb-26", "mar-26", "abr-26", "may-26", "jun-26", "jul-26", "ago-26", "sept-26", "oct-26", "nov-26"]) if vm == "📅 Anual (Por Meses)" else ('S_Programada', list(dict.fromkeys(["WK51", "WK52"] + [f"WK{i:02d}" for i in range(1, 53)])))
            df_pv = df_pb.groupby(['TAG', cp])['Contenido'].apply(lambda x: '\n---\n'.join(x)).unstack().fillna(""); li = [{"TAG": t, "Equipo": inventario_equipos[t][0] if t in inventario_equipos else "-", "Área": inventario_equipos[t][3].title() if t in inventario_equipos else "-"} for t in df_pv.index]
            df_mz = pd.concat([pd.DataFrame(li).set_index("TAG"), df_pv], axis=1).reset_index(); cb = ['TAG', 'Equipo', 'Área']
            for c in ct:
                if c not in df_mz.columns: df_mz[c] = ""
            df_mz = df_mz[cb + ct]; cf = cb.copy()
            if vm == "🔍 Por Mes (Zoom In)":
                wq = {wk: calcular_mes_minero(wk) for wk in ct}; omz = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]; qu = sorted(list(set(wq.values())), key=lambda x: omz.index(x) if x in omz else 99); qs = cm2.selectbox("Selecciona el Mes a enfocar:", qu, index=qu.index(mes_hoy) if mes_hoy in qu else 0); cf.extend([wk for wk, q in wq.items() if q == qs])
            else: cf.extend(ct)
            df_mc = df_mz[cf].set_index(['TAG', 'Equipo', 'Área'])
            def est_m(v):
                v = str(v).upper(); b = 'white-space: pre-wrap; line-height: 1.4; border-radius: 6px; padding: 6px; text-align: center; font-size: 0.85em; '
                if not v or v == "NAN": return ''
                if 'HECHO' in v: return b + 'background-color: #063f22; color: #6ee7b7; font-weight: bold; border-left: 4px solid #10b981;'
                if 'F/S' in v: return b + 'background-color: #471015; color: #ff8a93; font-weight: bold; border-left: 4px solid #ef4444;'
                if 'PENDIENTE' in v: return b + ('background-color: #0c2d48; color: #66c2ff;' if 'P1' in v else 'background-color: #4a2c00; color: #ffb04c;' if 'P2' in v else 'background-color: #301047; color: #d78aff;' if 'P3' in v else 'background-color: #471015; color: #ff8a93;' if 'P4' in v else 'background-color: #423205; color: #fde047;') + ' font-weight: bold; border-left: 4px solid #eab308;'
                return b + 'color: #8c9eb5; font-style: italic;'
            cpnt = [c for c in cf if c not in cb]
            st.dataframe(df_mc.style.map(est_m, subset=cpnt) if hasattr(df_mc.style, 'map') else df_mc.style.applymap(est_m, subset=cpnt), use_container_width=True, height=600) if cpnt else st.dataframe(df_mc, use_container_width=True, height=600)

    elif st.session_state.vista_actual == "catalogo" and not st.session_state.equipo_seleccionado:
        st.markdown("<div style='margin-bottom: 2.5rem; text-align: center;'><h1 style='color: #007CA6; font-size: 4em; font-weight: 800; margin: 0;'>Atlas Copco <span style='color: #FF6600;'>Spence</span></h1><p style='color: #8c9eb5; font-size: 1.2em; font-weight: 300; margin-top: -10px;'>Sistema Integrado de Control de Activos • Hidrometalurgia</p></div>", unsafe_allow_html=True)
        e_db = obtener_estados_actuales(); te = len(inventario_equipos); op = sum(1 for tag in inventario_equipos if e_db.get(tag, "Operativo") == "Operativo"); fs = te - op
        m1, m2, m3 = st.columns(3); m1.markdown(f"<div style='background: #1e2530; border-left: 5px solid #8c9eb5; padding: 20px; border-radius: 10px; text-align: center;'><p style='color: #8c9eb5; margin:0; font-weight:600;'>📦 Total Activos</p><h2 style='color: white; margin:0;'>{te}</h2></div>", unsafe_allow_html=True); m2.markdown(f"<div style='background: #1e2530; border-left: 5px solid #00e676; padding: 20px; border-radius: 10px; text-align: center;'><p style='color: #8c9eb5; margin:0; font-weight:600;'>🟢 Operativos</p><h2 style='color: #00e676; margin:0;'>{op}</h2></div>", unsafe_allow_html=True); m3.markdown(f"<div style='background: #1e2530; border-left: 5px solid #ff1744; padding: 20px; border-radius: 10px; text-align: center;'><p style='color: #8c9eb5; margin:0; font-weight:600;'>🔴 Fuera de Servicio</p><h2 style='color: #ff1744; margin:0;'>{fs}</h2></div>", unsafe_allow_html=True); st.markdown("<hr style='border-color: #2b3543;'>", unsafe_allow_html=True)
        cf, cb = st.columns([1.2, 2]); ft = cf.radio("🗂️ Categoría:", ["Todos", "Compresores", "Secadores"], horizontal=True); bsq = cb.text_input("🔍 Buscar...").lower(); cols = st.columns(4); c = 0
        for tag, (mod, ser, area, ubi) in inventario_equipos.items():
            if (ft == "Compresores" and "CD" in mod.upper()) or (ft == "Secadores" and "CD" not in mod.upper()): continue
            if bsq in tag.lower() or bsq in area.lower() or bsq in mod.lower() or bsq in ubi.lower():
                est = e_db.get(tag, "Operativo"); c_b = "#00e676" if est == "Operativo" else "#ff1744"; b_h = f"<div style='background: rgba({('0,230,118' if est == 'Operativo' else '255,23,68')},0.15); color: {c_b}; border: 1px solid {c_b}; padding: 4px 10px; border-radius: 20px; font-size: 0.75rem; font-weight: 700;'>{est.upper()}</div>"
                with cols[c % 4]:
                    with st.container(border=True):
                        st.markdown(f"<div style='border-top: 4px solid {c_b}; padding-top: 10px; text-align: center; margin-top:-10px;'>{b_h}</div>", unsafe_allow_html=True); st.button(tag, key=f"b_{tag}", on_click=seleccionar_equipo, args=(tag,), use_container_width=True); st.markdown(f"<p style='color: #8c9eb5; margin-top: 5px; font-size: 0.85rem; text-align: center;'><strong style='color:#007CA6;'>{mod}</strong> &bull; {area.title()}<br><small style='color: #556b82;'>{ubi.title()}</small></p>", unsafe_allow_html=True)
                c += 1

    elif st.session_state.equipo_seleccionado:
        t = st.session_state.equipo_seleccionado; mod, ser, area, ubi = inventario_equipos[t]; cb, ct = st.columns([1, 4]); cb.button("⬅️ Volver", on_click=volver_catalogo, use_container_width=True); ct.markdown(f"<h1 style='margin-top:-15px;'>⚙️ Ficha: <span style='color:#007CA6;'>{t}</span></h1>", unsafe_allow_html=True); st.markdown("<br>", unsafe_allow_html=True); SP = obtener_especificaciones(DEFAULT_SPECS)
        t1, t2, t3, t4 = st.tabs(["📋 Reporte", "📚 Ficha", "🔍 Bitácora", "👤 Contactos"])
        with t1:
            tp = st.selectbox("🛠️ Orden:", ["Inspección", "PM03"] if "CD" in t else ["Inspección", "P1", "P2", "P3", "PM03"]); c1, c2, c3, c4 = st.columns(4); c1.text_input("Modelo", mod, disabled=True); c2.text_input("Serie", ser, disabled=True); c3.text_input("Área", area, disabled=True); c4.text_input("Macro", ubi, disabled=True); c5, c6, c7, c8 = st.columns([1, 1, 1, 1.3]); fec = c5.text_input("Fecha", obtener_fecha_hoy_esp()); t1 = c6.text_input("Técnico 1", key="input_tec1"); t2 = c7.text_input("Técnico 2", key="input_tec2")
            with c8:
                cdb = obtener_contactos(); op = ["➕ Nuevo..."] + cdb; ci = op.index(st.session_state.input_cliente) if st.session_state.input_cliente in op else 1 if cdb else 0
                sc1, sc2 = st.columns([4, 1]); csel = sc1.selectbox("Cliente", op, index=ci)
                if csel != "➕ Nuevo..." and sc2.button("❌"): eliminar_contacto(csel); st.rerun()
                if csel == "➕ Nuevo...":
                    nc = st.text_input("Nombre:")
                    if st.button("💾 Guardar"): agregar_contacto(nc); st.session_state.input_cliente = nc.strip().title(); st.rerun()
                    cc = nc.strip().title()
                else: cc = csel; st.session_state.input_cliente = csel
            st.markdown("<hr>", unsafe_allow_html=True); c9, c10, c11, c12, c13, c14 = st.columns(6); hm = c9.number_input("Horas Marcha", step=1, value=int(st.session_state.input_h_marcha)); hc = c10.number_input("Horas Carga", step=1, value=int(st.session_state.input_h_carga)); up = c11.selectbox("Unidad", ["Bar", "psi"]); pc = c12.text_input("P. Carga", value=str(st.session_state.input_p_carga)); pd_d = c13.text_input("P. Descarga", value=str(st.session_state.input_p_descarga)); ts = c14.text_input("Temp Salida (°C)", value=str(st.session_state.input_temp)); pcc, pdc, tsc = pc.replace(',', '.'), pd_d.replace(',', '.'), ts.replace(',', '.')
            st.markdown("<hr>", unsafe_allow_html=True); ee = st.radio("Devolución:", ["Operativo", "Fuera de servicio"], key="input_estado_eq", horizontal=True); ent = st.text_area("Condición Final:", key="input_estado"); rec = st.text_area("Acciones Pendientes:", key="input_reco"); st.markdown("<br>", unsafe_allow_html=True)
            if st.button("📥 Guardar en Bandeja", type="primary", use_container_width=True):
                actualizar_estado_equipo_en_nube(t, ee); pl = "plantilla/secadorfueradeservicio.docx" if "CD" in t and ee == "Fuera de servicio" else "plantilla/inspeccionsecador.docx" if "CD" in t else "plantilla/fueradeservicio.docx" if ee == "Fuera de servicio" else f"plantilla/{tp.lower()}.docx" if tp in ["P1", "P2", "P3"] else "plantilla/inspeccion.docx"
                ctx = {"tipo_intervencion": tp, "modelo": mod, "tag": t, "area": area, "ubicacion": ubi, "cliente_contacto": cc, "p_carga": f"{pcc} {up}", "p_descarga": f"{pdc} {up}", "temp_salida": tsc, "horas_marcha": int(hm), "horas_carga": int(hc), "tecnico_1": t1, "tecnico_2": t2, "estado_equipo": ee, "estado_entrega": ent, "recomendaciones": rec, "serie": ser, "tipo_orden": tp.upper(), "fecha": fec}; na = f"Informe_{tp}_{t}_{fec.replace(' ','_')}.docx"; rd = os.path.join(RUTA_ONEDRIVE, na)
                with st.spinner("Creando borrador..."): docp = DocxTemplate(pl); docp.render({**ctx, 'firma_tecnico': "", 'firma_cliente': ""}); os.makedirs(RUTA_ONEDRIVE, exist_ok=True); rwp = os.path.join(RUTA_ONEDRIVE, f"PREVIEW_{na}"); docp.save(rwp); rpp = convertir_a_pdf(rwp)
                st.session_state.informes_pendientes.append({"tag": t, "file_plantilla": pl, "context": ctx, "tupla_db": (t, mod, ser, area, ubi, fec, cc, t1, t2, float(tsc) if tsc.replace('.','',1).isdigit() else 0.0, f"{pcc} {up}", f"{pdc} {up}", hm, hc, ent, tp, rec, ee, "", st.session_state.usuario_actual), "ruta_docx": rd, "nombre_archivo_base": na, "ruta_prev_pdf": rpp}); guardar_pendientes(st.session_state.usuario_actual, st.session_state.informes_pendientes); volver_catalogo(); st.rerun()
        with t2:
            with st.expander("✏️ Agregar Datos"):
                with st.form(f"fs_{t}"):
                    ce1, ce2 = st.columns(2); cls = ce1.selectbox("Dato:", ["Litros Aceite", "Tipo Aceite", "Cant. Filtro Aceite", "Cant. Filtro Aire", "N° Parte Kit", "N° Parte Separador", "Otro..."]); clf = ce1.text_input("Nombre:") if cls == "Otro..." else cls; vf = ce2.text_input("Valor:")
                    if st.form_submit_button("💾 Guardar"): guardar_especificacion_db(mod, clf.strip(), vf.strip()); st.rerun()
            if mod in SP:
                sp = {k: v for k, v in SP[mod].items() if k != "Manual"}; cs = st.columns(3)
                for i, (k, v) in enumerate(sp.items()): cs[i % 3].markdown(f"<div style='background-color: #1e2530; padding: 15px; border-radius: 8px; margin-bottom: 15px; border-left: 4px solid #007CA6;'><small style='color: #8c9eb5;'>{k.upper()}</small><br><span style='color: white;'>{v}</span></div>", unsafe_allow_html=True)
                if "Manual" in SP[mod] and os.path.exists(SP[mod]["Manual"]):
                    with open(SP[mod]["Manual"], "rb") as f: st.download_button(f"📕 Manual {mod} (PDF)", f, file_name=f"Manual_{mod}.pdf", mime="application/pdf")
        with t3:
            with st.form(f"fo_{t}"):
                no = st.text_area("Nueva Observación:"); 
                if st.form_submit_button("➕ Registrar"): agregar_observacion(t, st.session_state.usuario_actual, no); st.rerun()
            dfo = obtener_observaciones(t);
            if not dfo.empty:
                for _, r in dfo.iterrows():
                    co, cd = st.columns([11, 1]); co.markdown(f"<div style='background-color: #2b303b; padding: 15px; border-radius: 8px; margin-bottom: 10px; border-left: 4px solid #FF6600;'><small style='color: #aeb9cc;'>👤 {r['usuario']} | 📅 {r['fecha']}</small><br><span style='color: white;'>{r['texto']}</span></div>", unsafe_allow_html=True)
                    if cd.button("🗑️", key=f"do_{r['id']}"): eliminar_observacion(r['id']); st.rerun()
        with t4:
            with st.expander("✏️ Editar Contacto/Seguridad"):
                with st.form(f"fa_{t}"):
                    ca1, ca2 = st.columns(2); csa = ca1.selectbox("Dato:", ["Dueño Turno 1-3", "Dueño Turno 2-4", "PEA", "Frecuencia Radial", "Supervisor", "Jefe Turno", "Otro..."]); cfa = ca1.text_input("Cargo:") if csa == "Otro..." else csa; vfa = ca2.text_input("Info:")
                    if st.form_submit_button("💾 Guardar"): guardar_dato_equipo(t, cfa.strip(), vfa.strip()); st.rerun()
            de = obtener_datos_equipo(t); ca = st.columns(2)
            for i, (k, v) in enumerate(de.items()): ca[i % 2].markdown(f"<div style='background-color: #2b303b; padding: 15px; border-radius: 8px; margin-bottom: 15px; border-left: 4px solid #FF6600;'><small style='color: #aeb9cc;'>{k.upper()}</small><br><span style='color: white;'>{v}</span></div>", unsafe_allow_html=True)
        st.markdown("<br><hr>", unsafe_allow_html=True); st.markdown("### 📋 Trazabilidad Histórica"); dfh = obtener_todo_el_historial(t)
        if not dfh.empty: st.dataframe(dfh, use_container_width=True)

    elif st.session_state.vista_firmas or st.session_state.vista_actual == "firmas":
        cv1, cv2 = st.columns([1,4]); cv1.button("⬅️ Volver", on_click=volver_catalogo, use_container_width=True); cv2.markdown("<h1 style='margin-top:-15px;'>✍️ Pizarra de Firmas y Revisión</h1>", unsafe_allow_html=True); st.markdown("---")
        if not st.session_state.informes_pendientes: st.info("🎉 No tienes informes pendientes.")
        else:
            with st.expander("🧑‍🔧 Firma de Técnico", expanded=(not st.session_state.firma_tec_json)):
                ctg = st_canvas(stroke_width=4, stroke_color="#000", background_color="#fff", height=180, width=400, drawing_mode="freedraw", key="ctg", initial_drawing=st.session_state.firma_tec_json if st.session_state.firma_tec_json else None); cb1, cb2 = st.columns(2)
                if cb1.button("💾 Guardar Mi Firma", use_container_width=True):
                    if ctg.json_data is not None and len(ctg.json_data.get("objects", [])) > 0: st.session_state.update({'firma_tec_json': ctg.json_data, 'firma_tec_img': ctg.image_data}); st.success("✅ Guardada."); time.sleep(1); st.rerun()
                    else: st.warning("⚠️ Dibuja tu firma.")
                if cb2.button("🔄 Reiniciar Firma", use_container_width=True): st.session_state.update({'firma_tec_json': None, 'firma_tec_img': None}); st.rerun()
            st.markdown("<br>", unsafe_allow_html=True); ag = {}
            for inf in st.session_state.informes_pendientes:
                ar = inventario_equipos.get(inf['tag'], ["", "", "", "General"])[3].title(); ag.setdefault(ar, []).append(inf)
            for ma, ia in ag.items():
                st.markdown(f"### 🏢 {ma} ({len(ia)} pendientes)")
                with st.container(border=True):
                    for idx, inf in enumerate(ia):
                        ce, cdel = st.columns([11, 1])
                        with ce:
                            with st.expander(f"📝 {inf['tag']} ({inf['tipo_plan']})"):
                                tv, te = st.tabs(["📄 Borrador", "✏️ Corregir"])
                                with tv:
                                    if inf.get('ruta_prev_pdf') and os.path.exists(inf['ruta_prev_pdf']):
                                        cd1, cd2 = st.columns(2)
                                        with cd1:
                                            with open(inf['ruta_prev_pdf'], "rb") as f: st.download_button("⬇️ Borrador (PDF)", f, file_name=f"Borr_{inf['nombre_archivo_base'].replace('.docx', '.pdf')}", mime="application/pdf", key=f"dpdf_{inf['tag']}_{idx}")
                                        with cd2:
                                            rw = os.path.join(RUTA_ONEDRIVE, f"PREVIEW_{inf['nombre_archivo_base']}")
                                            if os.path.exists(rw):
                                                with open(rw, "rb") as f: st.download_button("⬇️ Borrador (Word)", f, file_name=f"Borr_{inf['nombre_archivo_base']}", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"dw_{inf['tag']}_{idx}")
                                        st.markdown("<br>", unsafe_allow_html=True); pdf_viewer(inf['ruta_prev_pdf'], width=950, height=900)
                                    else: st.warning("⚠️ No disponible.")
                                with te:
                                    with st.form(f"ef_{inf['tag']}_{idx}"):
                                        c1, c2, c3 = st.columns(3); nhm = c1.number_input("Horas Marcha", value=int(inf['context'].get('horas_marcha', 0))); nhc = c2.number_input("Horas Carga", value=int(inf['context'].get('horas_carga', 0))); nts = c3.text_input("Temp Salida", value=str(inf['context'].get('temp_salida', '0')))
                                        c4, c5 = st.columns(2); npc = c4.text_input("P. Carga", value=str(inf['context'].get('p_carga', ''))); npd = c5.text_input("P. Descarga", value=str(inf['context'].get('p_descarga', '')))
                                        nee, nre = st.text_area("Condición Final", value=str(inf['context'].get('estado_entrega', ''))), st.text_area("Recomendaciones", value=str(inf['context'].get('recomendaciones', '')))
                                        if st.form_submit_button("💾 Guardar y Regenerar", type="primary"):
                                            inf['context'].update({'horas_marcha': nhm, 'horas_carga': nhc, 'temp_salida': nts, 'p_carga': npc, 'p_descarga': npd, 'estado_entrega': nee, 'recomendaciones': nre})
                                            tl = list(inf['tupla_db']); 
                                            try: tl[9] = float(nts.replace(',', '.'))
                                            except: tl[9] = 0.0
                                            tl[10], tl[11], tl[12], tl[13], tl[14], tl[16] = npc, npd, nhm, nhc, nee, nre; inf['tupla_db'] = tuple(tl)
                                            dp = DocxTemplate(inf['file_plantilla']); ctxp = inf['context'].copy(); ctxp.update({'firma_tecnico': "", 'firma_cliente': ""}); dp.render(ctxp); rwp = os.path.join(RUTA_ONEDRIVE, f"PREVIEW_{inf['nombre_archivo_base']}"); dp.save(rwp); inf['ruta_prev_pdf'] = convertir_a_pdf(rwp); guardar_pendientes(st.session_state.usuario_actual, st.session_state.informes_pendientes); st.success("✅ Actualizado."); time.sleep(1); st.rerun()
                        if cdel.button("❌", key=f"del_{inf['tag']}_{idx}"): st.session_state.informes_pendientes.remove(inf); guardar_pendientes(st.session_state.usuario_actual, st.session_state.informes_pendientes); volver_catalogo() if not st.session_state.informes_pendientes else st.rerun()
                    st.markdown("---"); nc = " y ".join(list(set([i['cli'] for i in ia if i.get('cli')]))); st.markdown(f"<h3 style='text-align: center; color: #007CA6;'>Firma de Aprobación</h3><p style='text-align: center; color: #8c9eb5; margin-top: -10px;'><b>{nc if nc else 'Cliente'}</b></p>", unsafe_allow_html=True); _, cf, _ = st.columns([1, 2, 1])
                    with cf: 
                        with st.container(border=True): c_cli = st_canvas(stroke_width=4, stroke_color="#000", background_color="#f0f2f6", height=200, width=450, drawing_mode="freedraw", key=f"cc_{ma}"); st.markdown("<p style='text-align: center; font-size: 0.8em; color: gray;'>Firme arriba</p>", unsafe_allow_html=True)
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button(f"🚀 Aprobar y Enviar", type="primary", use_container_width=True, key=f"bs_{ma}"):
                        if not st.session_state.firma_tec_img is not None: st.warning("⚠️ Guarda tu firma técnico.")
                        elif not (c_cli.image_data is not None and c_cli.json_data is not None and len(c_cli.json_data.get("objects", [])) > 0): st.warning("⚠️ Falta firma cliente.")
                        else:
                            def pr_img(img): i = Image.fromarray(img.astype('uint8'), 'RGBA'); io_i = io.BytesIO(); i.save(io_i, format='PNG'); io_i.seek(0); return io_i
                            ifl = []
                            with st.spinner("Sellando..."):
                                try:
                                    for inf in ia:
                                        doc = DocxTemplate(inf['file_plantilla']); inf['context'].update({'firma_tecnico': InlineImage(doc, pr_img(st.session_state.firma_tec_img), width=Mm(40)), 'firma_cliente': InlineImage(doc, pr_img(c_cli.image_data), width=Mm(40))}); doc.render(inf['context']); doc.save(inf['ruta_docx']); rp = convertir_a_pdf(inf['ruta_docx']); rf, nf = (rp, inf['nombre_archivo_base'].replace(".docx", ".pdf")) if rp else (inf['ruta_docx'], inf['nombre_archivo_base']); tl = list(inf['tupla_db']); tl[18] = rf; guardar_registro(tuple(tl)); ifl.append({"tag": inf['tag'], "tipo": inf['tipo_plan'], "ruta": rf, "nombre_archivo": f"{ma}@@{inf['tag']}@@{nf}"})
                                    ok, mm = enviar_carrito_por_correo(MI_CORREO_CORPORATIVO, ifl)
                                    if ok: 
                                        st.success("✅ ¡Enviados!"); 
                                        for i in ia: st.session_state.informes_pendientes.remove(i)
                                        guardar_pendientes(st.session_state.usuario_actual, st.session_state.informes_pendientes); time.sleep(2); volver_catalogo() if not st.session_state.informes_pendientes else st.rerun()
                                    else: st.error(f"Error mail: {mm}")
                                except Exception as e: st.error(f"Error PDF: {e}")