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
# 0.2 ESTILOS PREMIUM (BOTÓN LATERAL BLINDADO)
# =============================================================================
st.set_page_config(page_title="Atlas Spence | Gestión de Reportes", layout="wide", page_icon="⚙️", initial_sidebar_state="expanded")

def aplicar_estilos_premium():
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;600;800&display=swap');
        :root { --ac-blue: #007CA6; --ac-dark: #005675; --bhp-orange: #FF6600; }
        html, body, p, h1, h2, h3, h4, h5, h6, span, div { font-family: 'Montserrat', sans-serif; }
        
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
    "Taller": ["GA 18", "API335343", "Taller", "Taller Central"]
}

# =============================================================================
# 2. CONEXIÓN A GOOGLE SHEETS (HÍBRIDA RENDER / LOCAL)
# =============================================================================
@st.cache_resource(show_spinner=False)
def get_gspread_client():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    try:
        creds_dict = json.loads(os.environ["gcp_json"])
    except:
        creds_dict = json.loads(st.secrets["gcp_json"])
    
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return gspread.authorize(creds)

def get_sheet(sheet_name):
    try:
        client = get_gspread_client()
        doc = client.open("BaseDatos")
        try: return doc.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound: 
            # MAGIA: Si no existe, la crea automáticamente con espacio de sobra.
            return doc.add_worksheet(title=sheet_name, rows="1000", cols="20")
    except Exception as e: return None

# =============================================================================
# 3. FUNCIONES DE BASE DE DATOS (FUSIÓN PERFECTA DE HOJAS)
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
            registros = sheet.get_all_values()
            fila_encontrada = -1
            for i, fila in enumerate(registros):
                if len(fila) > 0 and fila[0] == tag:
                    fila_encontrada = i + 1
                    break
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
# 4. FUNCIONES AUXILIARES
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

# =============================================================================
# 5. INICIALIZACIÓN DE ESTADOS Y MOTOR CMMS (NUEVO)
# =============================================================================
@st.cache_data(ttl=60, show_spinner=False)
def cargar_cmms():
    headers = ["TAG", "S_Programada", "Tipo", "Estado", "S_Realizada", "Observacion"]
    try:
        sheet = get_sheet("plan_cmms")
        if sheet:
            data = sheet.get_all_values()
            
            if len(data) > 0:
                # Convierte los datos a DataFrame
                df = pd.DataFrame(data[1:], columns=data[0]) if len(data) > 1 else pd.DataFrame(columns=data[0])
                
                # VERIFICACIÓN CRÍTICA: ¿Están bien escritos los títulos?
                if "S_Programada" in df.columns:
                    return df
                else:
                    # Si alguien borró los títulos o la hoja tiene basura, la resetea y la arregla
                    sheet.clear()
                    sheet.append_row(headers)
                    st.cache_data.clear()
                    return pd.DataFrame(columns=headers)
            else:
                # Si la hoja está 100% vacía
                sheet.append_row(headers)
                st.cache_data.clear()
                return pd.DataFrame(columns=headers)
    except Exception as e:
        print(f"Error cargando CMMS: {e}")
        
    # Si Google falla, devuelve un dataframe vacío pero con la estructura correcta para que no explote
    return pd.DataFrame(columns=headers)

def guardar_cmms(df):
    sheet = get_sheet("plan_cmms")
    if sheet:
        sheet.clear()
        sheet.append_rows([df.columns.values.tolist()] + df.values.tolist())
        st.cache_data.clear()

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
            
        if st.button("📊 Panel de Planificación (CMMS)", use_container_width=True, type="primary" if st.session_state.vista_actual == "planificacion" else "secondary"):
            st.session_state.vista_actual = "planificacion"; st.session_state.vista_firmas = False; st.session_state.equipo_seleccionado = None; st.rerun()
            
        if len(st.session_state.informes_pendientes) > 0:
            st.markdown("---")
            st.warning(f"📝 Tienes {len(st.session_state.informes_pendientes)} reportes esperando firmas.")
            if st.button("✍️ Ir a Pizarra de Firmas", use_container_width=True, type="primary" if st.session_state.vista_actual == "firmas" else "secondary"): 
                st.session_state.vista_firmas = True; st.session_state.vista_actual = "firmas"; st.session_state.equipo_seleccionado = None; st.rerun()
        st.markdown("---")
        if st.button("🚪 Cerrar Sesión", use_container_width=True): st.session_state.logged_in = False; st.rerun()

    # --- 7.1 VISTA PLANIFICACIÓN (NUEVO CMMS KANBAN) ---
    if st.session_state.vista_actual == "planificacion":
        df_cmms = cargar_cmms()
        hoy = datetime.date.today()
        semana_actual = f"WK{hoy.isocalendar()[1]:02d}"
        
        st.markdown(f"""
            <div style="margin-top: 1rem; margin-bottom: 1rem; background: linear-gradient(90deg, rgba(0,124,166,0.1) 0%, rgba(0,124,166,0.2) 50%, rgba(0,124,166,0.1) 100%); padding: 20px; border-radius: 15px; border-left: 5px solid var(--ac-blue);">
                <h2 style="color: white; margin: 0;">📅 Panel de Control CMMS</h2>
                <p style="color: #8c9eb5; margin: 0; font-weight: 600;">Semana en curso: {semana_actual}</p>
            </div>
        """, unsafe_allow_html=True)
        
        # --- PANEL DE INDICADORES (KPIs) ---
        df_semana = df_cmms[df_cmms["S_Programada"] == semana_actual]
        total_tareas = len(df_semana)
        hechas = len(df_semana[df_semana["Estado"] == "Hecho"])
        fs = len(df_semana[df_semana["Estado"] == "F/S"])
        cumplimiento = int((hechas / total_tareas * 100)) if total_tareas > 0 else 100
        
        c_kpi1, c_kpi2, c_kpi3, c_kpi4 = st.columns(4)
        c_kpi1.metric(label="📈 Cumplimiento Semanal", value=f"{cumplimiento}%")
        c_kpi2.metric(label="🎯 Tareas Programadas", value=total_tareas)
        c_kpi3.metric(label="✅ Tareas Completadas", value=hechas)
        c_kpi4.metric(label="🚨 Equipos F/S", value=fs)
        
        st.markdown("---")
        tab_gestion, tab_crear = st.tabs(["📋 Tablero Kanban de Mantenimiento", "➕ Programar Nueva Tarea"])
        
        # --- TABLERO INTERACTIVO ---
        with tab_gestion:
            st.info("Haz doble clic en la columna **Estado Actual** para marcar un equipo como Hecho o Fuera de Servicio. Luego dale a Guardar.")
            
            c_f1, c_f2 = st.columns([1, 3])
            with c_f1: 
                opciones_sem = ["Todas"] + sorted(df_cmms["S_Programada"].unique().tolist())
                filtro_sem = st.selectbox("Filtrar por Semana:", opciones_sem, index=opciones_sem.index(semana_actual) if semana_actual in opciones_sem else 0)
            
            df_mostrar = df_cmms.copy() if filtro_sem == "Todas" else df_cmms[df_cmms["S_Programada"] == filtro_sem]
            
            if not df_mostrar.empty:
                config_columnas = {
                    "TAG": st.column_config.TextColumn("Equipo", disabled=True),
                    "S_Programada": st.column_config.TextColumn("Semana Prog.", disabled=True),
                    "Tipo": st.column_config.TextColumn("Intervención", disabled=True),
                    "Estado": st.column_config.SelectboxColumn("Estado Actual", options=["Pendiente", "Hecho", "F/S"], required=True),
                    "S_Realizada": st.column_config.TextColumn("Semana Realizada (Ej: WK10)"),
                    "Observacion": st.column_config.TextColumn("Comentarios / Desviaciones")
                }
                
                def color_estado(val):
                    if val == 'Hecho': return 'background-color: #063f22; color: #6ee7b7; font-weight: bold;'
                    if val == 'Pendiente': return 'background-color: #423205; color: #fde047; font-weight: bold;'
                    if val == 'F/S': return 'background-color: #471015; color: #ff8a93; font-weight: bold;'
                    return ''
                    
                try: df_estilizado = df_mostrar.style.map(color_estado, subset=['Estado'])
                except AttributeError: df_estilizado = df_mostrar.style.applymap(color_estado, subset=['Estado'])
                
                df_editado = st.data_editor(df_estilizado, hide_index=True, use_container_width=True, column_config=config_columnas, height=400)
                
                if st.button("💾 Guardar Avances en la Nube", type="primary"):
                    # Compara si algo cambió y actualiza la matriz original
                    df_cmms.update(df_editado)
                    # Si algo se marcó como "Hecho" y no tenía semana realizada, se le pone la de hoy
                    df_cmms.loc[(df_cmms['Estado'] == 'Hecho') & (df_cmms['S_Realizada'] == ""), 'S_Realizada'] = semana_actual
                    guardar_cmms(df_cmms)
                    st.success("✅ Tablero sincronizado perfectamente.")
                    time.sleep(1.5); st.rerun()
            else:
                st.success(f"🎉 No hay tareas programadas para la semana {filtro_sem}. Ve a la siguiente pestaña para crear una.")

        # --- MOTOR DE PROGRAMACIÓN ---
        with tab_crear:
            with st.form("form_nueva_tarea"):
                st.markdown("### 📅 Programar Nueva Intervención")
                c1, c2, c3 = st.columns(3)
                n_tag = c1.selectbox("Equipo:", sorted(list(inventario_equipos.keys())))
                n_tipo = c2.selectbox("Tipo de Tarea:", ["INSP", "P1", "P2", "P3", "P4", "PM03"])
                n_sem = c3.text_input("Semana Programada (Ej: WK12):", value=semana_actual)
                
                n_obs = st.text_input("Observación inicial (Opcional):", placeholder="Ej: Se requiere cambio de aceite especial...")
                
                if st.form_submit_button("🚀 Añadir a la Planificación", type="primary", use_container_width=True):
                    nueva_fila = pd.DataFrame([{"TAG": n_tag, "S_Programada": n_sem.upper(), "Tipo": n_tipo, "Estado": "Pendiente", "S_Realizada": "", "Observacion": n_obs}])
                    df_cmms = pd.concat([df_cmms, nueva_fila], ignore_index=True)
                    guardar_cmms(df_cmms)
                    st.success(f"✅ ¡Listo! Tarea de {n_tipo} para {n_tag} programada exitosamente.")
                    time.sleep(1.5); st.rerun()

    # --- 7.2 VISTA DE FIRMAS (AGRUPADA POR MACRO-ÁREA) ---
    elif st.session_state.vista_firmas or st.session_state.vista_actual == "firmas":
        c_v1, c_v2 = st.columns([1,4])
        with c_v1: 
            if st.button("⬅️ Volver", use_container_width=True): volver_catalogo(); st.rerun()
        with c_v2: st.markdown("<h1 style='margin-top:-15px;'>✍️ Pizarra de Firmas por Área</h1>", unsafe_allow_html=True)
        st.markdown("---")
        
        if len(st.session_state.informes_pendientes) == 0:
            st.info("🎉 ¡Excelente! No tienes ningún informe pendiente por firmar.")
        else:
            areas_agrupadas = {}
            for inf in st.session_state.informes_pendientes:
                macro_area = inventario_equipos[inf['tag']][3].title() if inf['tag'] in inventario_equipos else "General"
                if macro_area not in areas_agrupadas: areas_agrupadas[macro_area] = []
                areas_agrupadas[macro_area].append(inf)

            for macro_area, informes_area in areas_agrupadas.items():
                st.markdown(f"### 🏢 Informes de {macro_area} ({len(informes_area)} pendientes)")
                with st.container(border=True):
                    for inf in informes_area:
                        c_exp, c_del = st.columns([12, 1])
                        with c_exp:
                            with st.expander(f"📄 Ver documento preliminar: {inf['tag']} ({inf['tipo_plan']} - {inf['area'].title()})"):
                                if inf.get('ruta_prev_pdf') and os.path.exists(inf['ruta_prev_pdf']):
                                    try: pdf_viewer(inf['ruta_prev_pdf'], width=700, height=600)
                                    except Exception as e: st.error(f"Error visor: {e}")
                                else: st.warning("⚠️ Vista preliminar no disponible.")
                        with c_del:
                            st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)
                            if st.button("❌", key=f"del_{inf['tag']}_{inf['tupla_db'][5].replace('/','')}_{inf['tipo_plan']}", help="Quitar este informe"):
                                st.session_state.informes_pendientes.remove(inf)
                                guardar_pendientes(st.session_state.usuario_actual, st.session_state.informes_pendientes) 
                                if len(st.session_state.informes_pendientes) == 0: volver_catalogo()
                                st.rerun()
                    
                    st.markdown("---")
                    c_tec, c_cli = st.columns(2)
                    with c_tec:
                        st.markdown(f"#### 🧑‍🔧 Firma Técnico ({macro_area})")
                        canvas_tec = st_canvas(stroke_width=4, stroke_color="#000", background_color="#fff", height=180, width=400, drawing_mode="freedraw", key=f"tec_{macro_area}")
                    with c_cli:
                        st.markdown(f"#### 👷 Firma Cliente ({macro_area})")
                        canvas_cli = st_canvas(stroke_width=4, stroke_color="#000", background_color="#fff", height=180, width=400, drawing_mode="freedraw", key=f"cli_{macro_area}")
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button(f"🚀 Aprobar, Firmar y Subir Informes de {macro_area}", type="primary", use_container_width=True, key=f"btn_subir_{macro_area}"):
                        tec_ok = canvas_tec.image_data is not None and canvas_tec.json_data is not None and len(canvas_tec.json_data.get("objects", [])) > 0
                        cli_ok = canvas_cli.image_data is not None and canvas_cli.json_data is not None and len(canvas_cli.json_data.get("objects", [])) > 0
                        
                        if tec_ok and cli_ok:
                            def procesar_imagen_firma(img_data):
                                img = Image.fromarray(img_data.astype('uint8'), 'RGBA'); img_io = io.BytesIO(); img.save(img_io, format='PNG'); img_io.seek(0); return img_io
                            io_tec = procesar_imagen_firma(canvas_tec.image_data); io_cli = procesar_imagen_firma(canvas_cli.image_data)
                            
                            informes_finales = []
                            with st.spinner(f"Procesando documentos de {macro_area}..."):
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
                                        for inf_enviado in informes_area:
                                            if inf_enviado in st.session_state.informes_pendientes: st.session_state.informes_pendientes.remove(inf_enviado)
                                        guardar_pendientes(st.session_state.usuario_actual, st.session_state.informes_pendientes) 
                                        time.sleep(2)
                                        if len(st.session_state.informes_pendientes) == 0: volver_catalogo()
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