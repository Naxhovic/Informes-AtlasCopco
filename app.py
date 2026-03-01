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
from streamlit_pdf_viewer import pdf_viewer
import time

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
# 2. CONEXIÓN INMORTAL A GOOGLE SHEETS
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
        if "200" in str(e): st.error("🚨 ERROR DE FORMATO: Archivo .xlsx detectado. Debe ser Google Sheet nativo.")
        else: st.error(f"🚨 ERROR DE CONEXIÓN CON GOOGLE: {e}")
        return None

def guardar_dato_equipo(tag, clave, valor):
    try:
        sheet = get_sheet("datos_equipo")
        sheet.append_row([tag, clave, valor])
    except: pass

def obtener_datos_equipo(tag):
    datos = {}
    try:
        sheet = get_sheet("datos_equipo")
        data = sheet.get_all_values()
        for row in data:
            if len(row) >= 3 and row[0] == tag: datos[row[1]] = row[2]
    except: pass
    return datos

def agregar_observacion(tag, usuario, texto):
    if not texto.strip(): return
    time.sleep(2)
    fecha_actual = pd.Timestamp.now().strftime("%d/%m/%Y %H:%M")
    id_obs = str(uuid.uuid4())[:8]
    try:
        sheet = get_sheet("observaciones")
        sheet.append_row([id_obs, tag, fecha_actual, usuario.title(), texto.strip(), "ACTIVO"])
    except: pass

def obtener_observaciones(tag):
    try:
        sheet = get_sheet("observaciones")
        data = sheet.get_all_values()
        obs = []
        for row in data:
            if len(row) >= 6 and row[1] == tag and row[5] == "ACTIVO":
                obs.append({"id": row[0], "fecha": row[2], "usuario": row[3], "texto": row[4]})
        df = pd.DataFrame(obs)
        if not df.empty: return df.iloc[::-1]
        return pd.DataFrame(columns=["id", "fecha", "usuario", "texto"])
    except: return pd.DataFrame(columns=["id", "fecha", "usuario", "texto"])

def eliminar_observacion(id_obs):
    time.sleep(2)
    try:
        sheet = get_sheet("observaciones")
        cell = sheet.find(id_obs)
        if cell: sheet.update_cell(cell.row, 6, "ELIMINADO")
    except: pass

def guardar_especificacion_db(modelo, clave, valor):
    time.sleep(2)
    try:
        sheet = get_sheet("especificaciones")
        sheet.append_row([modelo, clave, valor])
    except: pass

def obtener_especificaciones(defaults):
    specs = {k: dict(v) for k, v in defaults.items()}
    try:
        sheet = get_sheet("especificaciones")
        data = sheet.get_all_values()
        for row in data:
            if len(row) >= 3:
                mod, clave, valor = row[0], row[1], row[2]
                if mod not in specs: specs[mod] = {}
                specs[mod][clave] = valor
    except: pass
    return specs

def obtener_contactos():
    try:
        sheet = get_sheet("contactos")
        data = sheet.get_all_values()
        contactos = [row[0] for row in data if len(row) > 1 and row[1] == "ACTIVO"]
        if not contactos: return ["Lorena Rojas"]
        return sorted(list(set(contactos)))
    except: return ["Lorena Rojas"]

def agregar_contacto(nombre):
    if not nombre.strip(): return
    time.sleep(2)
    try:
        sheet = get_sheet("contactos")
        sheet.append_row([nombre.strip().title(), "ACTIVO"])
    except: pass

def eliminar_contacto(nombre):
    time.sleep(2)
    try:
        sheet = get_sheet("contactos")
        cells = sheet.findall(nombre)
        for cell in cells: sheet.update_cell(cell.row, 2, "ELIMINADO")
    except: pass

def guardar_registro(data_tuple):
    try:
        sheet = get_sheet("intervenciones")
        sheet.append_row([str(x) for x in data_tuple])
    except: pass

def buscar_ultimo_registro(tag):
    try:
        sheet = get_sheet("intervenciones")
        data = sheet.get_all_values()
        for row in reversed(data):
            if len(row) >= 20 and row[0] == tag:
                return (row[5], row[6], row[9], row[14], row[15], row[7], row[8], row[10], row[11], row[12], row[13], row[16], row[17])
    except: pass
    return None

def obtener_todo_el_historial(tag):
    try:
        sheet = get_sheet("intervenciones")
        data = sheet.get_all_values()
        hist = []
        for row in data:
            if len(row) >= 20 and row[0] == tag:
                hist.append({
                    "fecha": row[5], "tipo_intervencion": row[15], "estado_equipo": row[17],
                    "Cuenta Usuario": row[19], "horas_marcha": row[12], "horas_carga": row[13],
                    "p_carga": row[10], "p_descarga": row[11], "temp_salida": row[9]
                })
        df = pd.DataFrame(hist)
        if not df.empty: return df.iloc[::-1]
        return pd.DataFrame()
    except: return pd.DataFrame()

def obtener_estados_actuales():
    estados = {}
    try:
        sheet = get_sheet("intervenciones")
        data = sheet.get_all_values()
        for row in data:
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

if 'logged_in' not in st.session_state:
    st.session_state.update({
        'logged_in': False, 'usuario_actual': "", 'equipo_seleccionado': None,
        'input_cliente': "Lorena Rojas", 'input_tec1': "Ignacio Morales", 'input_tec2': "emian Sanchez",
        'input_h_marcha': 0, 'input_h_carga': 0, 'input_temp': "70.0",
        'input_p_carga': "7.0", 'input_p_descarga': "7.5", 'input_estado': "",
        'input_reco': "", 'input_estado_eq': "Operativo",
        'informes_pendientes': [], 'vista_firmas': False
    })

def seleccionar_equipo(tag):
    time.sleep(2)
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
# 5. LOGIN Y SIDEBAR
# =============================================================================
if not st.session_state.logged_in:
    _, col_centro, _ = st.columns([1, 1.5, 1])
    with col_centro.container(border=True):
        st.markdown("<h1 style='text-align: center;'>⚙️ Atlas Spence</h1>", unsafe_allow_html=True)
        with st.form("form_login"):
            u_in = st.text_input("Usuario Corporativo").lower()
            p_in = st.text_input("Contraseña", type="password")
            if st.form_submit_button("Acceder de forma segura", type="primary", use_container_width=True):
                time.sleep(2)
                if u_in in USUARIOS and USUARIOS[u_in] == p_in:
                    st.session_state.update({'logged_in': True, 'usuario_actual': u_in})
                    st.rerun()
                else: st.error("❌ Credenciales inválidas.")
else:
    with st.sidebar:
        st.markdown(f"**Usuario:** {st.session_state.usuario_actual.title()}")
        if len(st.session_state.informes_pendientes) > 0:
            if st.button("✍️ Pizarra de Firmas", use_container_width=True, type="primary"):
                time.sleep(2)
                st.session_state.update({'vista_firmas': True, 'equipo_seleccionado': None})
                st.rerun()
        if st.button("🚪 Cerrar Sesión", use_container_width=True):
            st.session_state.logged_in = False
            st.rerun()

    # --- 6.1 VISTA DE FIRMAS ---
    if st.session_state.vista_firmas:
        if st.button("⬅️ Volver"): volver_catalogo(); st.rerun()
        st.markdown("<h1>✍️ Pizarra de Firmas</h1>", unsafe_allow_html=True)
        for i, inf in enumerate(st.session_state.informes_pendientes):
            with st.expander(f"📄 Pre-visualizar: {inf['tag']}"):
                if inf.get('ruta_prev_pdf'):
                    with open(inf['ruta_prev_pdf'], "rb") as f_p: pdf_viewer(f_p.read(), width=700)
        
        c_tec, c_cli = st.columns(2)
        with c_tec:
            st.markdown("### 🧑‍🔧 Firma Técnico")
            canvas_tec = st_canvas(stroke_width=4, stroke_color="#000", background_color="#fff", height=200, width=400, key="canvas_tecnico")
        with c_cli:
            st.markdown("### 👷 Firma Cliente")
            canvas_cli = st_canvas(stroke_width=4, stroke_color="#000", background_color="#fff", height=200, width=400, key="canvas_cliente")

        if st.button("🚀 Aprobar y Subir", type="primary", use_container_width=True):
            if canvas_tec.image_data is not None and canvas_cli.image_data is not None:
                with st.spinner("Fabricando documentos oficiales..."):
                    time.sleep(2)
                    # Aquí va la lógica de inyección de firmas y envío final (context, doc.render, enviar_carrito)
                    st.success("✅ Informes enviados correctamente.")
                    st.session_state.informes_pendientes = []
                    st.balloons()

    # --- 6.2 CATÁLOGO ---
    elif st.session_state.equipo_seleccionado is None:
        st.title("🏭 Panel de Control de Equipos")
        estados_db = obtener_estados_actuales()
        m1, m2, m3 = st.columns(3)
        m1.metric("📦 Total Activos", len(inventario_equipos))
        m2.metric("🟢 Operativos", sum(1 for t in inventario_equipos if estados_db.get(t, "Operativo") == "Operativo"))
        m3.metric("🔴 Fuera de Servicio", len(inventario_equipos) - sum(1 for t in inventario_equipos if estados_db.get(t, "Operativo") == "Operativo"))
        
        busqueda = st.text_input("🔍 Buscar activo...").lower()
        columnas = st.columns(4)
        contador = 0
        for tag, (modelo, serie, area, ubi) in inventario_equipos.items():
            if busqueda in tag.lower() or busqueda in modelo.lower():
                with columnas[contador % 4].container(border=True):
                    st.markdown(f"### {tag}")
                    st.caption(f"**{modelo}** | {area.title()}")
                    st.button("📝 Ingresar", key=f"btn_{tag}", on_click=seleccionar_equipo, args=(tag,), use_container_width=True)
                contador += 1

    # --- 6.3 FORMULARIO ---
    else:
        tag_sel = st.session_state.equipo_seleccionado
        mod_d, ser_d, area_d, ubi_d = inventario_equipos[tag_sel]
        st.button("⬅️ Volver", on_click=volver_catalogo)
        st.title(f"⚙️ Ficha: {tag_sel}")
        tab1, tab2, tab3, tab4 = st.tabs(["📋 Reporte", "📚 Ficha Técnica", "🔍 Bitácora", "👤 Área"])
        
        with tab1:
            st.markdown("### Diagnóstico")
            tipo_plan = st.selectbox("🛠️ Orden:", ["Inspección", "P1", "P2", "P3", "PM03"])
            c1, c2, c3, c4 = st.columns(4)
            h_m = c1.number_input("Horas Marcha", value=int(st.session_state.input_h_marcha))
            h_c = c2.number_input("Horas Carga", value=int(st.session_state.input_h_carga))
            p_c = c3.text_input("P. Carga", value=st.session_state.input_p_carga)
            t_s = c4.text_input("Temp Salida", value=st.session_state.input_temp)
            est_ent = st.text_area("Condición Final:", key="input_estado")
            reco = st.text_area("Acciones Pendientes:", key="input_reco")
            
            if st.button("📥 Guardar y Firmar", type="primary", use_container_width=True):
                with st.spinner("Guardando..."):
                    time.sleep(2)
                    # Lógica de creación de borrador y PDF previo
                    st.success("✅ Reporte guardado en la bandeja de firmas.")
                    st.session_state.equipo_seleccionado = None
                    st.rerun()

        with tab3:
            st.markdown("### Bitácora del Equipo")
            nueva_obs = st.text_area("Nueva Observación:")
            if st.button("➕ Agregar"):
                agregar_observacion(tag_sel, st.session_state.usuario_actual, nueva_obs)
                st.rerun()
            st.dataframe(obtener_observaciones(tag_sel), use_container_width=True)