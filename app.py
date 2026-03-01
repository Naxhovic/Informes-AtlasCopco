import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import os, subprocess, io, base64, smtplib, json, uuid, time
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from streamlit_drawable_canvas import st_canvas
from PIL import Image
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
    cuerpo = f"Se adjuntan {len(lista_informes)} reportes de servicio técnico (Firmados).\n"
    msg.attach(MIMEText(cuerpo, 'plain'))

    for item in lista_informes:
        ruta = item['ruta']
        if os.path.exists(ruta):
            with open(ruta, "rb") as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{item["nombre_archivo"]}"')
            msg.attach(part)

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(CORREO_REMITENTE, PASSWORD_APLICACION)
        server.send_message(msg); server.quit()
        return True, "✅ Enviado con éxito."
    except Exception as e: return False, str(e)

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
# 2. CONEXIÓN INMORTAL A GOOGLE SHEETS (ESCUDO ANTI-BLOQUEO 429)
# =============================================================================
@st.cache_resource
def get_gspread_client():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds_dict = json.loads(st.secrets["gcp_json"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return gspread.authorize(creds)

def get_sheet(sheet_name):
    time.sleep(3) # Demora obligatoria antes de llamar a Google
    try:
        client = get_gspread_client()
        doc = client.open("BaseDatos")
        pestañas = [hoja.title for hoja in doc.worksheets()]
        if sheet_name in pestañas: return doc.worksheet(sheet_name)
        return doc.add_worksheet(title=sheet_name, rows="1000", cols="20")
    except Exception as e:
        st.error(f"🚨 ERROR DE CONEXIÓN CON GOOGLE: {e}")
        return None

# --- FUNCIONES DE BASE DE DATOS CON DELAY ---
def guardar_dato_equipo(tag, clave, valor):
    time.sleep(3)
    try:
        sheet = get_sheet("datos_equipo")
        sheet.append_row([tag, clave, valor])
        st.cache_data.clear()
    except: pass

@st.cache_data(ttl=60, show_spinner=False)
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
    time.sleep(3)
    fecha = pd.Timestamp.now().strftime("%d/%m/%Y %H:%M")
    try:
        sheet = get_sheet("observaciones")
        sheet.append_row([str(uuid.uuid4())[:8], tag, fecha, usuario.title(), texto.strip(), "ACTIVO"])
        st.cache_data.clear()
    except: pass

@st.cache_data(ttl=60, show_spinner=False)
def obtener_observaciones(tag):
    try:
        sheet = get_sheet("observaciones")
        data = sheet.get_all_values()
        obs = [{"id": r[0], "fecha": r[2], "usuario": r[3], "texto": r[4]} for r in data if len(r) >= 6 and r[1] == tag and r[5] == "ACTIVO"]
        df = pd.DataFrame(obs)
        return df.iloc[::-1] if not df.empty else pd.DataFrame(columns=["id", "fecha", "usuario", "texto"])
    except: return pd.DataFrame(columns=["id", "fecha", "usuario", "texto"])

def eliminar_observacion(id_obs):
    time.sleep(3)
    try:
        sheet = get_sheet("observaciones")
        cell = sheet.find(id_obs)
        if cell: 
            sheet.update_cell(cell.row, 6, "ELIMINADO")
            st.cache_data.clear()
    except: pass

def guardar_especificacion_db(modelo, clave, valor):
    time.sleep(3)
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
        data = sheet.get_all_values()
        for row in data:
            if len(row) >= 3:
                mod, clave, valor = row[0], row[1], row[2]
                if mod not in specs: specs[mod] = {}
                specs[mod][clave] = valor
    except: pass
    return specs

@st.cache_data(ttl=60, show_spinner=False)
def obtener_contactos():
    try:
        sheet = get_sheet("contactos")
        data = sheet.get_all_values()
        contactos = [row[0] for row in data if len(row) > 1 and row[1] == "ACTIVO"]
        return sorted(list(set(contactos))) if contactos else ["Lorena Rojas"]
    except: return ["Lorena Rojas"]

def agregar_contacto(nombre):
    time.sleep(3)
    try:
        sheet = get_sheet("contactos")
        sheet.append_row([nombre.strip().title(), "ACTIVO"])
        st.cache_data.clear()
    except: pass

def eliminar_contacto(nombre):
    time.sleep(3)
    try:
        sheet = get_sheet("contactos")
        cells = sheet.findall(nombre)
        for cell in cells: sheet.update_cell(cell.row, 2, "ELIMINADO")
        st.cache_data.clear()
    except: pass

def guardar_registro(data_tuple):
    time.sleep(3)
    try:
        sheet = get_sheet("intervenciones")
        sheet.append_row([str(x) for x in data_tuple])
        st.cache_data.clear()
    except: pass

@st.cache_data(ttl=60, show_spinner=False)
def buscar_ultimo_registro(tag):
    try:
        sheet = get_sheet("intervenciones")
        data = sheet.get_all_values()
        for row in reversed(data):
            if len(row) >= 20 and row[0] == tag:
                return (row[5], row[6], row[9], row[14], row[15], row[7], row[8], row[10], row[11], row[12], row[13], row[16], row[17])
    except: pass
    return None

@st.cache_data(ttl=60, show_spinner=False)
def obtener_todo_el_historial(tag):
    try:
        sheet = get_sheet("intervenciones")
        data = sheet.get_all_values()
        hist = [{"fecha": r[5], "tipo_intervencion": r[15], "estado_equipo": r[17], "Cuenta Usuario": r[19], "horas_marcha": r[12], "horas_carga": r[13], "p_carga": r[10], "p_descarga": r[11], "temp_salida": r[9]} for r in data if len(r) >= 20 and r[0] == tag]
        df = pd.DataFrame(hist)
        return df.iloc[::-1] if not df.empty else pd.DataFrame()
    except: return pd.DataFrame()

@st.cache_data(ttl=60, show_spinner=False)
def obtener_estados_actuales():
    try:
        sheet = get_sheet("intervenciones")
        data = sheet.get_all_values()
        return {row[0]: row[17] for row in data if len(row) >= 18}
    except: return {}

# =============================================================================
# 3. CONVERSIÓN A PDF HÍBRIDA
# =============================================================================
def convertir_a_pdf(ruta_docx):
    ruta_pdf = ruta_docx.replace(".docx", ".pdf")
    try:
        comando = ['libreoffice', '--headless', '--convert-to', 'pdf', os.path.abspath(ruta_docx), '--outdir', os.path.dirname(os.path.abspath(ruta_docx))]
        subprocess.run(comando, capture_output=True)
        if os.path.exists(ruta_pdf): return ruta_pdf
    except: pass
    return None
# =============================================================================
# 4. VARIABLES DE SESIÓN Y NAVEGACIÓN
# =============================================================================
ESPECIFICACIONES = obtener_especificaciones(DEFAULT_SPECS)
default_states = {'logged_in': False, 'usuario_actual': "", 'equipo_seleccionado': None, 'input_cliente': "Lorena Rojas", 'input_tec1': "Ignacio Morales", 'input_tec2': "emian Sanchez", 'input_h_marcha': 0, 'input_h_carga': 0, 'input_temp': "70.0", 'input_p_carga': "7.0", 'input_p_descarga': "7.5", 'input_estado': "", 'input_reco': "", 'input_estado_eq': "Operativo", 'informes_pendientes': [], 'vista_firmas': False}
for k, v in default_states.items():
    if k not in st.session_state: st.session_state[k] = v

def seleccionar_equipo(tag):
    st.session_state.equipo_seleccionado = tag
    st.session_state.vista_firmas = False
    reg = buscar_ultimo_registro(tag)
    if reg:
        st.session_state.update({'input_cliente': reg[1], 'input_tec1': reg[5], 'input_tec2': reg[6], 'input_estado': reg[3], 'input_reco': reg[11] or "", 'input_estado_eq': reg[12] or "Operativo", 'input_h_marcha': int(reg[9] or 0), 'input_h_carga': int(reg[10] or 0), 'input_temp': str(reg[2]).replace(',', '.') if reg[2] else "70.0"})
        try: st.session_state.input_p_carga = str(reg[7]).split()[0].replace(',', '.')
        except: st.session_state.input_p_carga = "7.0"
        try: st.session_state.input_p_descarga = str(reg[8]).split()[0].replace(',', '.')
        except: st.session_state.input_p_descarga = "7.5"
    else:
        st.session_state.update({'input_estado_eq': "Operativo", 'input_estado': "", 'input_reco': "", 'input_h_marcha': 0, 'input_h_carga': 0})

def volver_catalogo(): 
    st.session_state.equipo_seleccionado = None; st.session_state.vista_firmas = False

# --- PANTALLA LOGIN ---
if not st.session_state.logged_in:
    st.markdown("<br><br>", unsafe_allow_html=True)
    _, col, _ = st.columns([1, 1.5, 1])
    with col.container(border=True):
        st.markdown("<h1 style='text-align: center; border:none;'>⚙️ <span style='color:#007CA6;'>Atlas</span> <span style='color:#FF6600;'>Spence</span></h1>", unsafe_allow_html=True)
        with st.form("login_f"):
            u, p = st.text_input("Usuario").lower(), st.text_input("Contraseña", type="password")
            if st.form_submit_button("Ingresar", use_container_width=True, type="primary"):
                if u in USUARIOS and USUARIOS[u] == p: st.session_state.update({'logged_in': True, 'usuario_actual': u}); st.rerun()
                else: st.error("❌ Credenciales inválidas.")

# --- APP PRINCIPAL ---
else:
    with st.sidebar:
        st.markdown(f"**Usuario:** {st.session_state.usuario_actual.title()}")
        if st.session_state.informes_pendientes:
            if st.button(f"✍️ Ir a Firmas ({len(st.session_state.informes_pendientes)})", type="primary", use_container_width=True):
                st.session_state.vista_firmas = True; st.session_state.equipo_seleccionado = None; st.rerun()
        if st.button("🚪 Salir", use_container_width=True): st.session_state.logged_in = False; st.rerun()

    # --- 6.1 PIZARRA DE FIRMAS (NUEVO VISOR ROBUSTO POR DIBUJO) ---
    if st.session_state.vista_firmas:
        if st.button("⬅️ Volver"): volver_catalogo(); st.rerun()
        st.title("✍️ Pizarra de Firmas Digital")
        for i, inf in enumerate(st.session_state.informes_pendientes):
            with st.expander(f"📄 Revisar: {inf['tag']} ({inf['tipo_plan']})"):
                if inf.get('ruta_prev_pdf') and os.path.exists(inf['ruta_prev_pdf']):
                    try:
                        with open(inf['ruta_prev_pdf'], "rb") as f_pdf: pdf_bytes = f_pdf.read()
                        pdf_viewer(pdf_bytes, width=700, height=600)
                    except: st.error("No se pudo dibujar el PDF.")
                    with open(inf['ruta_prev_pdf'], "rb") as f2:
                        st.download_button("📥 Descargar Borrador", f2, file_name=f"Borrador_{inf['tag']}.pdf", key=f"dl_{i}")
        
        c1, c2 = st.columns(2)
        with c1: st.markdown("🧑‍🔧 **Técnico**"); can_t = st_canvas(stroke_width=4, height=200, width=400, key="ct")
        with c2: st.markdown("👷 **Cliente**"); can_c = st_canvas(stroke_width=4, height=200, width=400, key="cc")

        if st.button("🚀 Firmar y Subir a Nube", type="primary", use_container_width=True):
            if can_t.image_data is not None and can_c.image_data is not None:
                with st.spinner("Fabricando documentos oficiales..."):
                    img_t, img_c = io.BytesIO(), io.BytesIO()
                    Image.fromarray(can_t.image_data.astype('uint8'), 'RGBA').save(img_t, 'PNG')
                    Image.fromarray(can_c.image_data.astype('uint8'), 'RGBA').save(img_c, 'PNG')
                    
                    finales = []
                    for inf in st.session_state.informes_pendientes:
                        doc = DocxTemplate(inf['file_plantilla']); ctx = inf['context']
                        ctx['firma_tecnico'], ctx['firma_cliente'] = InlineImage(doc, img_t, width=Mm(40)), InlineImage(doc, img_c, width=Mm(40))
                        doc.render(ctx); doc.save(inf['ruta_docx'])
                        
                        pdf_gen = convertir_a_pdf(inf['ruta_docx']) or inf['ruta_docx']
                        nombre_cod = f"{inf['area'].title()}@@{inf['tag']}@@{inf['nombre_archivo_base'].replace('.docx', '.pdf')}"
                        
                        t_db = list(inf['tupla_db']); t_db[18] = pdf_gen; guardar_registro(tuple(t_db))
                        finales.append({"tag": inf['tag'], "tipo": inf['tipo_plan'], "ruta": pdf_gen, "nombre_archivo": nombre_cod})
                    
                    exito, msg = enviar_carrito_por_correo(MI_CORREO_CORPORATIVO, finales)
                    if exito: st.success("✅ ¡Listo!"); st.session_state.informes_pendientes = []; st.balloons()
                    else: st.error(msg)

    # --- 6.2 PANEL PRINCIPAL ---
    elif st.session_state.equipo_seleccionado is None:
        st.title("🏭 Panel de Control")
        est_actuales = obtener_estados_actuales()
        c_f, c_b = st.columns([1.2, 2])
        filt = c_f.radio("🗂️ Categoría:", ["Todos", "Compresores", "Secadores"], horizontal=True)
        busq = c_b.text_input("🔍 Buscar TAG o Área...").lower()
        
        cols = st.columns(4); count = 0
        for tag, (mod, ser, area, ubi) in inventario_equipos.items():
            if (filt == "Compresores" and "CD" in mod.upper()) or (filt == "Secadores" and "CD" not in mod.upper()): continue
            if busq in tag.lower() or busq in area.lower() or busq in mod.lower():
                est = est_actuales.get(tag, "Operativo")
                bg, txt, ico = ("#eaffea", "#004d00", "🟢") if est == "Operativo" else ("#ffeaea", "#800000", "🔴")
                with cols[count % 4].container(border=True):
                    st.markdown(f"<span style='background:{bg}; color:{txt}; padding:4px 8px; border-radius:4px; font-weight:bold;'>{ico} {est.upper()}</span>", unsafe_allow_html=True)
                    st.markdown(f"### {tag}\n**{mod}** | {area.title()}")
                    st.button("📝 Ingresar", key=f"btn_{tag}", on_click=seleccionar_equipo, args=(tag,), use_container_width=True)
                count += 1

    # --- 6.3 FICHA (4 PESTAÑAS) ---
    else:
        tag = st.session_state.equipo_seleccionado; mod, ser, area, ubi = inventario_equipos[tag]
        st.button("⬅️ Volver", on_click=volver_catalogo, use_container_width=True)
        st.title(f"Ficha de Servicio: {tag}")
        
        t1, t2, t3, t4 = st.tabs(["📋 1. Reporte", "📚 2. Ficha Técnica", "🔍 3. Bitácora", "👤 4. Área"])
        
        with t1:
            plan = st.selectbox("🛠️ Orden:", ["Inspección", "PM03"] if "CD" in tag else ["Inspección", "P1", "P2", "P3", "PM03"])
            c1, c2, c3, c4 = st.columns(4)
            c1.text_input("Modelo", mod, disabled=True); c2.text_input("Serie", ser, disabled=True); c3.text_input("Área", area, disabled=True); c4.text_input("Ubicación", ubi, disabled=True)
            
            c5, c6, c7, c8 = st.columns([1, 1, 1, 1.3])
            fec = c5.text_input("Fecha", "25 de febrero de 2026")
            t1_n = c6.text_input("Técnico 1", key="input_tec1"); t2_n = c7.text_input("Técnico 2", key="input_tec2")
            
            conts = obtener_contactos(); ops = ["➕ Nuevo..."] + conts
            cli = c8.selectbox("Cliente", ops, index=ops.index(st.session_state.input_cliente) if st.session_state.input_cliente in ops else 1)
            if cli == "➕ Nuevo...":
                nc = c8.text_input("Nombre Cliente:")
                if c8.button("Guardar"): agregar_contacto(nc); st.session_state.input_cliente = nc.title(); st.rerun()
            else: st.session_state.input_cliente = cli

            m1, m2, m3, m4, m5, m6 = st.columns(6)
            hm = m1.number_input("H. Marcha", value=int(st.session_state.input_h_marcha))
            hc = m2.number_input("H. Carga", value=int(st.session_state.input_h_carga))
            un = m3.selectbox("Unidad", ["Bar", "psi"])
            pc, p_desc, ts = m4.text_input("P. Carga", st.session_state.input_p_carga), m5.text_input("P. Descarga", st.session_state.input_p_descarga), m6.text_input("Temp Salida", st.session_state.input_temp)

            e_eq = st.radio("Estado Entrega:", ["Operativo", "Fuera de servicio"], horizontal=True)
            cond = st.text_area("Condición Final:", key="input_estado")
            reco = st.text_area("Recomendaciones:", key="input_reco")

            if st.button("📥 Añadir a Bandeja", type="primary", use_container_width=True):
                tpl = "plantilla/secadorfueradeservicio.docx" if "CD" in tag and e_eq == "Fuera de servicio" else ("plantilla/inspeccionsecador.docx" if "CD" in tag else (f"plantilla/{plan.lower()}.docx" if e_eq == "Operativo" and plan in ["P1","P2","P3"] else ("plantilla/fueradeservicio.docx" if e_eq == "Fuera de servicio" else "plantilla/inspeccion.docx")))
                ctx = {"tipo_intervencion": plan, "modelo": mod, "tag": tag, "area": area, "ubicacion": ubi, "cliente_contacto": st.session_state.input_cliente, "p_carga": f"{pc} {un}", "p_descarga": f"{p_desc} {un}", "temp_salida": ts, "horas_marcha": hm, "horas_carga": hc, "tecnico_1": t1_n, "tecnico_2": t2_n, "estado_equipo": e_eq, "estado_entrega": cond, "recomendaciones": reco, "serie": ser, "fecha": fec, "firma_tecnico": "", "firma_cliente": ""}
                
                fname = f"Inf_{plan}_{tag}.docx"; ruta = os.path.join(RUTA_ONEDRIVE, fname); os.makedirs(RUTA_ONEDRIVE, exist_ok=True)
                with st.spinner("Creando borrador..."):
                    d_p = DocxTemplate(tpl); d_p.render(ctx); d_p.save(os.path.join(RUTA_ONEDRIVE, f"PRE_{fname}"))
                    r_p = convertir_a_pdf(os.path.join(RUTA_ONEDRIVE, f"PRE_{fname}"))
                
                t_db = (tag, mod, ser, area, ubi, fec, st.session_state.input_cliente, t1_n, t2_n, float(ts.replace(',','.') or 0.0), f"{pc} {un}", f"{p_desc} {un}", hm, hc, cond, plan, reco, e_eq, "", st.session_state.usuario_actual)
                st.session_state.informes_pendientes.append({"tag": tag, "area": area, "tec1": t1_n, "cli": st.session_state.input_cliente, "tipo_plan": plan, "file_plantilla": tpl, "context": ctx, "tupla_db": t_db, "ruta_docx": ruta, "nombre_archivo_base": fname, "ruta_prev_pdf": r_p})
                st.success("✅ Guardado."); navegar()

        with t2:
            st.markdown(f"### 📘 Ficha Técnica ({mod})")
            with st.expander("✏️ Agregar Dato"):
                with st.form("fs"):
                    c_e1, c_e2 = st.columns(2)
                    k = c_e1.selectbox("Dato:", ["N° Parte Filtro Aceite", "N° Parte Kit", "Litros de Aceite", "Otro..."])
                    if k == "Otro...": k = c_e1.text_input("Nombre:")
                    v = c_e2.text_input("Valor:")
                    if st.form_submit_button("💾 Guardar"): guardar_especificacion_db(mod, k.strip(), v.strip()); st.rerun()
            s = ESPECIFICACIONES.get(mod, {})
            cols = st.columns(3); idx = 0
            for k, v in s.items():
                if k != "Manual":
                    with cols[idx % 3]: st.markdown(f"<div style='background:#1e2530; padding:15px; border-radius:8px; margin-bottom:10px; border-left:4px solid #007CA6;'><small style='color:#8c9eb5;'>{k}</small><br>{v}</div>", unsafe_allow_html=True)
                    idx += 1
            if "Manual" in s and os.path.exists(s["Manual"]):
                with open(s["Manual"], "rb") as f: st.download_button(f"📕 Manual {mod}", f, file_name=f"{mod}.pdf")

        with t3:
            st.markdown(f"### 🔍 Bitácora: {tag}")
            with st.form("fo"):
                n_o = st.text_area("Nueva nota técnica:"); 
                if st.form_submit_button("➕ Guardar"): agregar_observacion(tag, st.session_state.usuario_actual, n_o); st.rerun()
            for _, r in obtener_observaciones(tag).iterrows():
                col_o, col_d = st.columns([11, 1])
                with col_o: st.markdown(f"<div style='background:#2b303b; padding:15px; border-radius:8px; margin-bottom:10px; border-left:4px solid #FF6600;'><small>{r['usuario']} | {r['fecha']}</small><br>{r['texto']}</div>", unsafe_allow_html=True)
                with col_d: 
                    if st.button("🗑️", key=f"d_{r['id']}"): eliminar_observacion(r['id']); st.rerun()

        with t4:
            st.markdown(f"### 👤 Gestión de Área: {tag}")
            with st.expander("✏️ Editar Datos"):
                with st.form("fa"):
                    ka = st.selectbox("Cargo:", ["Dueño de Área", "PEA", "Frecuencia Radial", "Supervisor", "Otro..."])
                    if ka == "Otro...": ka = st.text_input("Nombre cargo:")
                    va = st.text_input("Valor:")
                    if st.form_submit_button("Actualizar"): guardar_dato_equipo(tag, ka.strip(), va.strip()); st.rerun()
            d_eq = obtener_datos_equipo(tag)
            cols = st.columns(2); i = 0
            for k, v in d_eq.items():
                with cols[i % 2]: st.markdown(f"<div style='background:#2b303b; padding:15px; border-radius:8px; margin-bottom:10px; border-left:4px solid #FF6600;'><span style='color:#aeb9cc;'>{k}</span><br>{v}</div>", unsafe_allow_html=True)
                i += 1

        st.markdown("---"); st.markdown("### 📋 Historial de Intervenciones")
        df_h = obtener_todo_el_historial(tag)
        if not df_h.empty: st.dataframe(df_h, use_container_width=True)