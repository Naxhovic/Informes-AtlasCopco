import streamlit as st
from docxtpl import DocxTemplate
import os, sqlite3, subprocess
import pandas as pd
import smtplib
from email.message import EmailMessage
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# =============================================================================
# 0.1 CONFIGURACIÓN DE NUBE (ENVÍO SILENCIOSO AUTOMÁTICO)
# =============================================================================

# Carpeta temporal donde el PC guardará el archivo antes de enviarlo por correo
RUTA_ONEDRIVE = "Reportes_Temporales" 

# 👇 Aquí el programa enviará el correo oculto para que Power Automate lo atrape 👇
MI_CORREO_CORPORATIVO = "ignacio.a.morales@atlascopco.com"  

CORREO_REMITENTE = "informeatlas.spence@gmail.com"  
PASSWORD_APLICACION = "jbumdljbdpyomnna"  # <-- Recuerda poner tu contraseña real de Gmail aquí

def enviar_carrito_por_correo(destinatario, lista_informes):
    msg = MIMEMultipart()
    msg['From'] = CORREO_REMITENTE
    msg['To'] = destinatario
    msg['Subject'] = f"REVISIÓN PREVIA: Reportes Atlas Copco - {pd.Timestamp.now().strftime('%d/%m/%Y')}"

    cuerpo = f"Estimado/a,\n\nSe adjuntan {len(lista_informes)} reportes de servicio técnico generados en la presente jornada para su revisión previa.\n\nEquipos intervenidos:\n"
    for item in lista_informes:
        cuerpo += f"- TAG: {item['tag']} | Orden: {item['tipo']}\n"
    cuerpo += "\nSaludos cordiales,\nSistema Integrado InforGem"

    msg.attach(MIMEText(cuerpo, 'plain'))

    for item in lista_informes:
        ruta = item['ruta']
        
        # 👇 MAGIA ANTI-OUTLOOK: Quitamos tildes del nombre del adjunto
        nombre_seguro = item["nombre_archivo"].replace("ó","o").replace("í","i").replace("á","a").replace("é","e").replace("ú","u")
        
        if os.path.exists(ruta):
            with open(ruta, "rb") as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
            encoders.encode_base64(part)
            
            # 👇 Forzamos la etiqueta para que Outlook no lo vuelva .bin
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
# 0.2 CONFIGURACIÓN DE PÁGINA Y ESTILOS CORPORATIVOS
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
        div.stButton > button:first-child:hover {
            background-color: var(--ac-dark); transform: translateY(-2px); box-shadow: 0 4px 8px rgba(0,0,0,0.2); color: white;
        }
        .stTextInput>div>div>input:focus, .stNumberInput>div>div>input:focus, .stSelectbox>div>div>select:focus {
            border-color: var(--ac-blue) !important; box-shadow: 0 0 0 1px var(--ac-blue) !important;
        }
        h1, h2, h3 { font-family: 'Segoe UI', sans-serif; font-weight: 700; }
        h1 { border-bottom: 3px solid var(--ac-blue); padding-bottom: 10px; }
        div[data-testid="stVerticalBlock"] div[data-testid="stContainer"] { transition: all 0.3s ease; }
        div[data-testid="stVerticalBlock"] div[data-testid="stContainer"]:hover { border-color: var(--ac-blue); }
        .stTabs [data-baseweb="tab-list"] { gap: 24px; }
        .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; border-radius: 4px 4px 0 0; gap: 1px; padding-top: 10px; padding-bottom: 10px; }
        .stTabs [aria-selected="true"] { background-color: var(--ac-light); border-bottom: 3px solid var(--ac-blue); color: var(--ac-dark); font-weight: 600; }
        </style>
    """, unsafe_allow_html=True)

aplicar_estilos_premium()

# =============================================================================
# 1. DATOS MAESTROS Y CONFIGURACIÓN
# =============================================================================
USUARIOS = {"ignacio morales": "spence2026", "emian": "spence2026", "ignacio veas": "spence2026", "admin": "admin123"}

ESPECIFICACIONES = {
    "GA 18": {"Litros de Aceite": "8 L", "Cant. Filtros Aceite": "1", "N° Parte Filtro Aceite": "1623 0514 00", "Cant. Filtros Aire": "1", "N° Parte Filtro Aire": "1622 1855 01", "N° Parte Separador": "1622 0871 00", "Tipo de Aceite": "Roto Inject Fluid", "Manual": "manuales/manual_ga18.pdf"},
    "GA 30": {"Litros de Aceite": "14.6 L", "Cant. Filtros Aceite": "1", "N° Parte Filtro Aceite": "1613 6105 00", "Cant. Filtros Aire": "1", "N° Parte Filtro Aire": "1613 7407 00", "N° Parte Kit": "2901-0326-00 - 2901 0325 00", "Tipo de Aceite": "Indurance - Xtend Duty", "Manual": "manuales/manual_ga30.pdf"},
    "GA 37": {"Litros de Aceite": "14.6 L", "N° Parte Filtro Aceite": "1613 6105 00", "N° Parte Filtro Aire": "1613 7407 00", "N° Parte Separador": "1613 7408 00", "N° Parte Kit": "2901 1626 00 - 10-1613 8397 02 / 2901-0326-00", "Tipo de Aceite": "Indurance - Xtend Duty", "Manual": "manuales/manual_ga37.pdf"},
    "GA 45": {"Litros de Aceite": "17.9 L", "Cant. Filtros Aceite": "1", "N° Parte Filtro Aceite": "1613 6105 00", "N° Parte Kit": "2901-0326-00 - 2901 0325 00", "Tipo de Aceite": "Indurance - Xtend Duty / Xtend duty", "Manual": "manuales/manual_ga45.pdf"},
    "GA 75": {"Litros de Aceite": "35.2 L", "Manual": "manuales/manual_ga75.pdf"},
    "GA 90": {"Litros de Aceite": "69 L", "Cant. Filtros Aceite": "3", "N° Parte Filtro Aceite": "1613 6105 00", "N° Parte Filtro Aire": "2914 5077 00", "N° Parte Kit": "2901-0776-00", "Manual": "manuales/manual_ga90.pdf"},
    "GA 132": {"Litros de Aceite": "93 L", "Cant. Filtros Aceite": "3", "N° Parte Filtro Aceite": "1613 6105 90", "Cant. Filtros Aire": "1", "N° Parte Filtro Aire": "2914 5077 00", "N° Parte Kit": "2906 0604 00", "Tipo de Aceite": "Indurance / Indurance - Xtend Duty", "Manual": "manuales/manual_ga132.pdf"},
    "GA 250": {"Litros de Aceite": "130 L", "Cant. Filtros Aceite": "3", "Cant. Filtros Aire": "2", "Tipo de Aceite": "Indurance", "Manual": "manuales/manual_ga250.pdf"},
    "ZT 37": {"Litros de Aceite": "11 L / 23 L", "Cant. Filtros Aceite": "1", "N° Parte Filtro Aceite": "1614 8747 00", "Cant. Filtros Aire": "1", "N° Parte Filtro Aire": "1613 7407 00", "N° Parte Kit": "2901-1122-00", "Tipo de Aceite": "Roto Z fluid", "Manual": "manuales/manual_zt37.pdf"},
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
    "TALLER MECÁNICO": ["GA 18", "API335343", "laboratorio", "taller mecánico"]
}

# =============================================================================
# 2. FUNCIONES DE BASE DE DATOS Y ESTADO LOCAL (SQLite)
# =============================================================================
DB_PATH = "historial_equipos.db"

def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute('''CREATE TABLE IF NOT EXISTS intervenciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT, tag TEXT, modelo TEXT, numero_serie TEXT, area TEXT, ubicacion TEXT,
            fecha TEXT, cliente_contacto TEXT, tecnico_1 TEXT, tecnico_2 TEXT, temp_salida REAL, p_carga TEXT, p_descarga TEXT,
            horas_marcha REAL, horas_carga REAL, estado_entrega TEXT, tipo_intervencion TEXT, recomendaciones TEXT, 
            estado_equipo TEXT DEFAULT 'Operativo', ruta_archivo TEXT, generado_por TEXT DEFAULT 'Desconocido'
        )''')
        cols = [info[1] for info in conn.execute("PRAGMA table_info(intervenciones)").fetchall()]
        if "estado_equipo" not in cols: conn.execute("ALTER TABLE intervenciones ADD COLUMN estado_equipo TEXT DEFAULT 'Operativo'")
        if "recomendaciones" not in cols: conn.execute("ALTER TABLE intervenciones ADD COLUMN recomendaciones TEXT")
        if "generado_por" not in cols: conn.execute("ALTER TABLE intervenciones ADD COLUMN generado_por TEXT DEFAULT 'Desconocido'")
        
        try: conn.execute('''CREATE TABLE IF NOT EXISTS contactos (id INTEGER PRIMARY KEY AUTOINCREMENT, nombre TEXT UNIQUE)''')
        except: pass
        cursor = conn.execute("SELECT COUNT(*) FROM contactos")
        if cursor.fetchone()[0] == 0: conn.execute("INSERT INTO contactos (nombre) VALUES ('Lorena Rojas')")

def obtener_contactos():
    try:
        with sqlite3.connect(DB_PATH) as conn:
            return [row[0] for row in conn.execute("SELECT nombre FROM contactos ORDER BY nombre").fetchall()]
    except: return ["Lorena Rojas"]

def agregar_contacto(nombre):
    if not nombre.strip(): return
    try:
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute("INSERT INTO contactos (nombre) VALUES (?)", (nombre.strip().title(),))
    except sqlite3.IntegrityError: pass 

def eliminar_contacto(nombre):
    try:
        with sqlite3.connect(DB_PATH) as conn: conn.execute("DELETE FROM contactos WHERE nombre = ?", (nombre,))
    except: pass

def obtener_estados_actuales():
    try:
        with sqlite3.connect(DB_PATH) as conn:
            return {row[0]: row[1] for row in conn.execute('''SELECT tag, estado_equipo FROM intervenciones WHERE id IN (SELECT MAX(id) FROM intervenciones GROUP BY tag)''').fetchall()}
    except: return {}

def guardar_registro(data_tuple):
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute('''INSERT INTO intervenciones 
            (tag, modelo, numero_serie, area, ubicacion, fecha, cliente_contacto, tecnico_1, tecnico_2, temp_salida, 
            p_carga, p_descarga, horas_marcha, horas_carga, estado_entrega, tipo_intervencion, recomendaciones, estado_equipo, ruta_archivo, generado_por)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''', data_tuple)

def buscar_ultimo_registro(tag):
    with sqlite3.connect(DB_PATH) as conn:
        return conn.execute('''SELECT fecha, cliente_contacto, temp_salida, estado_entrega, tipo_intervencion, tecnico_1, tecnico_2, p_carga, p_descarga, horas_marcha, horas_carga, recomendaciones, estado_equipo FROM intervenciones WHERE tag = ? ORDER BY id DESC LIMIT 1''', (tag,)).fetchone()

def obtener_todo_el_historial(tag):
    with sqlite3.connect(DB_PATH) as conn:
        return pd.read_sql_query("SELECT fecha, tipo_intervencion, estado_equipo, generado_por as 'Cuenta Usuario', horas_marcha, horas_carga, p_carga, p_descarga, temp_salida FROM intervenciones WHERE tag = ? ORDER BY id DESC", conn, params=(tag,))

# =============================================================================
# 3. CONVERSIÓN A PDF HÍBRIDA (NUBE LINUX / LOCAL WINDOWS)
# =============================================================================
def convertir_a_pdf(ruta_docx):
    ruta_pdf = ruta_docx.replace(".docx", ".pdf")
    
    # Intento 1: Servidor en la Nube (Linux + LibreOffice)
    try:
        carpeta_salida = os.path.dirname(ruta_docx)
        if not carpeta_salida: carpeta_salida = "."
        
        # Ejecuta LibreOffice de forma invisible
        comando = ['libreoffice', '--headless', '--convert-to', 'pdf', ruta_docx, '--outdir', carpeta_salida]
        subprocess.run(comando, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        
        if os.path.exists(ruta_pdf): 
            return ruta_pdf
    except Exception:
        pass # Si falla, significa que no estamos en la nube, pasamos al Intento 2

    # Intento 2: Computador Local (Windows + Microsoft Word)
    try:
        import pythoncom
        from docx2pdf import convert
        pythoncom.CoInitialize()
        convert(ruta_docx, ruta_pdf)
        if os.path.exists(ruta_pdf): 
            return ruta_pdf
    except Exception as e:
        print(f"Error PDF Windows: {e}")
        
    return None
# =============================================================================
# 4. INICIALIZACIÓN DE LA APLICACIÓN Y VARIABLES DE SESIÓN
# =============================================================================
init_db()

default_states = {
    'logged_in': False, 'usuario_actual': "", 'equipo_seleccionado': None,
    'input_cliente': "Lorena Rojas", 'input_tec1': "Ignacio Morales", 'input_tec2': "emian Sanchez",
    'input_h_marcha': 0, 'input_h_carga': 0, 'input_temp': "70.0",
    'input_p_carga': "7.0", 'input_p_descarga': "7.5", 'input_estado': "",
    'input_reco': "", 'input_estado_eq': "Operativo",
    'carrito_informes': []
}
for key, value in default_states.items():
    if key not in st.session_state: st.session_state[key] = value

def seleccionar_equipo(tag):
    st.session_state.equipo_seleccionado = tag
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

def volver_catalogo(): st.session_state.equipo_seleccionado = None
def eliminar_del_carrito(idx): st.session_state.carrito_informes.pop(idx)
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
                        st.rerun()
                    else: st.error("❌ Credenciales inválidas.")

# =============================================================================
# 6. PANTALLA PRINCIPAL: APLICACIÓN AUTENTICADA
# =============================================================================
else:
    # Sidebar Corporativo y Carrito
    with st.sidebar:
        st.markdown("<h2 style='text-align: center; border-bottom:none; margin-top: -20px;'><span style='color:#007CA6;'>Atlas</span> <span style='color:#FF6600;'>Spence</span></h2>", unsafe_allow_html=True)
        st.markdown(f"**Usuario Activo:**<br>{st.session_state.usuario_actual.title()}", unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("### 🛒 Historial de Sesión")
        
        if len(st.session_state.carrito_informes) == 0:
            st.info("No hay informes generados recientemente.")
        else:
            for i, item in enumerate(st.session_state.carrito_informes):
                c_name, c_btn = st.columns([4, 1])
                c_name.caption(f"📄 {item['tag']} ({item['tipo']})")
                c_btn.button("❌", key=f"del_cart_{i}", help="Quitar", on_click=eliminar_del_carrito, args=(i,))
                
            st.markdown("<hr style='margin: 10px 0;'>", unsafe_allow_html=True)
            correo_destino = st.text_input("Re-enviar a mi correo:", value=MI_CORREO_CORPORATIVO)
            
            if st.button("✉️ Re-enviar Informes Manualmente", use_container_width=True):
                with st.spinner("Enviando paquete de correos a tu bandeja..."):
                    exito, mensaje = enviar_carrito_por_correo(correo_destino, st.session_state.carrito_informes)
                    if exito:
                        st.success(mensaje)
                        st.session_state.carrito_informes = [] 
                    else:
                        st.error(mensaje)
        
        st.markdown("---")
        if st.button("🚪 Cerrar Sesión", use_container_width=True):
            st.session_state.logged_in = False
            st.rerun()

    # --- 6.1 VISTA CATÁLOGO (Dashboard interactivo) ---
    if st.session_state.equipo_seleccionado is None:
        
        st.markdown("""
            <div style="margin-top: 1.5rem; margin-bottom: 2rem; text-align: center;">
                <div style="background-color: white; height: 2px; width: 100%;"></div>
                <h1 style="color: #007CA6; font-size: 4.5em; font-weight: 900; margin: 20px 0; border-bottom: none; padding: 0;">Atlas Copco</h1>
                <div style="background-color: white; height: 2px; width: 100%;"></div>
            </div>
        """, unsafe_allow_html=True)
        
        st.title("🏭 Panel de Control de Equipos")
        
        estados_db = obtener_estados_actuales()
        total_equipos = len(inventario_equipos)
        
        operativos = sum(1 for tag in inventario_equipos.keys() if estados_db.get(tag, "Operativo") == "Operativo")
        detenidos = total_equipos - operativos
        
        m1, m2, m3 = st.columns(3)
        m1.metric("📦 Total Activos Mineros", total_equipos)
        m2.metric("🟢 Equipos Operativos", operativos)
        m3.metric("🔴 Fuera de Servicio", detenidos)
        
        st.markdown("---")
        
        col_filtro, col_busqueda = st.columns([1.2, 2])
        with col_filtro:
            filtro_tipo = st.radio("🗂️ Categoría de Equipo:", ["Todos", "Compresores", "Secadores"], horizontal=True)
        with col_busqueda:
            busqueda = st.text_input("🔍 Buscar activo por TAG, Modelo o Área...", placeholder="Ejemplo: GA 250, 35-GC-006...").lower()
            
        st.markdown("<br>", unsafe_allow_html=True)
        
        columnas = st.columns(4)
        contador = 0
        for tag, (modelo, serie, area, ubicacion) in inventario_equipos.items():
            
            es_secador = "CD" in modelo.upper()
            if filtro_tipo == "Compresores" and es_secador: continue
            if filtro_tipo == "Secadores" and not es_secador: continue

            if busqueda in tag.lower() or busqueda in area.lower() or busqueda in modelo.lower():
                estado = estados_db.get(tag, "Operativo")
                color_bg = "#eaffea" if estado == "Operativo" else "#ffeaea"
                color_text = "#004d00" if estado == "Operativo" else "#800000"
                icono = "🟢" if estado == "Operativo" else "🔴"
                
                with columnas[contador % 4]:
                    with st.container(border=True):
                        st.markdown(f"<span style='background-color:{color_bg}; color:{color_text}; padding: 4px 8px; border-radius:4px; font-size:0.85em; font-weight:bold; letter-spacing: 0.5px;'>{icono} {estado.upper()}</span>", unsafe_allow_html=True)
                        st.markdown(f"<h3 style='margin-top:10px; margin-bottom:0;'>{tag}</h3>", unsafe_allow_html=True)
                        st.caption(f"**{modelo}** | {area.title()}")
                        st.button("📝 Ingresar", key=f"btn_{tag}", on_click=seleccionar_equipo, args=(tag,), use_container_width=True)
                contador += 1

    # --- 6.2 VISTA FORMULARIO Y GENERACIÓN (Wizard con 3 Tabs) ---
    else:
        tag_sel = st.session_state.equipo_seleccionado
        mod_d, ser_d, area_d, ubi_d = inventario_equipos[tag_sel]
        
        c_btn, c_tit = st.columns([1, 4])
        with c_btn: st.button("⬅️ Volver", on_click=volver_catalogo, use_container_width=True)
        with c_tit: st.markdown(f"<h1 style='margin-top:-15px;'>⚙️ Ficha de Servicio: <span style='color:#007CA6;'>{tag_sel}</span></h1>", unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)

        tab1, tab2, tab3 = st.tabs(["📋 1. Configuración y Parámetros", "📝 2. Diagnóstico Final", "📚 3. Ficha Técnica"])
        
        with tab1:
            st.markdown("### Datos de la Intervención")
            tipo_plan = st.selectbox("🛠️ Tipo de Plan / Orden:", ["Inspección", "PM03"] if "CD" in tag_sel else ["Inspección", "P1", "P2", "P3", "PM03"])
            
            c1, c2, c3, c4 = st.columns(4)
            modelo = c1.text_input("Modelo", mod_d, disabled=True)
            numero_serie = c2.text_input("N° Serie", ser_d, disabled=True)
            area = c3.text_input("Área", area_d, disabled=True)
            ubicacion = c4.text_input("Ubicación", ubi_d, disabled=True)

            c5, c6, c7, c8 = st.columns([1, 1, 1, 1.3])
            fecha = c5.text_input("Fecha Ejecución", "25 de febrero de 2026")
            tec1 = c6.text_input("Técnico 1", key="input_tec1")
            tec2 = c7.text_input("Técnico 2", key="input_tec2")
            
            with c8:
                contactos_db = obtener_contactos()
                opciones = ["➕ Escribir nuevo..."] + contactos_db
                
                if st.session_state.input_cliente in opciones:
                    cli_idx = opciones.index(st.session_state.input_cliente)
                else:
                    cli_idx = 1 if len(contactos_db) > 0 else 0
                
                sc1, sc2 = st.columns([4, 1])
                with sc1:
                    cli_sel = st.selectbox("Contacto Cliente", opciones, index=cli_idx)
                with sc2:
                    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                    if cli_sel != "➕ Escribir nuevo...":
                        if st.button("❌", help="Eliminar permanentemente"):
                            eliminar_contacto(cli_sel)
                            st.session_state.input_cliente = obtener_contactos()[0] if obtener_contactos() else ""
                            st.rerun()
                
                if cli_sel == "➕ Escribir nuevo...":
                    nuevo_c = st.text_input("Nombre:", placeholder="Ej: Juan Pérez", label_visibility="collapsed")
                    if st.button("💾 Guardar y Seleccionar", use_container_width=True):
                        if nuevo_c.strip():
                            agregar_contacto(nuevo_c)
                            st.session_state.input_cliente = nuevo_c.strip().title()
                            st.rerun()
                    cli_cont = nuevo_c.strip().title()
                else:
                    cli_cont = cli_sel
                    st.session_state.input_cliente = cli_sel

            st.markdown("<hr>", unsafe_allow_html=True)
            
            st.markdown("### Mediciones del Equipo")
            c9, c10, c11, c12, c13, c14 = st.columns(6)
            
            h_m = c9.number_input("Horas Marcha Totales", step=1, value=int(st.session_state.input_h_marcha), format="%d")
            h_c = c10.number_input("Horas en Carga", step=1, value=int(st.session_state.input_h_carga), format="%d")
            unidad_p = c11.selectbox("Unidad de Presión", ["Bar", "psi"])
            
            p_c_str = c12.text_input("P. Carga", value=str(st.session_state.input_p_carga))
            p_d_str = c13.text_input("P. Descarga", value=str(st.session_state.input_p_descarga))
            t_salida_str = c14.text_input("Temp Salida (°C)", value=str(st.session_state.input_temp))
            
            p_c_clean = p_c_str.replace(',', '.')
            p_d_clean = p_d_str.replace(',', '.')
            t_salida_clean = t_salida_str.replace(',', '.')

        with tab2:
            st.markdown("### Evaluación y Conclusiones")
            est_eq = st.radio("Estado de Devolución del Activo:", ["Operativo", "Fuera de servicio"], key="input_estado_eq", horizontal=True)
            est_ent = st.text_area("Descripción Condición Final:", key="input_estado", height=100)
            reco = st.text_area("Recomendaciones / Acciones Pendientes:", key="input_reco", height=100)
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            # 👇 BOTÓN CON LA MAGIA DEL ENVÍO SILENCIOSO AUTOMÁTICO 👇
            if st.button("🚀 Generar y Enviar a Nube Central", type="primary", use_container_width=True):
                with st.spinner('Procesando datos y transmitiendo a la nube corporativa...'):
                    try:
                        if "CD" in tag_sel: file_plantilla = "plantilla/secadorfueradeservicio.docx" if est_eq == "Fuera de servicio" else "plantilla/inspeccionsecador.docx"
                        else:
                            if est_eq == "Fuera de servicio": file_plantilla = "plantilla/fueradeservicio.docx"
                            elif tipo_plan == "P1": file_plantilla = "plantilla/p1.docx"
                            elif tipo_plan == "P2": file_plantilla = "plantilla/p2.docx"
                            elif tipo_plan == "P3": file_plantilla = "plantilla/p3.docx"
                            else: file_plantilla = "plantilla/inspeccion.docx"
                            
                        doc = DocxTemplate(file_plantilla)
                        context = {"tipo_intervencion": tipo_plan, "modelo": mod_d, "tag": tag_sel, "area": area_d, "ubicacion": ubi_d, "cliente_contacto": cli_cont, "p_carga": f"{p_c_clean} {unidad_p}", "p_descarga": f"{p_d_clean} {unidad_p}", "temp_salida": t_salida_clean, "horas_marcha": int(h_m), "horas_carga": int(h_c), "tecnico_1": tec1, "tecnico_2": tec2, "estado_equipo": est_eq, "estado_entrega": est_ent, "recomendaciones": reco, "serie": ser_d, "tipo_orden": tipo_plan.upper(), "fecha": fecha, "equipo_modelo": mod_d}
                        
                        doc.render(context)
                        
                        nombre_archivo = f"Informe_{tipo_plan}_{tag_sel}_{fecha.replace(' ','_')}.docx"
                        folder = RUTA_ONEDRIVE
                        os.makedirs(folder, exist_ok=True)
                        ruta = os.path.join(folder, nombre_archivo)
                        doc.save(ruta)
                        
                        ruta_pdf_gen = convertir_a_pdf(ruta)
                        
                        try: temp_db = float(t_salida_clean)
                        except: temp_db = 0.0
                        
                        tupla_db = (tag_sel, mod_d, ser_d, area_d, ubi_d, fecha, cli_cont, tec1, tec2, temp_db, f"{p_c_clean} {unidad_p}", f"{p_d_clean} {unidad_p}", h_m, h_c, est_ent, tipo_plan, reco, est_eq, ruta, st.session_state.usuario_actual)
                        guardar_registro(tupla_db)
                        
                        ruta_final = ruta_pdf_gen if ruta_pdf_gen else ruta
                        nombre_final = nombre_archivo.replace(".docx", ".pdf") if ruta_pdf_gen else nombre_archivo
                        
                        st.session_state.carrito_informes.append({
                            "tag": tag_sel,
                            "tipo": tipo_plan,
                            "ruta": ruta_final,
                            "nombre_archivo": nombre_final
                        })

                        # --- EL ENVÍO SILENCIOSO HACIA POWER AUTOMATE ---
                        informe_actual = [{"tag": tag_sel, "tipo": tipo_plan, "ruta": ruta_final, "nombre_archivo": nombre_final}]
                        exito, mensaje_correo = enviar_carrito_por_correo(MI_CORREO_CORPORATIVO, informe_actual)
                        
                        if exito:
                            st.success(f"✅ ¡Reporte generado y transmitido exitosamente a la Nube Central (OneDrive)!")
                        else:
                            st.warning(f"⚠️ El reporte se generó localmente, pero hubo un error de red al transmitirlo: {mensaje_correo}")
                        
                        c_d1, c_d2 = st.columns(2)
                        with c_d1:
                            with open(ruta, "rb") as f: st.download_button("📄 Obtener Copia Local (Word)", f, file_name=nombre_archivo, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
                        with c_d2:
                            if ruta_pdf_gen:
                                with open(ruta_pdf_gen, "rb") as f_pdf: st.download_button("📕 Obtener Copia Local (PDF)", f_pdf, file_name=nombre_archivo.replace(".docx", ".pdf"), mime="application/pdf", use_container_width=True)
                            else: st.button("📕 PDF (En proceso)", disabled=True, use_container_width=True)
                    except Exception as e: st.error(f"Error sistémico generando reporte: {e}")

        with tab3:
            st.markdown(f"### 📘 Datos Técnicos y Repuestos ({mod_d})")
            if mod_d in ESPECIFICACIONES:
                specs = {k: v for k, v in ESPECIFICACIONES[mod_d].items() if k != "Manual"}
                
                cols = st.columns(3)
                for i, (k, v) in enumerate(specs.items()):
                    with cols[i % 3]:
                        st.markdown(f"""
                            <div style='background-color: #1e2530; padding: 15px; border-radius: 8px; margin-bottom: 15px; border-left: 4px solid #007CA6;'>
                                <span style='color: #8c9eb5; font-size: 0.85em; text-transform: uppercase; font-weight: bold;'>{k}</span><br>
                                <span style='color: white; font-size: 1.1em;'>{v}</span>
                            </div>
                        """, unsafe_allow_html=True)
                
                st.markdown("<hr>", unsafe_allow_html=True)
                st.markdown("### 📥 Documentación y Manuales")
                
                if "Manual" in ESPECIFICACIONES[mod_d] and os.path.exists(ESPECIFICACIONES[mod_d]["Manual"]):
                    with open(ESPECIFICACIONES[mod_d]["Manual"], "rb") as f:
                        st.download_button(label=f"📕 Descargar Manual de {mod_d} (PDF)", data=f, file_name=ESPECIFICACIONES[mod_d]["Manual"].split('/')[-1], mime="application/pdf")
                else:
                    st.info("ℹ️ El manual o despiece para este modelo aún no ha sido cargado en la plataforma.")
            else:
                st.warning("⚠️ No hay especificaciones técnicas registradas para este modelo.")

        st.markdown("<br><hr>", unsafe_allow_html=True)
        st.markdown("### 📋 Trazabilidad Histórica de Intervenciones")
        df_hist = obtener_todo_el_historial(tag_sel)
        if not df_hist.empty: st.dataframe(df_hist, use_container_width=True)