import streamlit as st
from docxtpl import DocxTemplate
import io, os, sqlite3, subprocess
import pandas as pd

# =============================================================================
# 0. CONFIGURACIÓN DE PÁGINA Y ESTILOS CORPORATIVOS (ATLAS COPCO)
# =============================================================================
st.set_page_config(page_title="InforGem | Atlas Copco", layout="wide", page_icon="🟦")

def aplicar_estilos_premium():
    st.markdown("""
        <style>
        /* Paleta Atlas Copco */
        :root {
            --ac-blue: #007CA6;
            --ac-dark: #005675;
            --ac-light: #e6f2f7;
        }
        
        /* Ocultar elementos por defecto de Streamlit para look White-Label */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
        
        /* Estilos de Botones Principales */
        div.stButton > button:first-child {
            background-color: var(--ac-blue);
            color: white;
            border-radius: 6px;
            border: none;
            font-weight: 600;
            padding: 0.5rem 1rem;
            transition: all 0.3s ease;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        div.stButton > button:first-child:hover {
            background-color: var(--ac-dark);
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.2);
            color: white;
        }
        
        /* Foco en inputs */
        .stTextInput>div>div>input:focus, .stNumberInput>div>div>input:focus, .stSelectbox>div>div>select:focus {
            border-color: var(--ac-blue) !important;
            box-shadow: 0 0 0 1px var(--ac-blue) !important;
        }

        /* Títulos */
        h1, h2, h3 {
            color: #1a1a1a;
            font-family: 'Segoe UI', sans-serif;
            font-weight: 700;
        }
        h1 { border-bottom: 3px solid var(--ac-blue); padding-bottom: 10px; }
        
        /* Tarjetas (Containers) en el Catálogo */
        div[data-testid="stVerticalBlock"] div[data-testid="stContainer"] {
            transition: all 0.3s ease;
        }
        div[data-testid="stVerticalBlock"] div[data-testid="stContainer"]:hover {
            border-color: var(--ac-blue);
            background-color: #fafafa;
        }
        
        /* Pestañas (Tabs) */
        .stTabs [data-baseweb="tab-list"] {
            gap: 24px;
        }
        .stTabs [data-baseweb="tab"] {
            height: 50px;
            white-space: pre-wrap;
            border-radius: 4px 4px 0 0;
            gap: 1px;
            padding-top: 10px;
            padding-bottom: 10px;
        }
        .stTabs [aria-selected="true"] {
            background-color: var(--ac-light);
            border-bottom: 3px solid var(--ac-blue);
            color: var(--ac-dark);
            font-weight: 600;
        }
        </style>
    """, unsafe_allow_html=True)

aplicar_estilos_premium()

# =============================================================================
# 1. DATOS MAESTROS Y CONFIGURACIÓN (Diccionarios de la aplicación)
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
    "TALLER": ["GA 18", "API335343", "taller", "laboratorio"]
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
# 3. UTILIDADES: CONVERSIÓN Y SINCRONIZACIÓN EN NUBE
# =============================================================================
def convertir_a_pdf(ruta_docx):
    ruta_pdf = ruta_docx.replace(".docx", ".pdf")
    dir_path = os.path.dirname(ruta_docx)
    try:
        subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", dir_path, ruta_docx], check=True, capture_output=True)
        if os.path.exists(ruta_pdf): return ruta_pdf
    except: pass
    try:
        from docx2pdf import convert
        convert(ruta_docx, ruta_pdf)
        if os.path.exists(ruta_pdf): return ruta_pdf
    except: pass
    return None

def sincronizar_con_nube(tag, tipo_plan):
    try:
        subprocess.run(["git", "add", "."], check=True, capture_output=True, text=True)
        subprocess.run(["git", "commit", "-m", f"Reporte: {tipo_plan} - {tag}"], check=True)
        subprocess.run(["git", "push"], check=True)
        return True, "☁️ Sincronización con Nube Exitosa"
    except: return False, "⚠️ Datos guardados localmente (Pendiente sincronización)"

# =============================================================================
# 4. INICIALIZACIÓN DE LA APLICACIÓN Y VARIABLES DE SESIÓN
# =============================================================================
init_db()

default_states = {
    'logged_in': False, 'usuario_actual': "", 'equipo_seleccionado': None,
    'input_cliente': "Lorena Rojas", 'input_tec1': "Ignacio Morales", 'input_tec2': "emian Sanchez",
    'input_h_marcha': 0.0, 'input_h_carga': 0.0, 'input_temp': 70.0,
    'input_p_carga': 7.0, 'input_p_descarga': 7.5, 'input_estado': "",
    'input_reco': "", 'input_estado_eq': "Operativo"
}
for key, value in default_states.items():
    if key not in st.session_state: st.session_state[key] = value

def seleccionar_equipo(tag):
    st.session_state.equipo_seleccionado = tag
    reg = buscar_ultimo_registro(tag)
    if reg:
        st.session_state.input_cliente, st.session_state.input_tec1, st.session_state.input_tec2 = reg[1], reg[5], reg[6]
        st.session_state.input_estado, st.session_state.input_reco = reg[3], (reg[11] or "")
        st.session_state.input_estado_eq = reg[12] or "Operativo"
        st.session_state.input_temp = float(reg[2])
        st.session_state.input_h_marcha, st.session_state.input_h_carga = float(reg[9] or 0), float(reg[10] or 0)
        try: st.session_state.input_p_carga = float(str(reg[7]).split()[0])
        except: st.session_state.input_p_carga = 7.0
        try: st.session_state.input_p_descarga = float(str(reg[8]).split()[0])
        except: st.session_state.input_p_descarga = 7.5
    else:
        st.session_state.update({'input_estado_eq': "Operativo", 'input_estado': "", 'input_reco': ""})

def volver_catalogo(): st.session_state.equipo_seleccionado = None

# =============================================================================
# 5. PANTALLA 1: SISTEMA DE LOGIN PREMIUM
# =============================================================================
if not st.session_state.logged_in:
    st.markdown("<br><br><br>", unsafe_allow_html=True)
    _, col_centro, _ = st.columns([1, 1.5, 1])
    with col_centro:
        with st.container(border=True):
            st.markdown("<h1 style='text-align: center; color:#007CA6; border-bottom:none;'>🟦 InforGem</h1>", unsafe_allow_html=True)
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
    # Sidebar Corporativo
    with st.sidebar:
        st.markdown(f"<h3 style='color:#007CA6;'>🟦 InforGem</h3>", unsafe_allow_html=True)
        st.markdown(f"**Usuario Activo:**<br>{st.session_state.usuario_actual.title()}", unsafe_allow_html=True)
        st.markdown("---")
        if st.button("🚪 Cerrar Sesión", use_container_width=True):
            st.session_state.logged_in = False
            st.rerun()

    # --- 6.1 VISTA CATÁLOGO (Dashboard interactivo) ---
    if st.session_state.equipo_seleccionado is None:
        st.title("🏭 Panel de Control de Equipos")
        
        estados_db = obtener_estados_actuales()
        total_equipos = len(inventario_equipos)
        operativos = sum(1 for e in estados_db.values() if e == "Operativo")
        detenidos = total_equipos - operativos
        
        # Dashboard Dashboard Metrics
        m1, m2, m3 = st.columns(3)
        m1.metric("📦 Total Activos Mineros", total_equipos)
        m2.metric("🟢 Equipos Operativos", operativos)
        m3.metric("🔴 Fuera de Servicio", detenidos)
        
        st.markdown("---")
        busqueda = st.text_input("🔍 Buscar activo por TAG, Modelo o Área...", placeholder="Ejemplo: GA 250, 35-GC-006...").lower()
        st.markdown("<br>", unsafe_allow_html=True)
        
        columnas = st.columns(4)
        contador = 0
        for tag, (modelo, serie, area, ubicacion) in inventario_equipos.items():
            if busqueda in tag.lower() or busqueda in area.lower() or busqueda in modelo.lower():
                estado = estados_db.get(tag, "Operativo")
                color_bg = "#eaffea" if estado == "Operativo" else "#ffeaea"
                icono = "🟢" if estado == "Operativo" else "🔴"
                
                with columnas[contador % 4]:
                    with st.container(border=True):
                        st.markdown(f"<span style='background-color:{color_bg}; padding: 4px 8px; border-radius:4px; font-size:0.85em; font-weight:bold;'>{icono} {estado.upper()}</span>", unsafe_allow_html=True)
                        st.markdown(f"<h3 style='margin-top:10px; margin-bottom:0;'>{tag}</h3>", unsafe_allow_html=True)
                        st.caption(f"**{modelo}** | {area.title()}")
                        st.button("📝 Ingresar", key=f"btn_{tag}", on_click=seleccionar_equipo, args=(tag,), use_container_width=True)
                contador += 1

    # --- 6.2 VISTA FORMULARIO Y GENERACIÓN (Wizard con Tabs) ---
    else:
        tag_sel = st.session_state.equipo_seleccionado
        mod_d, ser_d, area_d, ubi_d = inventario_equipos[tag_sel]
        
        c_btn, c_tit = st.columns([1, 4])
        with c_btn: st.button("⬅️ Volver", on_click=volver_catalogo, use_container_width=True)
        with c_tit: st.markdown(f"<h1 style='margin-top:-15px;'>⚙️ Ficha de Servicio: <span style='color:#007CA6;'>{tag_sel}</span></h1>", unsafe_allow_html=True)
            
        # Banner de Especificaciones Premium
        if mod_d in ESPECIFICACIONES:
            with st.expander(f"📘 Datos Técnicos y Repuestos ({mod_d})", expanded=False):
                specs = {k: v for k, v in ESPECIFICACIONES[mod_d].items() if k != "Manual"}
                cols_specs = st.columns(len(specs))
                for i, (k, v) in enumerate(specs.items()):
                    with cols_specs[i]: st.metric(k, v)
                if "Manual" in ESPECIFICACIONES[mod_d] and os.path.exists(ESPECIFICACIONES[mod_d]["Manual"]):
                    st.markdown("---")
                    with open(ESPECIFICACIONES[mod_d]["Manual"], "rb") as f:
                        st.download_button(label=f"📥 Descargar Manual {mod_d} (PDF)", data=f, file_name=ESPECIFICACIONES[mod_d]["Manual"].split('/')[-1], mime="application/pdf")
        
        st.markdown("<br>", unsafe_allow_html=True)

        # Creación del Wizard Interactivo con Pestañas
        tab1, tab2, tab3 = st.tabs(["📋 1. Datos Generales", "⚙️ 2. Parámetros", "📝 3. Diagnóstico Final"])
        
        with tab1:
            st.markdown("### Configuración de la Intervención")
            tipo_plan = st.selectbox("🛠️ Tipo de Plan / Orden:", ["Inspección", "PM03"] if "CD" in tag_sel else ["Inspección", "P1", "P2", "P3", "PM03"])
            
            c1, c2, c3, c4 = st.columns(4)
            modelo, numero_serie, area, ubicacion = c1.text_input("Modelo", mod_d), c2.text_input("N° Serie", ser_d), c3.text_input("Área", area_d), c4.text_input("Ubicación", ubi_d)

            c5, c6, c7, c8 = st.columns(4)
            fecha = c5.text_input("Fecha Ejecución", "23 de febrero de 2026")
            tec1 = c6.text_input("Técnico Principal", key="input_tec1")
            tec2 = c7.text_input("Técnico Asistente", key="input_tec2")
            cli_cont = c8.text_input("Contacto Cliente", key="input_cliente")

        with tab2:
            st.markdown("### Mediciones del Equipo")
            c9, c10, c11, c12, c13, c14 = st.columns(6)
            h_m = c9.number_input("Horas Marcha Totales", step=1.0, key="input_h_marcha")
            h_c = c10.number_input("Horas en Carga", step=1.0, key="input_h_carga")
            unidad_p = c11.selectbox("Unidad de Presión", ["bar", "psi"])
            p_c = c12.number_input("P. Carga", step=0.1, key="input_p_carga")
            p_d = c13.number_input("P. Descarga", step=0.1, key="input_p_descarga")
            t_salida = c14.number_input("Temp Salida (°C)", step=0.1, key="input_temp")

        with tab3:
            st.markdown("### Evaluación y Conclusiones")
            est_eq = st.radio("Estado de Devolución del Activo:", ["Operativo", "Fuera de servicio"], key="input_estado_eq", horizontal=True)
            est_ent = st.text_area("Descripción Condición Final:", key="input_estado", height=100)
            reco = st.text_area("Recomendaciones / Acciones Pendientes:", key="input_reco", height=100)
            
            st.markdown("<br>", unsafe_allow_html=True)
            # Botón de Guardado Prominente
            if st.button("🚀 Generar, Sincronizar y Guardar Reporte Oficial", type="primary", use_container_width=True):
                with st.spinner('Procesando datos y contactando con la base de datos...'):
                    try:
                        if "CD" in tag_sel: file_plantilla = "plantilla/secadorfueradeservicio.docx" if est_eq == "Fuera de servicio" else "plantilla/inspeccionsecador.docx"
                        else:
                            if est_eq == "Fuera de servicio": file_plantilla = "plantilla/fueradeservicio.docx"
                            elif tipo_plan == "P1": file_plantilla = "plantilla/p1.docx"
                            elif tipo_plan == "P2": file_plantilla = "plantilla/p2.docx"
                            elif tipo_plan == "P3": file_plantilla = "plantilla/p3.docx"
                            else: file_plantilla = "plantilla/inspeccion.docx"
                            
                        doc = DocxTemplate(file_plantilla)
                        context = {"tipo_intervencion": tipo_plan, "modelo": modelo, "tag": tag_sel, "area": area, "ubicacion": ubicacion, "cliente_contacto": cli_cont, "p_carga": f"{p_c} {unidad_p}", "p_descarga": f"{p_d} {unidad_p}", "temp_salida": t_salida, "horas_marcha": int(h_m), "horas_carga": int(h_c), "tecnico_1": tec1, "tecnico_2": tec2, "estado_equipo": est_eq, "estado_entrega": est_ent, "recomendaciones": reco, "serie": numero_serie, "tipo_orden": tipo_plan.upper(), "fecha": fecha, "equipo_modelo": modelo}
                        doc.render(context)
                        
                        nombre_archivo = f"Informe_{tipo_plan}_{tag_sel}_{fecha.replace(' ','_')}.docx"
                        folder = os.path.join("Historial_Informes", tag_sel)
                        os.makedirs(folder, exist_ok=True)
                        ruta = os.path.join(folder, nombre_archivo)
                        doc.save(ruta)
                        
                        ruta_pdf_gen = convertir_a_pdf(ruta)
                        tupla_db = (tag_sel, modelo, numero_serie, area, ubicacion, fecha, cli_cont, tec1, tec2, t_salida, f"{p_c} {unidad_p}", f"{p_d} {unidad_p}", h_m, h_c, est_ent, tipo_plan, reco, est_eq, ruta, st.session_state.usuario_actual)
                        guardar_registro(tupla_db)
                        
                        st.success(f"✅ ¡Operación Exitosa! Reporte '{nombre_archivo}' emitido correctamente.")
                        st.info(sincronizar_con_nube(tag_sel, tipo_plan)[1])
                        
                        c_d1, c_d2 = st.columns(2)
                        with c_d1:
                            with open(ruta, "rb") as f: st.download_button("📄 Obtener Original (Word)", f, file_name=nombre_archivo, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
                        with c_d2:
                            if ruta_pdf_gen:
                                with open(ruta_pdf_gen, "rb") as f_pdf: st.download_button("📕 Obtener Oficial (PDF)", f_pdf, file_name=nombre_archivo.replace(".docx", ".pdf"), mime="application/pdf", use_container_width=True)
                            else: st.button("📕 PDF (En proceso / Revisar nube)", disabled=True, use_container_width=True)
                    except Exception as e: st.error(f"Error sistémico generando reporte: {e}")

        # Renderizado del Historial Local
        st.markdown("<br><hr>", unsafe_allow_html=True)
        st.markdown("### 📋 Trazabilidad Histórica de Intervenciones")
        df_hist = obtener_todo_el_historial(tag_sel)
        if not df_hist.empty: st.dataframe(df_hist, use_container_width=True)