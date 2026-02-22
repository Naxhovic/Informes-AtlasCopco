import streamlit as st
from docxtpl import DocxTemplate
import io
import os
import sqlite3
import subprocess
import pandas as pd

# --- CONFIGURACIÓN DE USUARIOS ---
USUARIOS = {
    "Ignacio Morales": "spence2026",
    "Emian": "spence2026",
    "Ignacio Veas": "spence2026",
    "admin": "admin123"
}

# --- 1. MÓDULO DE BASE DE DATOS LOCAL CON AUTO-MIGRACIÓN ---
def init_db():
    conn = sqlite3.connect("historial_equipos.db")
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS intervenciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tag TEXT, modelo TEXT, numero_serie TEXT, area TEXT, ubicacion TEXT,
            fecha TEXT, cliente_contacto TEXT, tecnico_1 TEXT, tecnico_2 TEXT,
            temp_salida REAL, p_carga TEXT, p_descarga TEXT,
            horas_marcha REAL, horas_carga REAL,
            estado_entrega TEXT, tipo_intervencion TEXT, recomendaciones TEXT, 
            estado_equipo TEXT, ruta_archivo TEXT, generado_por TEXT
        )
    ''')
    
    cursor.execute("PRAGMA table_info(intervenciones)")
    columnas_actuales = [info[1] for info in cursor.fetchall()]
    
    if "estado_equipo" not in columnas_actuales:
        cursor.execute("ALTER TABLE intervenciones ADD COLUMN estado_equipo TEXT DEFAULT 'Operativo'")
    if "recomendaciones" not in columnas_actuales:
        cursor.execute("ALTER TABLE intervenciones ADD COLUMN recomendaciones TEXT")
    # NUEVA COLUMNA DE AUDITORÍA
    if "generado_por" not in columnas_actuales:
        cursor.execute("ALTER TABLE intervenciones ADD COLUMN generado_por TEXT DEFAULT 'Desconocido'")
        
    conn.commit()
    conn.close()

def obtener_estados_actuales():
    try:
        conn = sqlite3.connect("historial_equipos.db")
        cursor = conn.cursor()
        cursor.execute('''
            SELECT tag, estado_equipo FROM intervenciones 
            WHERE id IN (SELECT MAX(id) FROM intervenciones GROUP BY tag)
        ''')
        estados = {row[0]: row[1] for row in cursor.fetchall()}
        conn.close()
        return estados
    except: return {}

def guardar_registro(tag, mod, serie, area, ubi, fecha, cli, tec1, tec2, temp, p_c, p_d, h_m, h_c, est, tipo, reco, est_eq, ruta, usuario):
    conn = sqlite3.connect("historial_equipos.db")
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO intervenciones 
        (tag, modelo, numero_serie, area, ubicacion, fecha, cliente_contacto, tecnico_1, tecnico_2, 
        temp_salida, p_carga, p_descarga, horas_marcha, horas_carga, estado_entrega, tipo_intervencion, recomendaciones, estado_equipo, ruta_archivo, generado_por)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (tag, mod, serie, area, ubi, fecha, cli, tec1, tec2, temp, p_c, p_d, h_m, h_c, est, tipo, reco, est_eq, ruta, usuario))
    conn.commit()
    conn.close()

def buscar_ultimo_registro(tag):
    conn = sqlite3.connect("historial_equipos.db")
    cursor = conn.cursor()
    cursor.execute('''
        SELECT fecha, cliente_contacto, temp_salida, estado_entrega, tipo_intervencion, 
               tecnico_1, tecnico_2, p_carga, p_descarga, horas_marcha, horas_carga, recomendaciones, estado_equipo
        FROM intervenciones WHERE tag = ? ORDER BY id DESC LIMIT 1
    ''', (tag,))
    resultado = cursor.fetchone()
    conn.close()
    return resultado

def obtener_todo_el_historial(tag):
    conn = sqlite3.connect("historial_equipos.db")
    query = """
        SELECT fecha, tipo_intervencion, estado_equipo, generado_por, horas_marcha, horas_carga, 
               p_carga, p_descarga, temp_salida 
        FROM intervenciones WHERE tag = ? ORDER BY id DESC
    """
    df = pd.read_sql_query(query, conn, params=(tag,))
    conn.close()
    return df

# --- 2. MÓDULO DE NUBE ---
def sincronizar_con_nube(tag, tipo_plan):
    try:
        subprocess.run(["git", "add", "."], check=True, capture_output=True, text=True)
        subprocess.run(["git", "commit", "-m", f"Reporte: {tipo_plan} - {tag}"], check=True)
        subprocess.run(["git", "push"], check=True)
        return True, "☁️ Sincronización con GitHub exitosa."
    except Exception:
        return False, "⚠️ Pendiente de subir a GitHub."

# --- 3. CONFIGURACIÓN DE INTERFAZ Y SESIÓN ---
init_db()
st.set_page_config(page_title="InforGem Generador", layout="wide", page_icon="⚙️")

# Estado de autenticación
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'usuario_actual' not in st.session_state: st.session_state.usuario_actual = ""

# Memoria
if 'input_cliente' not in st.session_state: st.session_state.input_cliente = "Lorena Rojas"
if 'input_tec1' not in st.session_state: st.session_state.input_tec1 = "Ignacio Morales"
if 'input_tec2' not in st.session_state: st.session_state.input_tec2 = "emian Sanchez"
if 'input_h_marcha' not in st.session_state: st.session_state.input_h_marcha = 0.0
if 'input_h_carga' not in st.session_state: st.session_state.input_h_carga = 0.0
if 'input_temp' not in st.session_state: st.session_state.input_temp = 70.0
if 'input_p_carga' not in st.session_state: st.session_state.input_p_carga = 7.0
if 'input_p_descarga' not in st.session_state: st.session_state.input_p_descarga = 7.5
if 'input_estado' not in st.session_state: st.session_state.input_estado = ""
if 'input_reco' not in st.session_state: st.session_state.input_reco = ""
if 'input_estado_eq' not in st.session_state: st.session_state.input_estado_eq = "Operativo"

# --- LOGIN ---
if not st.session_state.logged_in:
    st.markdown("<h1 style='text-align: center;'>🔒 Acceso Sistema InforGem</h1>", unsafe_allow_html=True)
    st.markdown("---")
    
    col_l1, col_l2, col_l3 = st.columns([1, 2, 1])
    with col_l2:
        with st.form("form_login"):
            st.subheader("Ingresa tus credenciales")
            usuario_ingresado = st.text_input("Usuario (Ej: ignacio)").lower()
            password_ingresada = st.text_input("Contraseña", type="password")
            submit_login = st.form_submit_button("Ingresar a la Plataforma", type="primary", use_container_width=True)
            
            if submit_login:
                if usuario_ingresado in USUARIOS and USUARIOS[usuario_ingresado] == password_ingresada:
                    st.session_state.logged_in = True
                    st.session_state.usuario_actual = usuario_ingresado
                    st.success("✅ Acceso concedido.")
                    st.rerun()
                else:
                    st.error("❌ Usuario o contraseña incorrectos.")

# --- APP PRINCIPAL ---
else:
    with st.sidebar:
        st.success(f"👤 Conectado como:\n**{st.session_state.usuario_actual.capitalize()}**")
        if st.button("🚪 Cerrar Sesión", use_container_width=True):
            st.session_state.logged_in = False
            st.rerun()

    st.title("⚙️ Sistema de Mantenimiento InforGem")
    st.markdown("---")

    inventario_equipos = {
        "20-GC-001": ["GA 75", "AII482673", "truck shop", "mina"],
        "20-GC-002": ["GA 75", "AII482674", "truck shop", "mina"],
        "20-GC-003": ["GA 90", "AIF095178", "truck shop", "mina"],
        "20-GC-004": ["GA 37", "AII390776", "mina", "mina"],
        "35-GC-006": ["GA 250", "AIF095420", "chancado secundario", "área seca"],
        "35-GC-007": ["GA 250", "AIF095421", "chancado secundario", "área seca"],
        "35-GC-008": ["GA 250", "AIF095302", "chancado secundario", "área seca"],
        "50-GC-001": ["GA 45", "API542705", "planta SX", "área húmeda"],
        "50-GC-002": ["GA 45", "API542706", "planta SX", "área húmeda"],
        "50-GC-003": ["ZT 37", "API791692", "planta SX", "área húmeda"],
        "50-GC-004": ["ZT 37", "API791693", "planta SX", "área húmeda"],
        "50-CD-001": ["CD 80+", "API095825", "planta SX", "área húmeda"],
        "50-CD-002": ["CD 80+", "API095826", "planta SX", "área húmeda"],
        "55-GC-015": ["GA 30", "API501440", "planta borra", "área húmeda"],
        "65-GC-009": ["GA 250", "APF253608", "patio estanques", "área húmeda"],
        "65-GC-011": ["GA 250", "APF253581", "patio estanques", "área húmeda"],
        "65-CD-011": ["CD 630", "WXF300015", "patio estanques", "área húmeda"],
        "65-CD-012": ["CD 630", "WXF300016", "patio estanques", "área húmeda"],
        "70-GC-013": ["GA 132", "AIF095296", "descarga acido", "área húmeda"],
        "70-GC-014": ["GA 132", "AIF095297", "descarga acido", "área húmeda"],
        "TALLER-01": ["GA18", "API335343", "taller", "laboratorio"]
    }

    estados_db = obtener_estados_actuales()
    col_busqueda, col_plan = st.columns(2)
    with col_busqueda:
        def format_func(tag):
            estado = estados_db.get(tag, "Operativo")
            return f"{'🟢' if estado == 'Operativo' else '🔴'} {tag}"

        tag_sel = st.selectbox("🔍 TAG del Equipo:", list(inventario_equipos.keys()), format_func=format_func)
        mod_d, ser_d, area_d, ubi_d = inventario_equipos[tag_sel]
        
        if st.button("Buscar Historial"):
            reg = buscar_ultimo_registro(tag_sel)
            if reg:
                st.session_state.input_cliente = reg[1]
                st.session_state.input_tec1 = reg[5]
                st.session_state.input_tec2 = reg[6]
                st.session_state.input_estado = reg[3]
                st.session_state.input_reco = reg[11] if reg[11] else ""
                st.session_state.input_estado_eq = reg[12] if reg[12] else "Operativo"
                st.session_state.input_temp = float(reg[2])
                st.session_state.input_h_marcha = float(reg[9]) if reg[9] else 0.0
                st.session_state.input_h_carga = float(reg[10]) if reg[10] else 0.0
                try: st.session_state.input_p_carga = float(str(reg[7]).split()[0])
                except: st.session_state.input_p_carga = 7.0
                try: st.session_state.input_p_descarga = float(str(reg[8]).split()[0])
                except: st.session_state.input_p_descarga = 7.5
                st.success("✅ Datos cargados.")
                st.rerun()

    with col_plan:
        if "CD" in tag_sel:
            opciones_plan = ["Inspección", "PM03"]
        else:
            opciones_plan = ["Inspección", "P1", "P2", "P3", "PM03"]
            
        tipo_plan = st.selectbox("🛠️ Tipo Intervención:", opciones_plan)

    st.markdown("---")

    c1, c2, c3, c4 = st.columns(4)
    modelo = c1.text_input("Modelo:", value=mod_d)
    numero_serie = c2.text_input("N° Serie:", value=ser_d)
    area = c3.text_input("Área:", value=area_d)
    ubicacion = c4.text_input("Ubicación:", value=ubi_d)

    c5, c6, c7, c8 = st.columns(4)
    fecha = c5.text_input("Fecha:", value="21 de febrero de 2026")
    tecnico_1 = c6.text_input("Técnico 1:", key="input_tec1")
    tecnico_2 = c7.text_input("Técnico 2:", key="input_tec2")
    cliente_contacto = c8.text_input("Contacto Cliente:", key="input_cliente")

    st.subheader("📊 Parámetros Técnicos")
    c9, c10, c11, c12, c13, c14 = st.columns(6)
    horas_marcha = c9.number_input("Horas Marcha:", step=1.0, key="input_h_marcha")
    horas_carga = c10.number_input("Horas Carga:", step=1.0, key="input_h_carga")
    unidad_p = c11.selectbox("Unidad:", ["bar", "psi"])
    p_carga_val = c12.number_input(f"P. Carga:", step=0.1, key="input_p_carga")
    p_descarga_val = c13.number_input(f"P. Descarga:", step=0.1, key="input_p_descarga")
    temp_salida = c14.number_input("Temp Salida (°C):", step=0.1, key="input_temp")

    st.subheader("📝 Notas y Estado Final")
    col_est1, col_est2 = st.columns([1, 2])
    with col_est1:
        estado_equipo = st.radio("Estado final del equipo:", ["Operativo", "Fuera de servicio"], key="input_estado_eq", horizontal=True)
    with col_est2:
        estado_entrega = st.text_area("Estado de Entrega:", key="input_estado")

    recomendaciones = st.text_area("Nota Técnica / Recomendaciones:", key="input_reco")

    if st.button("🚀 Generar Reporte Industrial", type="primary"):
        try:
            if "CD" in tag_sel:
                if estado_equipo == "Fuera de servicio": file_plantilla = "plantilla/secadorfueradeservicio.docx"
                else: file_plantilla = "plantilla/inspeccionsecador.docx"
            else:
                if estado_equipo == "Fuera de servicio": file_plantilla = "plantilla/fueradeservicio.docx"
                elif tipo_plan == "P1": file_plantilla = "plantilla/p1.docx"
                elif tipo_plan == "P2": file_plantilla = "plantilla/p2.docx"
                elif tipo_plan == "P3": file_plantilla = "plantilla/p3.docx"
                else: file_plantilla = "plantilla/inspeccion.docx"
                
            doc = DocxTemplate(file_plantilla)
            context = {
                "tipo_intervencion": tipo_plan, "modelo": modelo, "tag": tag_sel,
                "area": area, "ubicacion": ubicacion, "cliente_contacto": cliente_contacto,
                "p_carga": f"{p_carga_val} {unidad_p}", "p_descarga": f"{p_descarga_val} {unidad_p}",
                "temp_salida": temp_salida, "horas_marcha": int(horas_marcha), "horas_carga": int(horas_carga),
                "tecnico_1": tecnico_1, "tecnico_2": tecnico_2, "estado_equipo": estado_equipo,
                "estado_entrega": estado_entrega, "recomendaciones": recomendaciones,
                "serie": numero_serie, "tipo_orden": tipo_plan.upper(), "fecha": fecha, "equipo_modelo": modelo
            }
            doc.render(context)
            
            nombre_archivo = f"Informe_{tipo_plan}_{tag_sel}_{fecha.replace(' ','_')}.docx"
            folder = os.path.join("Historial_Informes", tag_sel)
            os.makedirs(folder, exist_ok=True)
            ruta = os.path.join(folder, nombre_archivo)
            doc.save(ruta)
            
            # --- GUARDADO EN DB INCLUYENDO EL USUARIO LOGUEADO ---
            guardar_registro(tag_sel, modelo, numero_serie, area, ubicacion, fecha, cliente_contacto, 
                             tecnico_1, tecnico_2, temp_salida, f"{p_carga_val} {unidad_p}", f"{p_descarga_val} {unidad_p}", 
                             horas_marcha, horas_carga, estado_entrega, tipo_plan, recomendaciones, estado_equipo, ruta, 
                             st.session_state.usuario_actual) # ¡Aquí pasa el usuario real!
            
            st.success(f"✅ Reporte generado utilizando plantilla: {file_plantilla.split('/')[-1]}")
            st.info(sincronizar_con_nube(tag_sel, tipo_plan)[1])
            
            with st.expander("👁️ Vista Previa de Datos del Reporte", expanded=True):
                st.markdown(f"**📍 Equipo:** {modelo} ({tag_sel}) | **N° Serie:** {numero_serie}")
                st.markdown(f"**🛠️ Tipo de Orden:** {tipo_plan.upper()} | **Fecha:** {fecha}")
                st.markdown(f"**👨‍🔧 Técnicos Registrados:** {tecnico_1} y {tecnico_2}")
                st.markdown(f"**✍️ Documento Creado por:** {st.session_state.usuario_actual.capitalize()}") # Mostrar en pantalla quién lo hizo
                if estado_equipo == "Operativo": st.success(f"**Estado Final:** {estado_equipo}")
                else: st.error(f"**Estado Final:** {estado_equipo}")
                st.info(f"**Comentarios de Entrega:**\n{estado_entrega}")
                if recomendaciones: st.warning(f"**Nota Técnica:**\n{recomendaciones}")
            
            with open(ruta, "rb") as file:
                st.download_button(
                    label="⬇️ Descargar Reporte",
                    data=file,
                    file_name=nombre_archivo,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
        except Exception as e:
            st.error(f"Error: {e}")

    # --- TABLA HISTÓRICA CON TEMA OSCURO Y NUEVA COLUMNA DE USUARIO ---
    st.markdown("---")
    df_hist = obtener_todo_el_historial(tag_sel)
    if not df_hist.empty:
        # Reordenar/renombrar columnas para que sea más legible en la tabla
        df_hist = df_hist.rename(columns={"generado_por": "Cuenta Usuario"})
        st.dataframe(df_hist, use_container_width=True)