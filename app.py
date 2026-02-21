import streamlit as st
from docxtpl import DocxTemplate
import io
import os
import sqlite3
import subprocess

# --- 1. MÓDULO DE BASE DE DATOS LOCAL ---
def init_db():
    conn = sqlite3.connect("historial_equipos.db")
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS intervenciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tag TEXT,
            modelo TEXT,
            fecha TEXT,
            cliente_contacto TEXT,
            temp_salida REAL,
            estado_entrega TEXT,
            tipo_intervencion TEXT
        )
    ''')
    
    # Actualización automática de la base de datos para los nuevos campos
    columnas_nuevas = ["ruta_archivo", "numero_serie", "area_especifica", "sector"]
    for col in columnas_nuevas:
        try:
            cursor.execute(f'ALTER TABLE intervenciones ADD COLUMN {col} TEXT')
        except sqlite3.OperationalError:
            pass # Si la columna ya existe, sigue adelante
            
    conn.commit()
    conn.close()

def guardar_registro(tag, modelo, serie, area, sector, fecha, cliente, temp, estado, tipo, ruta_archivo):
    conn = sqlite3.connect("historial_equipos.db")
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO intervenciones (tag, modelo, numero_serie, area_especifica, sector, fecha, cliente_contacto, temp_salida, estado_entrega, tipo_intervencion, ruta_archivo)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (tag, modelo, serie, area, sector, fecha, cliente, temp, estado, tipo, ruta_archivo))
    conn.commit()
    conn.close()

def buscar_ultimo_registro(tag):
    conn = sqlite3.connect("historial_equipos.db")
    cursor = conn.cursor()
    cursor.execute('''
        SELECT fecha, cliente_contacto, temp_salida, estado_entrega, tipo_intervencion 
        FROM intervenciones 
        WHERE tag = ? 
        ORDER BY id DESC LIMIT 1
    ''', (tag,))
    resultado = cursor.fetchone()
    conn.close()
    return resultado

# --- 2. MÓDULO DE NUBE ---
def sincronizar_con_nube(tag, tipo_plan):
    try:
        subprocess.run(["git", "add", "."], check=True, capture_output=True, text=True)
        mensaje = f"InforGem: Auto-guardado de {tipo_plan} para el equipo {tag}"
        subprocess.run(["git", "commit", "-m", mensaje], check=True, capture_output=True, text=True)
        subprocess.run(["git", "push"], check=True, capture_output=True, text=True)
        return True, "☁️ ¡Respaldo total en la nube (Word + Base de Datos) exitoso!"
    except subprocess.CalledProcessError as e:
        error_msg = e.stderr.strip() if e.stderr else "Sin cambios nuevos para subir."
        return False, f"⚠️ Aviso de Nube: {error_msg}"
    except FileNotFoundError:
        return False, "⚠️ Aviso: No se detectó Git."

# --- 3. INICIO DE LA APLICACIÓN VISUAL ---
init_db()

st.set_page_config(page_title="InforGem Generador", layout="wide", page_icon="⚙️")

if 'input_fecha' not in st.session_state:
    st.session_state.input_fecha = "21 de febrero de 2026"
    st.session_state.input_cliente = "Lorena Rojas"
    st.session_state.input_temp = 66.5
    st.session_state.input_estado = "El equipo se encuentra funcionando en óptimas condiciones..."
    st.session_state.fecha_ultima = ""
    st.session_state.tipo_ultimo = ""

st.title("⚙️ Sistema de Mantenimiento InforGem")
st.markdown("---")

# --- INVENTARIO MAESTRO (Actualizado) ---
# Formato: "TAG": ["Modelo", "Serie", "Área Específica", "Sector"]
inventario_equipos = {
    "70-GC-013": ["GA 132", "AIF095296", "descarga acido", "área húmeda"],
    "70-GC-014": ["GA 132", "AIF095297", "descarga acido", "área húmeda"],
    "050-GD-001": ["GA 45", "API542705", "planta sx", "área húmeda"],
    "050-GD-002": ["GA 45", "API542706", "planta sx", "área húmeda"],
    "050-GC-003": ["ZT 37", "API791692", "planta sx", "área húmeda"],
    "050-GC-004": ["ZT 37", "API791693", "planta sx", "área húmeda"],
    "050-CD-001": ["CD 80+", "API095825", "planta sx", "área húmeda"],
    "050-CD-002": ["CD 80+", "API095826", "planta sx", "área húmeda"],
    "050-GC-015": ["GA 30", "API501440", "planta borra", "área húmeda"],
    "65-GC-011": ["GA 250", "APF253581", "patio estanques", "área húmeda"],
    "65-GC-009": ["GA 250", "APF253608", "patio estanques", "área húmeda"],
    "65-GD-011": ["CD 630", "WXF300015", "patio estanques", "área húmeda"],
    "65-GD-012": ["CD 630", "WXF300016", "patio estanques", "área húmeda"],
    "35-GC-006": ["GA 250", "AIF095420", "chancado secundario", "área seca"],
    "35-GC-007": ["GA 250", "AIF095421", "chancado secundario", "área seca"],
    "35-GC-008": ["GA 250", "AIF095302", "chancado secundario", "área seca"],
    "20-GC-004": ["GA 37", "AII390776", "mina", "mina"],
    "20-GC-001": ["GA 75", "AII482673", "truck shop", "mina"],
    "20-GC-002": ["GA 75", "AII482674", "truck shop", "mina"],
    "20-GC-003": ["GA 90", "AIF095178", "truck shop", "mina"],
    "TALLER-01": ["GA18", "API335343", "taller", "área seca"]
}

# --- SECCIÓN 1: BÚSQUEDA ---
col_busqueda, col_plan = st.columns(2)

with col_busqueda:
    lista_tags = list(inventario_equipos.keys())
    tag_seleccionado = st.selectbox("🔍 Seleccionar TAG del Equipo:", lista_tags)
    
    # Extraemos los datos predeterminados del inventario
    datos_eq = inventario_equipos[tag_seleccionado]
    modelo_default = datos_eq[0]
    serie_default = datos_eq[1]
    area_default = datos_eq[2]
    sector_default = datos_eq[3]

    if st.button("Buscar Historial en Base de Datos"):
        registro = buscar_ultimo_registro(tag_seleccionado)
        if registro:
            st.session_state.fecha_ultima = registro[0]
            st.session_state.input_cliente = registro[1]
            try:
                st.session_state.input_temp = float(registro[2])
            except (ValueError, TypeError):
                st.session_state.input_temp = 0.0
            st.session_state.input_estado = registro[3]
            st.session_state.tipo_ultimo = registro[4]
            st.success("✅ ¡Historial encontrado! Se recuperaron los parámetros de la última visita.")
        else:
            st.session_state.fecha_ultima = ""
            st.session_state.tipo_ultimo = ""
            st.warning("No hay registros previos para este equipo. ¡Es su primera vez en InforGem!")

with col_plan:
    tipo_plan = st.selectbox("🛠️ Tipo de Intervención a realizar HOY:", ["Inspección", "P1", "P2", "P3", "P4"])

if st.session_state.fecha_ultima != "":
    st.info(f"📌 **Referencia Histórica:** El último trabajo realizado fue un(a) **{st.session_state.tipo_ultimo}**, el **{st.session_state.fecha_ultima}**.")

st.markdown("---")

# --- SECCIÓN 2: DATOS DEL EQUIPO (Editables) ---
st.subheader("📋 Datos del Activo")
col_eq1, col_eq2, col_eq3, col_eq4 = st.columns(4)

with col_eq1:
    modelo = st.text_input("Modelo:", value=modelo_default)
with col_eq2:
    numero_serie = st.text_input("N° de Serie:", value=serie_default)
with col_eq3:
    area_especifica = st.text_input("Área Específica:", value=area_default)
with col_eq4:
    sector = st.text_input("Sector:", value=sector_default)

st.markdown("---")

# --- SECCIÓN 3: DATOS DEL REPORTE NUEVO ---
st.subheader(f"🔧 Datos de la Intervención")
col1, col2 = st.columns(2)

with col1:
    fecha = st.text_input("Fecha de Intervención (Actual):", key="input_fecha")
    cliente_contacto = st.text_input("Contacto Cliente:", key="input_cliente")

with col2:
    temp_salida = st.number_input("Temperatura de Salida (°C):", step=0.1, key="input_temp")
    estado_entrega = st.text_area("Estado de Entrega:", key="input_estado")

st.markdown("---")

# --- SECCIÓN 4: GENERACIÓN Y GUARDADO ---
if st.button(f"Generar, Guardar y Registrar Word de {tipo_plan}", type="primary"):
    try:
        if tipo_plan == "Inspección":
            plantilla_path = "plantilla/inspeccion.docx"
        else:
            plantilla_path = "plantilla/inspeccion.docx" 
            
        doc = DocxTemplate(plantilla_path)
        
        # INYECTAMOS TODOS LOS DATOS NUEVOS A LA PLANTILLA
        context = {
            "tag": tag_seleccionado, 
            "modelo": modelo, 
            "numero_serie": numero_serie,
            "area_especifica": area_especifica,
            "sector": sector,
            "fecha": fecha,
            "cliente_contacto": cliente_contacto, 
            "temp_salida": temp_salida,
            "estado_entrega": estado_entrega, 
            "tipo_intervencion": tipo_plan
        }
        doc.render(context)

        # Guardado Local
        carpeta_equipo = os.path.join("Historial_Informes", tag_seleccionado)
        os.makedirs(carpeta_equipo, exist_ok=True)
        nombre_archivo = f"Informe_{tipo_plan}_{tag_seleccionado}.docx"
        ruta_completa = os.path.join(carpeta_equipo, nombre_archivo)
        doc.save(ruta_completa)

        # Guardar todo en la BD
        guardar_registro(tag_seleccionado, modelo, numero_serie, area_especifica, sector, fecha, cliente_contacto, temp_salida, estado_entrega, tipo_plan, ruta_completa)
        st.success(f"✅ ¡Word creado localmente y BD actualizada!")

        # Módulo de Nube
        exito_nube, mensaje_nube = sincronizar_con_nube(tag_seleccionado, tipo_plan)
        if exito_nube:
            st.success(mensaje_nube)
        else:
            st.warning(mensaje_nube)

        # Preparar Descarga
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        st.download_button(
            label=f"⬇️ Descargar Copia Manual ({tipo_plan})",
            data=buffer,
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        st.error(f"⚠️ Error. Detalle: {e}")