import streamlit as st
from docxtpl import DocxTemplate
import io
import os
import sqlite3

# --- CONFIGURACIÓN DE LA BASE DE DATOS ---
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
    conn.commit()
    conn.close()

def guardar_registro(tag, modelo, fecha, cliente, temp, estado, tipo):
    conn = sqlite3.connect("historial_equipos.db")
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO intervenciones (tag, modelo, fecha, cliente_contacto, temp_salida, estado_entrega, tipo_intervencion)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    ''', (tag, modelo, fecha, cliente, temp, estado, tipo))
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

# Iniciar la base de datos
init_db()

# --- CONFIGURACIÓN DE PÁGINA Y MEMORIA SEGURA ---
st.set_page_config(page_title="InforGem Generador", page_icon="⚙️")

# Inicializar las llaves de memoria solo si no existen (evita el bloqueo y la pantalla negra)
if 'input_fecha' not in st.session_state:
    st.session_state.input_fecha = "21 de febrero de 2026"
    st.session_state.input_cliente = "Lorena Rojas"
    st.session_state.input_temp = 66.5
    st.session_state.input_estado = "El equipo se encuentra funcionando en óptimas condiciones..."
    st.session_state.fecha_ultima = ""
    st.session_state.tipo_ultimo = ""

st.title("⚙️ Sistema de Mantenimiento InforGem")
st.markdown("---")

# Inventario de prueba
inventario_equipos = {
    "25-GC-007": "GA250",
    "0505-GC-015": "GA30",
    "10-GC-002": "GA90",
    "99-GC-100": "ZR400"
}

# --- SECCIÓN 1: BÚSQUEDA Y TIPO DE PLAN ---
col_busqueda, col_plan = st.columns(2)

with col_busqueda:
    lista_tags = list(inventario_equipos.keys())
    tag_seleccionado = st.selectbox("🔍 Seleccionar TAG del Equipo:", lista_tags)
    modelo_automatico = inventario_equipos[tag_seleccionado]

    if st.button("Buscar Historial en Base de Datos"):
        registro = buscar_ultimo_registro(tag_seleccionado)
        if registro:
            st.session_state.fecha_ultima = registro[0]
            st.session_state.input_cliente = registro[1]
            
            # Seguro contra errores de conversión de números (evita que la pantalla colapse)
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

st.markdown("---")

# --- AVISO VISUAL DE LA ÚLTIMA INTERVENCIÓN ---
if st.session_state.fecha_ultima != "":
    st.info(f"📌 **Referencia Histórica:** El último trabajo realizado en este equipo fue un(a) **{st.session_state.tipo_ultimo}**, el **{st.session_state.fecha_ultima}**.")

# --- SECCIÓN 2: DATOS DEL REPORTE NUEVO ---
st.subheader(f"Datos para el Reporte Nuevo ({tag_seleccionado})")
col1, col2 = st.columns(2)

with col1:
    modelo = st.text_input("Modelo del Equipo:", value=modelo_automatico)
    fecha = st.text_input("Fecha de Intervención (Actual):", key="input_fecha")
    cliente_contacto = st.text_input("Contacto Cliente:", key="input_cliente")

with col2:
    temp_salida = st.number_input("Temperatura de Salida (°C):", step=0.1, key="input_temp")
    estado_entrega = st.text_area("Estado de Entrega:", key="input_estado")

st.markdown("---")

# --- SECCIÓN 3: GENERACIÓN Y GUARDADO ---
if st.button(f"Generar, Guardar y Registrar Word de {tipo_plan}", type="primary"):
    try:
        # Selección de plantilla
        if tipo_plan == "Inspección":
            plantilla_path = "plantilla/inspeccion.docx"
        else:
            plantilla_path = "plantilla/inspeccion.docx" 
            
        doc = DocxTemplate(plantilla_path)

        context = {
            "tag": tag_seleccionado,
            "modelo": modelo,
            "fecha": fecha,
            "cliente_contacto": cliente_contacto,
            "temp_salida": temp_salida,
            "estado_entrega": estado_entrega,
            "tipo_intervencion": tipo_plan
        }
        doc.render(context)

        # Creación de carpetas y guardado físico
        carpeta_principal = "Historial_Informes"
        carpeta_equipo = os.path.join(carpeta_principal, tag_seleccionado)
        os.makedirs(carpeta_equipo, exist_ok=True)
        
        nombre_archivo = f"Informe_{tipo_plan}_{tag_seleccionado}.docx"
        ruta_completa = os.path.join(carpeta_equipo, nombre_archivo)
        doc.save(ruta_completa)

        # Guardado en base de datos SQLite
        guardar_registro(tag_seleccionado, modelo, fecha, cliente_contacto, temp_salida, estado_entrega, tipo_plan)

        # Descarga
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        st.success(f"✅ ¡Todo listo! Registro en BD exitoso y Word guardado en: {ruta_completa}")
        st.download_button(
            label=f"⬇️ Descargar Copia Manual ({tipo_plan})",
            data=buffer,
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        st.error(f"⚠️ Error al procesar la plantilla. Detalle: {e}")