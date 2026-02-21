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
            tag TEXT, modelo TEXT, numero_serie TEXT, area TEXT, ubicacion TEXT,
            fecha TEXT, cliente_contacto TEXT, tecnico_1 TEXT, tecnico_2 TEXT,
            temp_salida REAL, p_carga REAL, p_descarga REAL,
            estado_entrega TEXT, tipo_intervencion TEXT, ruta_archivo TEXT
        )
    ''')
    conn.commit()
    conn.close()

def guardar_registro(tag, mod, serie, area, ubi, fecha, cli, tec1, tec2, temp, p_c, p_d, est, tipo, ruta):
    conn = sqlite3.connect("historial_equipos.db")
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO intervenciones 
        (tag, modelo, numero_serie, area, ubicacion, fecha, cliente_contacto, tecnico_1, tecnico_2, temp_salida, p_carga, p_descarga, estado_entrega, tipo_intervencion, ruta_archivo)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (tag, mod, serie, area, ubi, fecha, cli, tec1, tec2, temp, p_c, p_d, est, tipo, ruta))
    conn.commit()
    conn.close()

def buscar_ultimo_registro(tag):
    conn = sqlite3.connect("historial_equipos.db")
    cursor = conn.cursor()
    cursor.execute('''
        SELECT fecha, cliente_contacto, temp_salida, estado_entrega, tipo_intervencion, tecnico_1, tecnico_2, p_carga, p_descarga
        FROM intervenciones WHERE tag = ? ORDER BY id DESC LIMIT 1
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
        return True, "☁️ ¡Respaldo total en la nube exitoso!"
    except subprocess.CalledProcessError:
        return False, "⚠️ Aviso de Nube: Pendiente de sincronización (Posible conflicto resuelto con git pull)."

# --- 3. INICIO DE LA APLICACIÓN ---
init_db()
st.set_page_config(page_title="InforGem Generador", layout="wide", page_icon="⚙️")

if 'input_fecha' not in st.session_state:
    st.session_state.input_fecha = "21 de febrero de 2026"
    st.session_state.input_cliente = "Lorena Rojas"
    st.session_state.input_tec1 = "Ignacio"
    st.session_state.input_tec2 = "Pendiente"
    st.session_state.input_temp = 66.5
    st.session_state.input_p_carga = 7.5
    st.session_state.input_p_descarga = 7.0
    st.session_state.input_estado = "El equipo se encuentra funcionando en óptimas condiciones..."

st.title("⚙️ Sistema de Mantenimiento InforGem")
st.markdown("---")

# Inventario actualizado según tu lista
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

# --- SECCIÓN SUPERIOR ---
col_busqueda, col_plan = st.columns(2)
with col_busqueda:
    tag_seleccionado = st.selectbox("🔍 TAG del Equipo:", list(inventario_equipos.keys()))
    mod_d, ser_d, area_d, ubi_d = inventario_equipos[tag_seleccionado]
    if st.button("Buscar Historial"):
        reg = buscar_ultimo_registro(tag_seleccionado)
        if reg:
            st.session_state.input_cliente, st.session_state.input_temp = reg[1], float(reg[2])
            st.session_state.input_estado, st.session_state.input_tec1, st.session_state.input_tec2 = reg[3], reg[5], reg[6]
            st.session_state.input_p_carga, st.session_state.input_p_descarga = float(reg[7]), float(reg[8])
            st.success("Historial cargado.")
with col_plan:
    tipo_plan = st.selectbox("🛠️ Tipo Intervención:", ["Inspección", "P1", "P2", "P3"])

st.markdown("---")
# --- FORMULARIO DE DATOS ---
c1, c2, c3, c4 = st.columns(4)
modelo = c1.text_input("Modelo:", value=mod_d)
numero_serie = c2.text_input("Serie:", value=ser_d)
area = c3.text_input("Área:", value=area_d)
ubicacion = c4.text_input("Ubicación:", value=ubi_d)

c5, c6, c7, c8 = st.columns(4)
fecha = c5.text_input("Fecha:", key="input_fecha")
cliente_contacto = c6.text_input("Contacto Cliente:", key="input_cliente")
tecnico_1 = c7.text_input("Técnico 1:", key="input_tec1")
tecnico_2 = c8.text_input("Técnico 2:", key="input_tec2")

c9, c10, c11 = st.columns(3)
p_carga = c9.number_input("Presión Carga (bar):", step=0.1, key="input_p_carga")
p_descarga = c10.number_input("Presión Descarga (bar):", step=0.1, key="input_p_descarga")
temp_salida = c11.number_input("Temp Salida (°C):", step=0.1, key="input_temp")

estado_entrega = st.text_area("Estado de Entrega:", key="input_estado")

# --- ACCIÓN ---
if st.button(f"Generar Word de {tipo_plan}", type="primary"):
    try:
        doc = DocxTemplate("plantilla/inspeccion.docx")
        
        # MAPEO EXACTO A TU IMAGEN 96c602.png
        context = {
            "tipo_intervencion": tipo_plan,
            "modelo": modelo,
            "tag": tag_seleccionado,
            "area": area,
            "ubicacion": ubicacion,
            "cliente_contacto": cliente_contacto,
            "p_carga": p_carga,
            "p_descarga": p_descarga,
            "temp_salida": temp_salida,
            "tecnico_1": tecnico_1,
            "tecnico_2": tecnico_2,
            "estado_entrega": estado_entrega,
            "fecha": fecha
        }
        doc.render(context)

        # Guardado local y registro
        folder = os.path.join("Historial_Informes", tag_seleccionado)
        os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, f"Informe_{tipo_plan}_{tag_seleccionado}.docx")
        doc.save(path)
        guardar_registro(tag_seleccionado, modelo, numero_serie, area, ubicacion, fecha, cliente_contacto, tecnico_1, tecnico_2, temp_salida, p_carga, p_descarga, estado_entrega, tipo_plan, path)
        
        st.success(f"Guardado local en: {path}")
        exito_nube, msg_nube = sincronizar_con_nube(tag_seleccionado, tipo_plan)
        st.info(msg_nube)

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        st.download_button("⬇️ Descargar Word", data=buffer, file_name=f"Informe_{tag_seleccionado}.docx")
    except Exception as e:
        st.error(f"Error: {e}")