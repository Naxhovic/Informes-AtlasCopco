import streamlit as st
from docxtpl import DocxTemplate
import io
import os
import sqlite3
import subprocess
import pandas as pd

# --- 1. MÓDULO DE BASE DE DATOS LOCAL ---
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
            estado_entrega TEXT, tipo_intervencion TEXT, recomendaciones TEXT, ruta_archivo TEXT
        )
    ''')
    conn.commit()
    conn.close()

def guardar_registro(tag, mod, serie, area, ubi, fecha, cli, tec1, tec2, temp, p_c, p_d, h_m, h_c, est, tipo, reco, ruta):
    conn = sqlite3.connect("historial_equipos.db")
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO intervenciones 
        (tag, modelo, numero_serie, area, ubicacion, fecha, cliente_contacto, tecnico_1, tecnico_2, 
        temp_salida, p_carga, p_descarga, horas_marcha, horas_carga, estado_entrega, tipo_intervencion, recomendaciones, ruta_archivo)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (tag, mod, serie, area, ubi, fecha, cli, tec1, tec2, temp, p_c, p_d, h_m, h_c, est, tipo, reco, ruta))
    conn.commit()
    conn.close()

def buscar_ultimo_registro(tag):
    conn = sqlite3.connect("historial_equipos.db")
    cursor = conn.cursor()
    cursor.execute('''
        SELECT fecha, cliente_contacto, temp_salida, estado_entrega, tipo_intervencion, 
               tecnico_1, tecnico_2, p_carga, p_descarga, horas_marcha, horas_carga, recomendaciones
        FROM intervenciones WHERE tag = ? ORDER BY id DESC LIMIT 1
    ''', (tag,))
    resultado = cursor.fetchone()
    conn.close()
    return resultado

def obtener_todo_el_historial(tag):
    conn = sqlite3.connect("historial_equipos.db")
    query = "SELECT fecha, tipo_intervencion, horas_marcha, p_carga, p_descarga, temp_salida, tecnico_1 FROM intervenciones WHERE tag = ? ORDER BY id DESC"
    df = pd.read_sql_query(query, conn, params=(tag,))
    conn.close()
    return df

# --- 2. MÓDULO DE NUBE (GITHUB) ---
def sincronizar_con_nube(tag, tipo_plan):
    try:
        subprocess.run(["git", "add", "."], check=True, capture_output=True, text=True)
        mensaje = f"Reporte: {tipo_plan} - {tag}"
        subprocess.run(["git", "commit", "-m", mensaje], check=True, capture_output=True, text=True)
        subprocess.run(["git", "push"], check=True, capture_output=True, text=True)
        return True, "☁️ Sincronización con GitHub exitosa."
    except Exception:
        return False, "⚠️ Pendiente de subir a GitHub."

# --- 3. CONFIGURACIÓN DE INTERFAZ ---
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
    st.session_state.input_h_marcha = 0.0
    st.session_state.input_h_carga = 0.0
    st.session_state.input_reco = "Realizar limpieza de filtros y chequeo de tensiones en la próxima visita."

st.title("⚙️ Sistema de Mantenimiento InforGem")
st.markdown("---")

# Inventario Maestro
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

# --- BÚSQUEDA ---
col_busqueda, col_plan = st.columns(2)
with col_busqueda:
    tag_seleccionado = st.selectbox("🔍 TAG del Equipo:", list(inventario_equipos.keys()))
    mod_d, ser_d, area_d, ubi_d = inventario_equipos[tag_seleccionado]
    if st.button("Buscar Historial"):
        reg = buscar_ultimo_registro(tag_seleccionado)
        if reg:
            st.session_state.input_cliente, st.session_state.input_temp = reg[1], float(reg[2])
            st.session_state.input_estado, st.session_state.input_tec1 = reg[3], reg[5]
            st.session_state.input_tec2 = reg[6]
            try: st.session_state.input_p_carga = float(str(reg[7]).split()[0])
            except: st.session_state.input_p_carga = 7.0
            try: st.session_state.input_p_descarga = float(str(reg[8]).split()[0])
            except: st.session_state.input_p_descarga = 7.5
            st.session_state.input_h_marcha, st.session_state.input_h_carga = float(reg[9]), float(reg[10])
            st.session_state.input_reco = reg[11] if reg[11] else ""
            st.success(f"Cargado. Última visita: {reg[0]}")

with col_plan:
    tipo_plan = st.selectbox("🛠️ Tipo Intervención:", ["Inspección", "P1", "P2", "P3", "Correctivo"])

st.markdown("---")

# --- FORMULARIO ---
st.subheader("📋 Datos del Activo")
c1, c2, c3, c4 = st.columns(4)
modelo = c1.text_input("Modelo:", value=mod_d)
numero_serie = c2.text_input("N° Serie:", value=ser_d)
area = c3.text_input("Área:", value=area_d)
ubicacion = c4.text_input("Ubicación:", value=ubi_d)

st.subheader("👨‍🔧 Personal y Fecha")
c5, c6, c7, c8 = st.columns(4)
fecha = c5.text_input("Fecha:", key="input_fecha")
tecnico_1 = c6.text_input("Técnico 1:", key="input_tec1")
tecnico_2 = c7.text_input("Técnico 2:", key="input_tec2")
cliente_contacto = c8.text_input("Contacto Cliente:", key="input_cliente")

st.subheader("📊 Operación y Horómetro")
c9, c10, c11, c12, c13, c14 = st.columns(6)
horas_marcha = c9.number_input("Horas Marcha:", step=1.0, key="input_h_marcha")
horas_carga = c10.number_input("Horas Carga:", step=1.0, key="input_h_carga")
unidad_p = c11.selectbox("Unidad Presión:", ["bar", "psi"])
p_carga_val = c12.number_input(f"P. Carga ({unidad_p}):", step=0.1 if unidad_p=="bar" else 1.0, key="input_p_carga")
p_descarga_val = c13.number_input(f"P. Descarga ({unidad_p}):", step=0.1 if unidad_p=="bar" else 1.0, key="input_p_descarga")
temp_salida = c14.number_input("Temp Salida (°C):", step=0.1, key="input_temp")

st.subheader("📝 Notas y Recomendaciones")
estado_entrega = st.text_area("Estado de Entrega:")
recomendaciones = st.text_area("Nota Técnica / Recomendaciones:", key="input_reco")

# --- ACCIÓN ---
if st.button(f"🚀 Generar Reporte Industrial", type="primary"):
    try:
        doc = DocxTemplate("plantilla/inspeccion.docx")
        
        p_carga_full = f"{p_carga_val} {unidad_p}"
        p_descarga_full = f"{p_descarga_val} {unidad_p}"
        
        # MAPEO A LA NUEVA PLANTILLA
        context = {
            "tipo_intervencion": tipo_plan,
            "modelo": modelo,
            "equipo_modelo": modelo, # Se usa en la tabla inferior
            "tag": tag_seleccionado,
            "area": area,
            "ubicacion": ubicacion,
            "cliente_contacto": cliente_contacto,
            "p_carga": p_carga_full,
            "p_descarga": p_descarga_full,
            "temp_salida": temp_salida,
            "horas_marcha": int(horas_marcha),
            "horas_carga": int(horas_carga),
            "tecnico_1": tecnico_1,
            "tecnico_2": tecnico_2,
            "estado_entrega": estado_entrega,
            "recomendaciones": recomendaciones,
            "serie": numero_serie, # Vinculado al N° Serie del equipo
            "tipo_orden": "INSPECCIÓN", # Siempre en mayúsculas
            "fecha": fecha
        }
        doc.render(context)

        fecha_l = fecha.replace(" ", "_").replace("/", "-")
        nombre_archivo = f"Informe_{tipo_plan}_{tag_seleccionado}_{fecha_l}.docx"
        folder = os.path.join("Historial_Informes", tag_seleccionado)
        os.makedirs(folder, exist_ok=True)
        ruta = os.path.join(folder, nombre_archivo)
        doc.save(ruta)
        
        guardar_registro(tag_seleccionado, modelo, numero_serie, area, ubicacion, fecha, cliente_contacto, 
                         tecnico_1, tecnico_2, temp_salida, p_carga_full, p_descarga_full, horas_marcha, horas_carga, 
                         estado_entrega, tipo_plan, recomendaciones, ruta)
        
        st.success(f"✅ Reporte generado: {nombre_archivo}")
        st.info(sincronizar_con_nube(tag_seleccionado, tipo_plan)[1])
        
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        st.download_button("⬇️ Descargar Word", data=buffer, file_name=nombre_archivo)
    except Exception as e:
        st.error(f"Error: {e}")

st.markdown("---")
st.subheader(f"📜 Historial de {tag_seleccionado}")
st.dataframe(obtener_todo_el_historial(tag_seleccionado), use_container_width=True)