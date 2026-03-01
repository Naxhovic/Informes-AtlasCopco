img.save(img_io, format='PNG')
                    img_io.seek(0)
                    return img_io

                io_tec = procesar_imagen_firma(canvas_tec.image_data)
                io_cli = procesar_imagen_firma(canvas_cli.image_data)
                
                informes_finales = []
                with st.spinner("Fabricando documentos oficiales, inyectando firmas y transformando a PDF..."):
                    try:
                        for inf in st.session_state.informes_pendientes:
                            doc = DocxTemplate(inf['file_plantilla'])
                            context = inf['context']
                            
                            context['firma_tecnico'] = InlineImage(doc, io_tec, width=Mm(40))
                            context['firma_cliente'] = InlineImage(doc, io_cli, width=Mm(40))
                            
                            doc.render(context)
                            doc.save(inf['ruta_docx'])
                            
                            ruta_pdf_gen = convertir_a_pdf(inf['ruta_docx'])
                            
                            if ruta_pdf_gen:
                                ruta_final = ruta_pdf_gen
                                nombre_final = inf['nombre_archivo_base'].replace(".docx", ".pdf")
                            else:
                                ruta_final = inf['ruta_docx']
                                nombre_final = inf['nombre_archivo_base']
                            
                            nombre_codificado = f"{inf['area'].title()}@@{inf['tag']}@@{nombre_final}"
                            
                            # Guardamos en Google Sheets
                            tupla_lista = list(inf['tupla_db'])
                            tupla_lista[18] = ruta_final
                            guardar_registro(tuple(tupla_lista))
                            
                            informes_finales.append({"tag": inf['tag'], "tipo": inf['tipo_plan'], "ruta": ruta_final, "nombre_archivo": nombre_codificado})
                            
                        exito, mensaje_correo = enviar_carrito_por_correo(MI_CORREO_CORPORATIVO, informes_finales)
                        
                        if exito:
                            st.success("✅ ¡PERFECTO! Los documentos oficiales se firmaron, convirtieron a PDF y ya están camino a tu OneDrive.")
                            st.session_state.informes_pendientes = []  
                            st.balloons()
                        else:
                            st.error(f"Error de red: {mensaje_correo}")
                            
                    except Exception as e:
                        st.error(f"Error sistémico procesando las firmas: {e}")
            else:
                st.warning("⚠️ Asegúrate de dibujar en ambas pizarras antes de generar los PDFs finales.")

    # --- 6.2 VISTA CATÁLOGO (Dashboard interactivo) ---
    elif st.session_state.equipo_seleccionado is None:
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
        with col_filtro: filtro_tipo = st.radio("🗂️ Categoría de Equipo:", ["Todos", "Compresores", "Secadores"], horizontal=True)
        with col_busqueda: busqueda = st.text_input("🔍 Buscar activo por TAG, Modelo o Área...", placeholder="Ejemplo: GA 250, 35-GC-006...").lower()
            
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

    # --- 6.3 VISTA FORMULARIO Y GENERACIÓN ---
    else:
        tag_sel = st.session_state.equipo_seleccionado
        mod_d, ser_d, area_d, ubi_d = inventario_equipos[tag_sel]
        
        c_btn, c_tit = st.columns([1, 4])
        with c_btn: st.button("⬅️ Volver", on_click=volver_catalogo, use_container_width=True)
        with c_tit: st.markdown(f"<h1 style='margin-top:-15px;'>⚙️ Ficha de Servicio: <span style='color:#007CA6;'>{tag_sel}</span></h1>", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        # 👇 AHORA SOLO HAY 4 PESTAÑAS 👇
        tab1, tab2, tab3, tab4 = st.tabs(["📋 1. Reporte y Diagnóstico", "📚 2. Ficha Técnica", "🔍 3. Bitácora de Observaciones", "👤 4. Gestión de Área"])
        
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
                
                if st.session_state.input_cliente in opciones: cli_idx = opciones.index(st.session_state.input_cliente)
                else: cli_idx = 1 if len(contactos_db) > 0 else 0
                
                sc1, sc2 = st.columns([4, 1])
                with sc1: cli_sel = st.selectbox("Contacto Cliente", opciones, index=cli_idx)
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

            # 👇 AQUÍ SE UNIÓ EL DIAGNÓSTICO FINAL A LA PESTAÑA 1 👇
            st.markdown("<hr>", unsafe_allow_html=True)
            st.markdown("### Evaluación y Diagnóstico Final")
            est_eq = st.radio("Estado de Devolución del Activo:", ["Operativo", "Fuera de servicio"], key="input_estado_eq", horizontal=True)
            est_ent = st.text_area("Descripción Condición Final:", key="input_estado", height=100)
            reco = st.text_area("Recomendaciones / Acciones Pendientes:", key="input_reco", height=100)
            
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("📥 Guardar y Añadir a la Bandeja de Firmas", type="primary", use_container_width=True):
                if "CD" in tag_sel: file_plantilla = "plantilla/secadorfueradeservicio.docx" if est_eq == "Fuera de servicio" else "plantilla/inspeccionsecador.docx"
                else:
                    if est_eq == "Fuera de servicio": file_plantilla = "plantilla/fueradeservicio.docx"
                    elif tipo_plan == "P1": file_plantilla = "plantilla/p1.docx"
                    elif tipo_plan == "P2": file_plantilla = "plantilla/p2.docx"
                    elif tipo_plan == "P3": file_plantilla = "plantilla/p3.docx"
                    else: file_plantilla = "plantilla/inspeccion.docx"
                    
                context = {"tipo_intervencion": tipo_plan, "modelo": mod_d, "tag": tag_sel, "area": area_d, "ubicacion": ubi_d, "cliente_contacto": cli_cont, "p_carga": f"{p_c_clean} {unidad_p}", "p_descarga": f"{p_d_clean} {unidad_p}", "temp_salida": t_salida_clean, "horas_marcha": int(h_m), "horas_carga": int(h_c), "tecnico_1": tec1, "tecnico_2": tec2, "estado_equipo": est_eq, "estado_entrega": est_ent, "recomendaciones": reco, "serie": ser_d, "tipo_orden": tipo_plan.upper(), "fecha": fecha, "equipo_modelo": mod_d}
                
                nombre_archivo = f"Informe_{tipo_plan}_{tag_sel}_{fecha.replace(' ','_')}.docx"
                ruta = os.path.join(RUTA_ONEDRIVE, nombre_archivo)
                
                try: temp_db = float(t_salida_clean)
                except: temp_db = 0.0
                tupla_db = (tag_sel, mod_d, ser_d, area_d, ubi_d, fecha, cli_cont, tec1, tec2, temp_db, f"{p_c_clean} {unidad_p}", f"{p_d_clean} {unidad_p}", h_m, h_c, est_ent, tipo_plan, reco, est_eq, "", st.session_state.usuario_actual)
                
                with st.spinner("Creando borrador del documento para vista preliminar..."):
                    doc_prev = DocxTemplate(file_plantilla)
                    ctx_prev = context.copy()
                    ctx_prev['firma_tecnico'] = "" 
                    ctx_prev['firma_cliente'] = ""
                    doc_prev.render(ctx_prev)
                    
                    os.makedirs(RUTA_ONEDRIVE, exist_ok=True)
                    ruta_prev_docx = os.path.join(RUTA_ONEDRIVE, f"PREVIEW_{nombre_archivo}")
                    doc_prev.save(ruta_prev_docx)
                    ruta_prev_pdf = convertir_a_pdf(ruta_prev_docx)
                
                st.session_state.informes_pendientes.append({
                    "tag": tag_sel, "area": area_d, "tec1": tec1, "cli": cli_cont, "tipo_plan": tipo_plan,
                    "file_plantilla": file_plantilla, "context": context, "tupla_db": tupla_db,
                    "ruta_docx": ruta, "nombre_archivo_base": nombre_archivo, "ruta_prev_pdf": ruta_prev_pdf
                })
                st.success("✅ Datos guardados. Agrega otro equipo o ve a la bandeja para firmar.")
                st.session_state.equipo_seleccionado = None
                st.rerun()

        # 👇 PESTAÑA 2: Ficha Técnica (Antes Pestaña 3) 👇
        with tab2:
            st.markdown(f"### 📘 Datos Técnicos y Repuestos ({mod_d})")
            with st.expander("✏️ Agregar o Corregir Datos Faltantes (N° Parte, Aceite, etc.)"):
                with st.form(key=f"form_specs_{tag_sel}"):
                    st.info(f"💡 Lo que guardes aquí se actualizará para todos los equipos modelo **{mod_d}**.")
                    c_e1, c_e2 = st.columns(2)
                    opc_claves = ["N° Parte Filtro Aceite", "N° Parte Filtro Aire", "N° Parte Kit", "N° Parte Separador", "Litros de Aceite", "Tipo de Aceite", "Cant. Filtros Aceite", "Cant. Filtros Aire", "Otro dato nuevo..."]
                    clave_sel = c_e1.selectbox("¿Qué dato vas a ingresar?", opc_claves)
                    if clave_sel == "Otro dato nuevo...": clave_final = c_e1.text_input("Escribe el nombre del dato:")
                    else: clave_final = clave_sel
                    valor_final = c_e2.text_input("Ingresa el valor (Ej: 1613 6105 00):")
                    
                    if st.form_submit_button("💾 Guardar en Base de Datos", use_container_width=True):
                        with st.spinner("Guardando en la nube..."):
                            if clave_final and valor_final:
                                guardar_especificacion_db(mod_d, clave_final.strip(), valor_final.strip())
                                st.success("✅ ¡Dato guardado en Google Sheets!")
                                st.rerun()
                            else: st.error("⚠️ Debes llenar el valor a guardar.")
            
            if mod_d in ESPECIFICACIONES:
                specs = {k: v for k, v in ESPECIFICACIONES[mod_d].items() if k != "Manual"}
                if specs:
                    cols = st.columns(3)
                    for i, (k, v) in enumerate(specs.items()):
                        with cols[i % 3]: st.markdown(f"<div style='background-color: #1e2530; padding: 15px; border-radius: 8px; margin-bottom: 15px; border-left: 4px solid #007CA6;'><span style='color: #8c9eb5; font-size: 0.85em; text-transform: uppercase; font-weight: bold;'>{k}</span><br><span style='color: white; font-size: 1.1em;'>{v}</span></div>", unsafe_allow_html=True)
                else: st.info("No hay datos técnicos registrados aún. ¡Sé el primero en agregar uno arriba!")
                    
                st.markdown("<hr>", unsafe_allow_html=True)
                st.markdown("### 📥 Documentación y Manuales")
                if "Manual" in ESPECIFICACIONES[mod_d] and os.path.exists(ESPECIFICACIONES[mod_d]["Manual"]):
                    with open(ESPECIFICACIONES[mod_d]["Manual"], "rb") as f: st.download_button(label=f"📕 Descargar Manual de {mod_d} (PDF)", data=f, file_name=ESPECIFICACIONES[mod_d]["Manual"].split('/')[-1], mime="application/pdf")
                else: st.info("ℹ️ El manual o despiece para este modelo aún no ha sido cargado en la plataforma.")
            else: st.warning(f"⚠️ No hay especificaciones técnicas base para el modelo {mod_d}.")

        # 👇 PESTAÑA 3: Bitácora (Antes Pestaña 4) 👇
        with tab3:
            st.markdown(f"### 🔍 Bitácora Permanente del Equipo: {tag_sel}")
            st.info("💡 Usa este espacio para dejar notas importantes sobre el estado general del equipo.")
            
            with st.form(key=f"form_obs_{tag_sel}"):
                nueva_obs = st.text_area("Escribe una nueva observación o nota técnica para este equipo:", height=100)
                if st.form_submit_button("➕ Dejar constancia en la bitácora", use_container_width=True):
                    with st.spinner("Guardando en la nube..."):
                        if nueva_obs:
                            agregar_observacion(tag_sel, st.session_state.usuario_actual, nueva_obs)
                            st.success("✅ Observación registrada con éxito en Google Sheets.")
                            st.rerun()
                        else:
                            st.warning("⚠️ Debes escribir algo antes de guardar.")
            
            st.markdown("---")
            st.markdown("#### 📜 Historial de Observaciones Anteriores")
            df_obs = obtener_observaciones(tag_sel)
            
            if not df_obs.empty:
                for _, row in df_obs.iterrows():
                    col_obs, col_del = st.columns([11, 1])
                    with col_obs:
                        st.markdown(f"""
                            <div style='background-color: #2b303b; padding: 15px; border-radius: 8px; margin-bottom: 10px; border-left: 4px solid #FF6600;'>
                                <small style='color: #aeb9cc;'><b>👤 Técnico: {row['usuario']}</b> &nbsp;|&nbsp; 📅 Fecha: {row['fecha']}</small><br>
                                <span style='color: white; font-size: 1.05em;'>{row['texto']}</span>
                            </div>
                        """, unsafe_allow_html=True)
                    with col_del:
                        st.markdown("<div style='margin-top: 20px;'></div>", unsafe_allow_html=True)
                        if st.button("🗑️", key=f"del_obs_{row['id']}", help="Borrar esta nota"):
                            eliminar_observacion(row['id'])
                            st.rerun()
            else:
                st.caption("No hay observaciones registradas para este equipo aún. ¡Escribe la primera!")

        # 👇 PESTAÑA 4: Gestión de Área (Antes Pestaña 5) 👇
        with tab4:
            st.markdown(f"### 👤 Información de Contactos y Seguridad del Área: {tag_sel}")
            st.info("💡 Asigna y actualiza los dueños del área, el PEA y la frecuencia radial correspondientes a este equipo.")

            with st.expander("✏️ Editar o Agregar Contacto / Dato de Seguridad"):
                with st.form(key=f"form_area_{tag_sel}"):
                    c_a1, c_a2 = st.columns(2)
                    opc_area = ["Dueño de Área (Turno 1-3)", "Dueño de Área (Turno 2-4)", "PEA", "Frecuencia Radial", "Supervisor a cargo", "Jefe de Turno", "Otro cargo..."]
                    clave_sel_area = c_a1.selectbox("¿Qué dato vas a ingresar?", opc_area)
                    if clave_sel_area == "Otro cargo...": clave_final_area = c_a1.text_input("Escribe el nombre del cargo/dato:")
                    else: clave_final_area = clave_sel_area
                    valor_final_area = c_a2.text_input("Ingresa la información:")
                    
                    if st.form_submit_button("💾 Guardar Información", use_container_width=True):
                        with st.spinner("Guardando en la nube..."):
                            if clave_final_area and valor_final_area:
                                guardar_dato_equipo(tag_sel, clave_final_area.strip(), valor_final_area.strip())
                                st.success("✅ ¡Dato actualizado exitosamente en Google Sheets!")
                                st.rerun()
                            else: st.error("⚠️ Debes llenar ambos campos.")

            datos_equipo = obtener_datos_equipo(tag_sel)
            if "Dueño de Área (Turno 1-3)" not in datos_equipo: datos_equipo["Dueño de Área (Turno 1-3)"] = "No asignado"
            if "Dueño de Área (Turno 2-4)" not in datos_equipo: datos_equipo["Dueño de Área (Turno 2-4)"] = "No asignado"
            if "PEA" not in datos_equipo: datos_equipo["PEA"] = "No asignado"
            if "Frecuencia Radial" not in datos_equipo: datos_equipo["Frecuencia Radial"] = "No asignada"

            cols_area = st.columns(2)
            for i, (k, v) in enumerate(datos_equipo.items()):
                with cols_area[i % 2]:
                    st.markdown(f"""
                        <div style='background-color: #2b303b; padding: 15px; border-radius: 8px; margin-bottom: 15px; border-left: 4px solid #FF6600;'>
                            <span style='color: #aeb9cc; font-size: 0.85em; text-transform: uppercase; font-weight: bold;'>{k}</span><br>
                            <span style='color: white; font-size: 1.1em;'>{v}</span>
                        </div>
                    """, unsafe_allow_html=True)

        st.markdown("<br><hr>", unsafe_allow_html=True)
        st.markdown("### 📋 Trazabilidad Histórica de Intervenciones (Reportes Creados)")
        df_hist = obtener_todo_el_historial(tag_sel)
        if not df_hist.empty: st.dataframe(df_hist, use_container_width=True)