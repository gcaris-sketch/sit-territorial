st.markdown("---")
    st.markdown("### 🗺️ Guía de Navegación")
    
    vendedora_destino = st.text_input("Vendedora a cargo:", placeholder="Ej: María Pérez")
    
    total_ruta = len(st.session_state.ruta_ferreterias) + len(st.session_state.ruta_obras)
    st.markdown(f"**Puntos en ruta:** {total_ruta}")
    
    if total_ruta > 0:
        texto_guia = f"GUÍA DE NAVEGACIÓN COMERCIAL\n"
        texto_guia += f"================================\n"
        texto_guia += f"Fecha: {date.today().strftime('%d/%m/%Y')}\n"
        texto_guia += f"Vendedora Asignada: {vendedora_destino.upper() if vendedora_destino else 'NO ESPECIFICADA'}\n\n"
        
        if st.session_state.ruta_ferreterias:
            texto_guia += "🛠️ FERRETERÍAS A VISITAR\n"
            texto_guia += "--------------------------------\n"
            puntos_ruta = []
            for i, f_id in enumerate(st.session_state.ruta_ferreterias):
                row_f = df_ferreterias.loc[f_id]
                tel = row_f.get('telefono', 'S/N')
                link = row_f.get('url_google_maps') or f"https://www.google.com/maps/search/?api=1&query={row_f['lat']},{row_f['lon']}"
                texto_guia += f"{i+1}. {row_f['nombre'].upper()}\n"
                texto_guia += f"   - Dirección: {row_f['direccion']}, {row_f['comuna']}\n"
                texto_guia += f"   - Teléfono: {tel}\n"
                texto_guia += f"   - Mapa: {link}\n\n"
                puntos_ruta.append(f"{row_f['lat']},{row_f['lon']}")
                
            if puntos_ruta:
                url_gmaps = f"https://www.google.com/maps/dir/?api=1&destination={puntos_ruta[-1]}" + (f"&waypoints={urllib.parse.quote('|'.join(puntos_ruta[:-1]))}" if len(puntos_ruta)>1 else "")
                st.link_button("🚗 Ruta Ferreterías en Maps", url_gmaps, use_container_width=True)

        if st.session_state.ruta_obras:
            texto_guia += "🏗️ OBRAS SEIA A VISITAR\n"
            texto_guia += "--------------------------------\n"
            for i, o_id in enumerate(st.session_state.ruta_obras):
                row_o = df_obras.loc[o_id]
                link = row_o.get('url_seia', 'S/N')
                mapa_obra = f"https://www.google.com/maps/search/?api=1&query={row_o['lat']},{row_o['lon']}"
                texto_guia += f"{i+1}. {str(row_o['titulo']).upper()}\n"
                texto_guia += f"   - Titular: {row_o.get('empresa', 'S/N')}\n"
                texto_guia += f"   - Comuna: {row_o['comuna']}\n"
                texto_guia += f"   - Ficha SEIA: {link}\n"
                texto_guia += f"   - Mapa: {mapa_obra}\n\n"

        # Botón de Descarga
        st.download_button(
            label="📄 Descargar Guía (TXT)",
            data=texto_guia,
            file_name=f"Guia_Ruta_{vendedora_destino.replace(' ', '_')}_{date.today()}.txt",
            mime="text/plain",
            use_container_width=True
        )

        if st.button("🗑️ Limpiar Ruta Completa", use_container_width=True): 
            st.session_state.ruta_ferreterias = []
            st.session_state.ruta_obras = []
            st.rerun()
