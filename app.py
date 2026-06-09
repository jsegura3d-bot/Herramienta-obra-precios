# ... (todo el código anterior hasta aquí)

# Mostrar resultados
st.subheader("📋 Partidas analizadas")

cols_mostrar = ["codigo", "subpartida", "subpartida_detalle", "texto", "unidad", "cantidad", "precio", "importe", "valoracion"]

st.dataframe(df_view[cols_mostrar].reset_index(drop=True), use_container_width=True,
             column_config={
                 "codigo": "Código", 
                 "subpartida": "Capítulo", 
                 "subpartida_detalle": "Subpartida",
                 "texto": "Descripción", 
                 "unidad": "Ud",
                 "cantidad": st.column_config.NumberColumn("Cantidad", format="%.2f"),
                 "precio": st.column_config.NumberColumn("Precio (€)", format="%.2f €"),
                 "importe": st.column_config.NumberColumn("Importe (€)", format="%.2f €"),
                 "valoracion": "Valoración"
             })

# =========================
# EXPORTACIÓN CON BOTONES
# =========================

st.divider()
st.subheader("📥 Exportar resultados")

col_export1, col_export2 = st.columns(2)

with col_export1:
    if st.button("📊 CSV completo", use_container_width=True):
        export_df = df[["codigo", "subpartida", "subpartida_detalle", "texto", "unidad", 
                        "cantidad", "precio", "importe", "valoracion"]].copy()
        csv = export_df.to_csv(index=False, encoding='utf-8-sig')
        st.download_button(
            label="⬇️ Descargar",
            data=csv,
            file_name="auditoria_completa.csv",
            mime="text/csv",
            key="descarga_completa"
        )

with col_export2:
    cols_problemas = ["alzada", "incompleta", "sobran_elementos", "debe_dividirse",
                     "sin_medicion", "sin_descomposicion", "sin_texto", "duplicada",
                     "contradictoria", "precio_bajo", "mo_irreal", "critica_coste"]
    
    if st.button("⚠️ Solo incidencias", use_container_width=True):
        df_problemas = df[df[cols_problemas].any(axis=1)]
        export_df = df_problemas[["codigo", "subpartida", "subpartida_detalle", "texto", 
                                   "unidad", "cantidad", "precio", "importe", "valoracion"]].copy()
        csv = export_df.to_csv(index=False, encoding='utf-8-sig')
        st.download_button(
            label="⬇️ Descargar",
            data=csv,
            file_name="auditoria_incidencias.csv",
            mime="text/csv",
            key="descarga_incidencias"
        )

# Botón de resumen
if st.button("📝 Resumen ejecutivo", use_container_width=True):
    total_partidas = len(df)
    partidas_con_problemas = df[cols_problemas].any(axis=1).sum()
    
    resumen = f"""
    ========================================
    AUDITORÍA DE PRESUPUESTO - RESUMEN
    ========================================
    
    📊 DATOS GENERALES:
    - Total partidas analizadas: {total_partidas}
    - Partidas con incidencias: {partidas_con_problemas}
    - Porcentaje de incidencias: {(partidas_con_problemas/total_partidas*100):.1f}%
    
    📋 DETALLE DE ALERTAS:
    - Partidas alzadas: {int(df['alzada'].sum())}
    - Incompletas: {int(df['incompleta'].sum())}
    - Con elementos que sobran: {int(df['sobran_elementos'].sum())}
    - Deben dividirse: {int(df['debe_dividirse'].sum())}
    - Sin medición: {int(df['sin_medicion'].sum())}
    - Sin descomposición: {int(df['sin_descomposicion'].sum())}
    - Sin texto informativo: {int(df['sin_texto'].sum())}
    - Duplicadas: {int(df['duplicada'].sum())}
    - Contradictorias: {int(df['contradictoria'].sum())}
    - Precios bajos: {int(df['precio_bajo'].sum())}
    - Mano de obra irreal: {int(df['mo_irreal'].sum())}
    - Críticas por coste: {int(df['critica_coste'].sum())}
    """
    
    if faltantes_global:
        resumen += f"\n\n🚨 SISTEMAS FALTANTES:\n"
        for s in faltantes_global:
            resumen += f"- {s}\n"
    
    resumen += f"""
    
    💰 RESUMEN ECONÓMICO:
    - Importe total: {df['importe'].sum():,.2f} €
    - Importe partidas críticas: {df[df['critica_coste']]['importe'].sum():,.2f} €
    
    ========================================
    Fecha: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}
    ========================================
    """
    
    st.download_button(
        label="⬇️ Descargar resumen",
        data=resumen,
        file_name="resumen_auditoria.txt",
        mime="text/plain",
        key="descarga_resumen"
    )

# Mostrar información de depuración (opcional)
with st.expander("ℹ️ Información del análisis"):
    st.write(f"**Formato detectado:** {detectar_formato_bc3(bc3_file.getvalue().decode('latin-1', errors='ignore'))}")
    st.write(f"**Total partidas procesadas:** {len(df)}")
    st.write(f"**Soporte PDF:** {'✅ Sí' if PDF_SUPPORT else '❌ No'}")
    st.write(f"**Soporte DOCX:** {'✅ Sí' if DOCX_SUPPORT else '❌ No'}")
