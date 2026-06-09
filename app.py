# =========================
# EXPORTACIÓN A EXCEL
# =========================

def exportar_a_excel(df_completo, df_incidencias, resumen_texto, faltantes_list):
    """Exporta los datos a un archivo Excel con múltiples hojas"""
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Hoja 1: Resumen ejecutivo
        resumen_data = {
            'Métrica': ['Total partidas', 'Partidas con incidencias', '% Incidencias',
                       'Importe total (€)', 'Importe partidas críticas (€)',
                       'Fecha análisis', 'Sistemas faltantes'],
            'Valor': [
                len(df_completo),
                len(df_incidencias),
                f"{(len(df_incidencias)/len(df_completo)*100):.1f}%",
                f"{df_completo['importe'].sum():,.2f}",
                f"{df_completo[df_completo['critica_coste']]['importe'].sum():,.2f}",
                pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S'),
                ', '.join(faltantes_list) if faltantes_list else 'No detectados'
            ]
        }
        df_resumen = pd.DataFrame(resumen_data)
        df_resumen.to_excel(writer, sheet_name='Resumen', index=False)
        
        # Hoja 2: Conteo de alertas
        alertas_count = {
            'Alerta': ['Partidas alzadas', 'Incompletas', 'Elementos que sobran', 
                      'Deben dividirse', 'Sin medición', 'Sin descomposición',
                      'Sin texto', 'Duplicadas', 'Contradictorias', 
                      'Precios bajos', 'Mano de obra irreal', 'Críticas por coste'],
            'Cantidad': [
                int(df_completo['alzada'].sum()),
                int(df_completo['incompleta'].sum()),
                int(df_completo['sobran_elementos'].sum()),
                int(df_completo['debe_dividirse'].sum()),
                int(df_completo['sin_medicion'].sum()),
                int(df_completo['sin_descomposicion'].sum()),
                int(df_completo['sin_texto'].sum()),
                int(df_completo['duplicada'].sum()),
                int(df_completo['contradictoria'].sum()),
                int(df_completo['precio_bajo'].sum()),
                int(df_completo['mo_irreal'].sum()),
                int(df_completo['critica_coste'].sum())
            ]
        }
        df_alertas = pd.DataFrame(alertas_count)
        df_alertas.to_excel(writer, sheet_name='Conteo_alertas', index=False)
        
        # Hoja 3: Todas las partidas
        cols_export = ["codigo", "subpartida", "subpartida_detalle", "texto", "unidad", 
                       "cantidad", "precio", "importe", "valoracion"]
        df_completo[cols_export].to_excel(writer, sheet_name='Todas_partidas', index=False)
        
        # Hoja 4: Solo incidencias
        if not df_incidencias.empty:
            df_incidencias[cols_export].to_excel(writer, sheet_name='Incidencias', index=False)
        
        # Hoja 5: Resumen de texto
        df_texto_resumen = pd.DataFrame({'Resumen': [resumen_texto]})
        df_texto_resumen.to_excel(writer, sheet_name='Resumen_texto', index=False)
        
        # Ajustar anchos de columnas
        for sheetname in writer.sheets:
            worksheet = writer.sheets[sheetname]
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
    
    return output.getvalue()


# =========================
# INTERFAZ CON BOTONES DE EXCEL
# =========================

st.divider()
st.subheader("📥 Exportar resultados")

# Generar resumen de texto para exportar
total_partidas = len(df)
cols_problemas = ["alzada", "incompleta", "sobran_elementos", "debe_dividirse",
                 "sin_medicion", "sin_descomposicion", "sin_texto", "duplicada",
                 "contradictoria", "precio_bajo", "mo_irreal", "critica_coste"]
df_incidencias = df[df[cols_problemas].any(axis=1)]

resumen_texto = f"""
========================================
AUDITORÍA DE PRESUPUESTO - RESUMEN
========================================

📊 DATOS GENERALES:
- Total partidas analizadas: {len(df)}
- Partidas con incidencias: {len(df_incidencias)}
- Porcentaje de incidencias: {(len(df_incidencias)/len(df)*100):.1f}%

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

💰 RESUMEN ECONÓMICO:
- Importe total: {df['importe'].sum():,.2f} €
- Importe partidas críticas: {df[df['critica_coste']]['importe'].sum():,.2f} €

========================================
"""

col1, col2, col3 = st.columns(3)

with col1:
    if st.button("📊 Exportar a EXCEL", use_container_width=True, type="primary"):
        with st.spinner("Generando archivo Excel..."):
            excel_data = exportar_a_excel(df, df_incidencias, resumen_texto, faltantes_global)
            st.download_button(
                label="⬇️ Descargar Excel",
                data=excel_data,
                file_name=f"auditoria_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="descarga_excel"
            )

with col2:
    if st.button("📄 CSV completo", use_container_width=True):
        export_df = df[["codigo", "subpartida", "subpartida_detalle", "texto", "unidad", 
                        "cantidad", "precio", "importe", "valoracion"]].copy()
        csv = export_df.to_csv(index=False, encoding='utf-8-sig')
        st.download_button(
            label="⬇️ Descargar CSV",
            data=csv,
            file_name="auditoria_completa.csv",
            mime="text/csv",
            key="descarga_csv_completa"
        )

with col3:
    if st.button("⚠️ Solo incidencias CSV", use_container_width=True):
        if not df_incidencias.empty:
            export_df = df_incidencias[["codigo", "subpartida", "subpartida_detalle", "texto", 
                                         "unidad", "cantidad", "precio", "importe", "valoracion"]].copy()
            csv = export_df.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="⬇️ Descargar CSV",
                data=csv,
                file_name="auditoria_incidencias.csv",
                mime="text/csv",
                key="descarga_csv_incidencias"
            )
        else:
            st.info("No hay incidencias para exportar")

# Botón adicional para resumen en texto
st.divider()
if st.button("📝 Descargar resumen (TXT)", use_container_width=True):
    st.download_button(
        label="⬇️ Descargar TXT",
        data=resumen_texto,
        file_name=f"resumen_auditoria_{pd.Timestamp.now().strftime('%Y%m%d')}.txt",
        mime="text/plain",
        key="descarga_txt"
    )
