import streamlit as st
import pandas as pd
import io
import openpyxl

st.set_page_config(page_title="Revisor IVE BDC25 - Valencia", layout="wide")
st.title("🛠️ Revisor de Precios Pro - Corrección de Comerciales (Bugarra)")
st.caption("Configuración activa: Valores comerciales estables por categoría y marcas premium.")

# --- PREFIJOS MAESTROS IVE PARA DETECCIÓN ---
codigos_ive_referencia = ["0AF010", "EIEB20", "DRT030", "EIEC", "PIBB", "DAISA"]

uploaded_file = st.file_uploader("Sube tu Excel de Bugarra", type=["xlsx"])

if uploaded_file:
    try:
        file_bytes = uploaded_file.read()
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
        
        if 'Hoja5' in wb.sheetnames:
            ws = wb['Hoja5']
        else:
            ws = wb.active

        # Índices fijos basados estrictamente en tu plantilla real
        col_codigo_idx = 1      # Columna A (Código)
        col_resumen_idx = 2     # Columna B (Descripción/Resumen)
        col_comercial_idx = 3   # Columna C (DATOS COMERCIALES)

        # La columna REVISIÓN IA se añade al final a la derecha
        col_ia_destino = ws.max_column + 1
        ws.cell(row=1, column=col_ia_destino, value="REVISIÓN IA").font = openpyxl.styles.Font(bold=True, color="0000FF")

        resultados_vista = []

        for row_idx in range(2, ws.max_row + 1):
            codigo = str(ws.cell(row=row_idx, column=col_codigo_idx).value or '').strip()
            resumen = str(ws.cell(row=row_idx, column=col_resumen_idx).value or '').strip()
            valor_comercial = str(ws.cell(row=row_idx, column=col_comercial_idx).value or '').strip().lower()
            
            if codigo == "" or codigo.lower() == "none" or codigo.lower() == "código" or "capítulo" in codigo.lower():
                continue

            val_texto = ""
            codigo_upper = codigo.upper()
            resumen_lower = resumen.lower()

            # --- PRIORIDAD 1: CASO 3 (Comerciales - Columna C llena con info) ---
            if valor_comercial != "":
                # 1. Valores base por defecto según la rama comercial introducida
                if "iluminación" in valor_comercial or "iluminacion" in valor_comercial:
                    precio_mercado = "55.00€"
                    # Ajustes específicos para equipos especiales de iluminación detectados en el texto
                    if "koban" in resumen_lower or "luxomat" in resumen_lower or "detector" in resumen_lower:
                        precio_mercado = "120.00€"
                    elif "carandini" in resumen_lower:
                        precio_mercado = "185.00€"
                    elif "schreder" in resumen_lower or "socelec" in resumen_lower:
                        precio_mercado = "320.00€"
                        
                    val_texto = f"Mismo equipo o equivalente | Precio aprox mercado: {precio_mercado}"

                elif "mecanismos" in valor_comercial:
                    precio_mercado = "55.00€"
                    if "koban" in resumen_lower or "luxomat" in resumen_lower or "detector" in resumen_lower:
                        precio_mercado = "120.00€"
                    elif "schneider" in resumen_lower or "artec" in resumen_lower:
                        precio_mercado = "16.80€"
                        
                    val_texto = f"Mismo equipo o equivalente | Precio aprox mercado: {precio_mercado}"

                elif "aerotermia" in valor_comercial:
                    val_texto = "Mismo equipo o equivalente | Precio aprox mercado: 4200.00€"
                elif "fotovoltaica" in valor_comercial:
                    val_texto = "Mismo equipo o equivalente | Precio aprox mercado: 1850.00€"
                else:
                    # Salvaguarda general por si escribes cualquier otra rama comercial
                    val_texto = "Mismo equipo o equivalente | Precio aprox mercado: 55.00€"

            # --- PRIORIDAD 2: ANÁLISIS DE CÓDIGOS (Columna C vacía) ---
            else:
                # CASO 1: Código Nativo IVE
                if any(ive_ref in codigo_upper for ive_ref in codigos_ive_referencia) or (len(codigo) >= 6 and codigo[0].isdigit() and codigo[1].isalpha()):
                    val_texto = "Código IVE Para precios se deberían de usar de la base de precios del IVE actual."

                # CASO 2: Código Estructura CYPE
                elif any(c in codigo_upper for c in [".", "-", "_"]) or len(codigo) >= 5:
                    val_texto = "Código CYPE Para precios se deberían de usar de la base de precios del IVE, son superiores y más acorde al mercado"
                
                # CASO 4: Inventado o Erróneo
                else:
                    val_texto = "❌ CODIGO ERRONEO / INVENTADO | Cotejar con IVE"

            # Inyección en la celda correspondiente a la derecha
            ws.cell(row=row_idx, column=col_ia_destino, value=val_texto)
            
            resultados_vista.append({
                "Código": codigo,
                "Datos Comerciales (Col C)": valor_comercial if valor_comercial != "" else "—",
                "Resultado REVISIÓN IA": val_texto
            })

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("✅ Estructura reparada. El motor comercial vuelve a asignar los precios estables de tu plantilla.")
        st.dataframe(pd.DataFrame(resultados_vista), use_container_width=True)
        
        st.download_button(
            label="📥 DESCARGAR EXCEL REVISADO (.XLSX)",
            data=output.getvalue(),
            file_name=f"{uploaded_file.name.split('.')[0]}_Revisado_IA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
            
    except Exception as e:
        st.error(f"Error en la ejecución del script: {e}")
