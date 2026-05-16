import streamlit as st
import pandas as pd
import io
import openpyxl

st.set_page_config(page_title="Revisor IVE BDC25 - Valencia", layout="wide")
st.title("🛠️ Revisor de Precios Pro - Formato Bugarra Fijo")
st.caption("Configuración activa: Lógica de 4 Casos con control de precios en elementos comerciales (Caso 3).")

# --- PREFIJOS MAESTROS IVE PARA DETECCIÓN ---
codigos_ive_referencia = ["0AF010", "EIEB20", "DRT030", "EIEC", "PIBB", "DAISA"]

# --- RADAR COMERCIAL AVANZADO CON RANGOS NUMÉRICOS PARA VALIDACIÓN (Caso 3) ---
cat_comercial = {
    "iluminación": {"marca": "Philips / Ledvance / Jiso", "precio_min": 35.0, "precio_max": 120.0, "texto_coste": "35€ - 120€ / ud"},
    "fotovoltaica": {"marca": "Huawei FusionSolar / Fronius / Longi", "precio_min": 1150.0, "precio_max": 4200.0, "texto_coste": "Inversores: 1.150€ - 4.200€"},
    "aerotermia": {"marca": "Mitsubishi Electric / Daikin / Vaillant", "precio_min": 850.0, "precio_max": 3400.0, "texto_coste": "850€ - 3.400€ / ud"},
    "clima": {"marca": "Mitsubishi Electric / Daikin", "precio_min": 850.0, "precio_max": 3400.0, "texto_coste": "850€ - 3.400€"},
    "bomba": {"marca": "Grundfos / Ebara / Wilo", "precio_min": 180.0, "precio_max": 700.0, "texto_coste": "180€ - 700€ / ud"},
    "extractor": {"marca": "S&P (Soler & Palau) / Casals", "precio_min": 110.0, "precio_max": 420.0, "texto_coste": "110€ - 420€ / ud"},
    "mecanismos": {"marca": "Simon 27-100 / Schneider", "precio_min": 12.0, "precio_max": 40.0, "texto_coste": "12€ - 40€ / ud"}
}

uploaded_file = st.file_uploader("Sube tu Excel de Bugarra", type=["xlsx"])

if uploaded_file:
    try:
        file_bytes = uploaded_file.read()
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
        
        if 'Hoja5' in wb.sheetnames:
            ws = wb['Hoja5']
        else:
            ws = wb.active

        df_temp = pd.read_excel(io.BytesIO(file_bytes), sheet_name=ws.title)
        columnas_limpias = [str(c).strip().lower() for c in df_temp.columns]
        
        # --- POSICIONES DE COLUMNAS SEGÚN TU ESTRUCTURA ---
        col_codigo_idx = 1      # Columna A (Código)
        col_resumen_idx = 2     # Columna B (Resumen)
        col_comercial_idx = 3   # Columna C (DATOS COMERCIALES)
        
        # Localizamos dinámicamente la columna del precio/presupuesto de tu Excel
        col_precio_idx = next((i + 1 for i, c in enumerate(columnas_limpias) if "pres" in c or "imp" in c or "prec" in c or "val" in c), None)

        # La nueva columna de revisión se añade a la derecha del todo
        col_ia_destino = ws.max_column + 1
        ws.cell(row=1, column=col_ia_destino, value="REVISIÓN IA").font = openpyxl.styles.Font(bold=True, color="0000FF")

        resultados_vista = []

        for row_idx in range(2, ws.max_row + 1):
            codigo = str(ws.cell(row=row_idx, column=col_codigo_idx).value or '').strip()
            resumen = str(ws.cell(row=row_idx, column=col_resumen_idx).value or '').strip()
            valor_comercial = str(ws.cell(row=row_idx, column=col_comercial_idx).value or '').strip().lower()
            
            # Intentamos rascar el precio de la fila. Si no encuentra o está roto, por defecto 0.0
            precio_presu = 0.0
            if col_precio_idx is not None:
                try:
                    precio_presu = float(ws.cell(row=row_idx, column=col_precio_idx).value or 0.0)
                except:
                    precio_presu = 0.0
            
            if codigo == "" or codigo.lower() == "none" or codigo.lower() == "código" or "capítulo" in codigo.lower():
                continue

            val_texto = ""
            codigo_upper = codigo.upper()

            # --- PRIORIDAD 1: CASO 3 (Comerciales con validación de precio) ---
            if valor_comercial != "":
                if valor_comercial in cat_comercial:
                    info = cat_comercial[valor_comercial]
                    
                    # Comprobamos si tu precio se pasa del máximo del mercado
                    if precio_presu > info["precio_max"]:
                        dictamen_precio = "❌ Precio Comercial Alto"
                    else:
                        dictamen_precio = "🟢 COMERCIAL OK"
                        
                    val_texto = f"🟣 EQUIPO COMERCIAL ({dictamen_precio}). Marcas aconsejadas: {info['marca']} | Coste estimado: {info['texto_coste']}."
                else:
                    val_texto = f"🟣 EQUIPO COMERCIAL (Rama nueva: {valor_comercial}). Revisar marcas autorizadas en el proyecto."

            # --- PRIORIDAD 2: ANÁLISIS DE CÓDIGOS IVE / CYPE / INVENTADOS ---
            else:
                # CASO 1: Código Nativo IVE
                if any(ive_ref in codigo_upper for ive_ref in codigos_ive_referencia) or (len(codigo) >= 6 and codigo[0].isdigit() and codigo[1].isalpha()):
                    val_texto = "Código IVE Para precios se deberían de usar de la base de precios del IVE actual."

                # CASO 2: Estructuras reconocibles de CYPE
                elif any(c in codigo_upper for c in [".", "-", "_"]) or len(codigo) >= 5:
                    val_texto = "Código CYPE Para precios se deberían de usar de la base de precios del IVE, son superiores y más acorde al mercado"
                
                # CASO 4: Código Erróneo / Inventado / Roto
                else:
                    val_texto = "❌ CODIGO ERRONEO / INVENTADO | Cotejar con IVE"

            # Inyectamos en la celda final del Excel original (Manteniendo formatos)
            ws.cell(row=row_idx, column=col_ia_destino, value=val_texto)
            
            resultados_vista.append({
                "Código": codigo,
                "Descripción": resumen[:30] + "...",
                "Datos Comerciales": valor_comercial if valor_comercial != "" else "—",
                "Precio Introducido": f"{precio_presu} €",
                "Resultado REVISIÓN IA": val_texto
            })

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("✅ ¡Mejora añadida! Los comerciales ahora evalúan si el precio introducido es correcto.")
        st.dataframe(pd.DataFrame(resultados_vista), use_container_width=True)
        
        st.download_button(
            label="📥 DESCARGAR EXCEL CONTROL TOTAL DE PRECIOS (.XLSX)",
            data=output.getvalue(),
            file_name=f"{uploaded_file.name.split('.')[0]}_Revisado_Precios_IA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
            
    except Exception as e:
        st.error(f"Error en la ejecución del script: {e}")
