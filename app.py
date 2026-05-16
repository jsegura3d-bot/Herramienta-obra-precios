import streamlit as st
import pandas as pd
import io
import openpyxl

st.set_page_config(page_title="Revisor IVE BDC25 - Valencia", layout="wide")
st.title("🛠️ Revisor de Precios Pro - Extractor de Marcas (Bugarra)")
st.caption("Configuración activa: Extracción y mapeo dinámico de fabricantes y precios de mercado de forma individual.")

# --- PREFIJOS MAESTROS IVE PARA DETECCIÓN ---
codigos_ive_referencia = ["0AF010", "EIEB20", "DRT030", "EIEC", "PIBB", "DAISA"]

# --- DICCIONARIO INDEPENDIENTE DE DETECCIÓN POR TEXTO ---
# Se evalúa la descripción (Columna B) para asignar marca y precio exacto sin depender de la Columna C
reglas_marcas = [
    {"keywords": ["carandini", "veka"], "marca_real": "Carandini Veka", "precio": "185.00€"},
    {"keywords": ["schreder", "schréder", "socelec", "briteline"], "marca_real": "Schréder Socelec", "precio": "320.00€"},
    {"keywords": ["normalux", "naos"], "marca_real": "Normalux Naos", "precio": "42.50€"},
    {"keywords": ["koban", "luxomat", "beg", "detector de presencia"], "marca_real": "BEG Luxomat / Koban", "precio": "120.00€"},
    {"keywords": ["schneider", "artec"], "marca_real": "Schneider Artec", "precio": "16.80€"},
    {"keywords": ["philips", "coreline"], "marca_real": "Philips", "precio": "65.00€"},
    {"keywords": ["ledvance", "osram"], "marca_real": "Ledvance", "precio": "45.00€"},
    {"keywords": ["jiso"], "marca_real": "Jiso", "precio": "38.00€"},
    {"keywords": ["simon 100", "simon100"], "marca_real": "Simon 100", "precio": "32.00€"},
    {"keywords": ["simon 27", "simon27"], "marca_real": "Simon 27", "precio": "14.50€"},
    {"keywords": ["simon 82", "simon82"], "marca_real": "Simon 82", "precio": "25.00€"},
    {"keywords": ["huawei", "sun2000"], "marca_real": "Huawei FusionSolar", "precio": "1850.00€"},
    {"keywords": ["daikin"], "marca_real": "Daikin", "precio": "4600.00€"}
]

uploaded_file = st.file_uploader("Sube tu Excel de Bugarra", type=["xlsx"])

if uploaded_file:
    try:
        file_bytes = uploaded_file.read()
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
        
        if 'Hoja5' in wb.sheetnames:
            ws = wb['Hoja5']
        else:
            ws = wb.active

        # Índices de columnas fijos según la estructura de tu presupuesto
        col_codigo_idx = 1      # Columna A (Código)
        col_resumen_idx = 2     # Columna B (Descripción/Resumen)
        col_comercial_idx = 3   # Columna C (DATOS COMERCIALES)

        # Generamos la columna REVISIÓN IA a la derecha del todo de forma limpia
        col_ia_destino = ws.max_column + 1
        ws.cell(row=1, column=col_ia_destino, value="REVISIÓN IA").font = openpyxl.styles.Font(bold=True, color="0000FF")

        resultados_vista = []

        for row_idx in range(2, ws.max_row + 1):
            codigo = str(ws.cell(row=row_idx, column=col_codigo_idx).value or '').strip()
            resumen = str(ws.cell(row=row_idx, column=col_resumen_idx).value or '').strip()
            valor_comercial = str(ws.cell(row=row_idx, column=col_comercial_idx).value or '').strip().lower()
            
            # Filtro para ignorar celdas de títulos vacías o filas de capítulos
            if codigo == "" or codigo.lower() == "none" or codigo.lower() == "código" or "capítulo" in codigo.lower():
                continue

            val_texto = ""
            codigo_upper = codigo.upper()
            resumen_lower = resumen.lower()

            # --- CASO 3: ELEMENTOS COMERCIALES (Prioridad si se detecta marca en la descripción o columna C con texto) ---
            marca_detectada = None
            precio_mercado = None

            # Buscamos de forma directa analizando el contenido de la descripción larga (Columna B)
            for regla in reglas_marcas:
                if any(kw in resumen_lower for kw in regla["keywords"]):
                    marca_detectada = regla["marca_real"]
                    precio_mercado = regla["precio"]
                    break

            if marca_detectada and precio_mercado:
                val_texto = f"Mismo equipo o equivalente ({marca_detectada}) | Precio aprox mercado: {precio_mercado}"
            
            elif valor_comercial != "":
                # Si la columna C tiene datos comerciales pero la descripción no incluye una marca de la lista premium
                precio_defecto = "55.00€"
                if "aerotermia" in valor_comercial: precio_defecto = "4200.00€"
                elif "fotovoltaica" in valor_comercial: precio_defecto = "1850.00€"
                elif "iluminación" in valor_comercial or "iluminacion" in valor_comercial:
                    if "emergencia" in resumen_lower or "naos" in resumen_lower: precio_defecto = "42.50€"
                    else: precio_defecto = "55.00€"
                val_texto = f"Mismo equipo o equivalente | Precio aprox mercado: {precio_defecto}"

            # --- CASOS RESTANTES: ANÁLISIS POR CÓDIGO (Columna C vacía y sin marca comercial identificada) ---
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

            # Guardamos el resultado en la celda del Excel manteniendo intacto el resto
            ws.cell(row=row_idx, column=col_ia_destino, value=val_texto)
            
            resultados_vista.append({
                "Código": codigo,
                "Descripción": resumen[:40] + "...",
                "Resultado REVISIÓN IA": val_texto
            })

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("✅ Sistema corregido. Extracción individual de marcas y precios estabilizada.")
        st.dataframe(pd.DataFrame(resultados_vista), use_container_width=True)
        
        st.download_button(
            label="📥 DESCARGAR EXCEL CORREGIDO (.XLSX)",
            data=output.getvalue(),
            file_name=f"{uploaded_file.name.split('.')[0]}_Revisado_IA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
            
    except Exception as e:
        st.error(f"Error en la ejecución del script: {e}")
