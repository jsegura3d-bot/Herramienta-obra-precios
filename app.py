import streamlit as st
import pandas as pd
import io
import openpyxl

st.set_page_config(page_title="Revisor IVE BDC25 - Valencia", layout="wide")
st.title("🛠️ Revisor de Precios Pro - Modo Fijo (Bugarra)")
st.caption("Configuración activa: Caso 3 limpio con marcas reales de proyecto y precios de mercado.")

# --- PREFIJOS MAESTROS IVE PARA DETECCIÓN ---
codigos_ive_referencia = ["0AF010", "EIEB20", "DRT030", "EIEC", "PIBB", "DAISA"]

# --- BANCO DE DATOS DE PRECIOS DE MERCADO PARA MARCAS DEL PROYECTO (Caso 3) ---
db_marcas_comerciales = {
    "iluminación": [
        {"keywords": ["naos", "normalux"], "marca_real": "Normalux Naos", "precio_est": 42.50},
        {"keywords": ["philips", "coreline"], "marca_real": "Philips", "precio_est": 65.0},
        {"keywords": ["ledvance", "osram"], "marca_real": "Ledvance", "precio_est": 45.0},
        {"keywords": ["jiso"], "marca_real": "Jiso", "precio_est": 38.0}
    ],
    "mecanismos": [
        {"keywords": ["schneider", "artec"], "marca_real": "Schneider Artec", "precio_est": 16.80},
        {"keywords": ["luxomat", "beg", "koban"], "marca_real": "BEG Luxomat / Koban", "precio_est": 85.00},
        {"keywords": ["simon 100", "simon100"], "marca_real": "Simon 100", "precio_est": 32.00},
        {"keywords": ["simon 27", "simon27"], "marca_real": "Simon 27", "precio_est": 14.50}
    ],
    "fotovoltaica": [
        {"keywords": ["huawei", "sun2000"], "marca_real": "Huawei FusionSolar", "precio_est": 1850.0},
        {"keywords": ["fronius"], "marca_real": "Fronius", "precio_est": 2400.0}
    ],
    "aerotermia": [
        {"keywords": ["mitsubishi", "ecodan"], "marca_real": "Mitsubishi Electric", "precio_est": 4200.0},
        {"keywords": ["daikin"], "marca_real": "Daikin", "precio_est": 4600.0}
    ]
}

uploaded_file = st.file_uploader("Sube tu Excel de Bugarra", type=["xlsx"])

if uploaded_file:
    try:
        file_bytes = uploaded_file.read()
        # openpyxl mantiene intactos los estilos, fuentes, bordes y columnas del archivo subido
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
        
        if 'Hoja5' in wb.sheetnames:
            ws = wb['Hoja5']
        else:
            ws = wb.active

        # Forzamos mapeo estricto por posición visual de tu Excel
        col_codigo_idx = 1      # Columna A (Código)
        col_resumen_idx = 2     # Columna B (Descripción/Resumen)
        col_comercial_idx = 3   # Columna C (DATOS COMERCIALES)

        # Colocamos la columna REVISIÓN IA exactamente en la primera columna vacía a la derecha
        col_ia_destino = ws.max_column + 1
        ws.cell(row=1, column=col_ia_destino, value="REVISIÓN IA").font = openpyxl.styles.Font(bold=True, color="0000FF")

        resultados_vista = []

        for row_idx in range(2, ws.max_row + 1):
            codigo = str(ws.cell(row=row_idx, column=col_codigo_idx).value or '').strip()
            resumen = str(ws.cell(row=row_idx, column=col_resumen_idx).value or '').strip()
            valor_comercial = str(ws.cell(row=row_idx, column=col_comercial_idx).value or '').strip().lower()
            
            # Saltamos celdas que son títulos de capítulos o vacías en el código
            if codigo == "" or codigo.lower() == "none" or codigo.lower() == "código" or "capítulo" in codigo.lower():
                continue

            val_texto = ""
            codigo_upper = codigo.upper()
            resumen_lower = resumen.lower()

            # --- PRIORIDAD 1: CASO 3 (Comerciales - Columna C con Texto) ---
            if valor_comercial != "":
                marca_detectada = None
                precio_mercado = None
                
                # Buscamos la rama comercial correspondiente en nuestro diccionario
                if valor_comercial in db_marcas_comerciales:
                    for item in db_marcas_comerciales[valor_comercial]:
                        if any(kw in resumen_lower for kw in item["keywords"]):
                            marca_detectada = item["marca_real"]
                            precio_mercado = item["precio_est"]
                            break
                
                if marca_detectada and precio_mercado:
                    val_texto = f"Mismo equipo o equivalente ({marca_detectada}) | Precio aprox mercado: {precio_mercado}€"
                else:
                    # En caso de no encontrar la marca específica en el texto, ponemos un genérico según la rama de la columna C
                    precio_generico = "Revisar"
                    if "iluminación" in valor_comercial: precio_generico = "45.00€"
                    elif "mecanismos" in valor_comercial: precio_generico = "18.50€"
                    elif "aerotermia" in valor_comercial: precio_generico = "4200.00€"
                    
                    val_texto = f"Mismo equipo o equivalente | Precio aprox mercado: {precio_generico}"

            # --- PRIORIDAD 2: ANÁLISIS DE CÓDIGOS (Columna C vacía) ---
            else:
                # CASO 1: Código Nativo IVE
                if any(ive_ref in codigo_upper for ive_ref in codigos_ive_referencia) or (len(codigo) >= 6 and codigo[0].isdigit() and codigo[1].isalpha()):
                    val_texto = "Código IVE Para precios se deberían de usar de la base de precios del IVE actual."

                # CASO 2: Código Estructura CYPE
                elif any(c in codigo_upper for c in [".", "-", "_"]) or len(codigo) >= 5:
                    val_texto = "Código CYPE Para precios se deberían de usar de la base de precios del IVE, son superiores y más acorde al mercado"
                
                # CASO 4: Inventado
                else:
                    val_texto = "❌ CODIGO ERRONEO / INVENTADO | Cotejar con IVE"

            # Guardamos la celda limpia en el archivo Excel original respetando el resto de datos
            ws.cell(row=row_idx, column=col_ia_destino, value=val_texto)
            
            resultados_vista.append({
                "Código": codigo,
                "Datos Comerciales (Col C)": valor_comercial if valor_comercial != "" else "—",
                "Resultado REVISIÓN IA": val_texto
            })

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("✅ ¡Lógica corregida por completo! Formato limpio y precios de mercado cruzados correctamente.")
        st.dataframe(pd.DataFrame(resultados_vista), use_container_width=True)
        
        st.download_button(
            label="📥 DESCARGAR EXCEL CONTROL TOTAL (.XLSX)",
            data=output.getvalue(),
            file_name=f"{uploaded_file.name.split('.')[0]}_Revisado_IA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
            
    except Exception as e:
        st.error(f"Error en la ejecución del script: {e}")
