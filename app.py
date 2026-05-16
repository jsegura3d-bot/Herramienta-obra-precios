import streamlit as st
import pandas as pd
import io
import openpyxl

st.set_page_config(page_title="Revisor IVE BDC25 - Valencia", layout="wide")
st.title("🛠️ Revisor de Precios Pro - Rastreador de Marcas (Bugarra)")
st.caption("Configuración activa: El Caso 3 detecta marcas específicas en la descripción y calcula precios equivalentes.")

# --- PREFIJOS MAESTROS IVE PARA DETECCIÓN ---
codigos_ive_referencia = ["0AF010", "EIEB20", "DRT030", "EIEC", "PIBB", "DAISA"]

# --- DICCIONARIO AVANZADO DE MARCAS Y PRECIOS OBJETIVOS (Caso 3) ---
# Si en la columna C pones la rama, el sistema buscará estas sub-marcas dentro de la descripción (Columna B)
db_marcas_comerciales = {
    "iluminación": [
        {"keywords": ["philips", "coreline"], "marca_real": "Philips", "precio_est": 65.0},
        {"keywords": ["ledvance", "osram"], "marca_real": "Ledvance", "precio_est": 45.0},
        {"keywords": ["jiso"], "marca_real": "Jiso Iluminación", "precio_est": 38.0}
    ],
    "mecanismos": [
        {"keywords": ["simon 100", "simon100"], "marca_real": "Simon 100", "precio_est": 32.0},
        {"keywords": ["simon 27", "simon27"], "marca_real": "Simon 27", "precio_est": 14.5},
        {"keywords": ["schneider", "asfora"], "marca_real": "Schneider Electric", "precio_est": 18.0}
    ],
    "fotovoltaica": [
        {"keywords": ["huawei", "sun2000"], "marca_real": "Huawei FusionSolar", "precio_est": 1850.0},
        {"keywords": ["fronius", "symo"], "marca_real": "Fronius", "precio_est": 2400.0},
        {"keywords": ["longi"], "marca_real": "Longi Solar (Paneles)", "precio_est": 145.0}
    ],
    "aerotermia": [
        {"keywords": ["mitsubishi", "ecodan"], "marca_real": "Mitsubishi Electric", "precio_est": 4200.0},
        {"keywords": ["daikin", "altherma"], "marca_real": "Daikin Altherma", "precio_est": 4600.0},
        {"keywords": ["vaillant", "arotherm"], "marca_real": "Vaillant", "precio_est": 3900.0}
    ]
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
        
        # --- POSICIONES DE COLUMNAS FIJAS ---
        col_codigo_idx = 1      # Columna A (Código)
        col_resumen_idx = 2     # Columna B (Descripción/Resumen)
        col_comercial_idx = 3   # Columna C (DATOS COMERCIALES)
        
        # Localizamos la columna del precio introducido en el presupuesto
        col_precio_idx = next((i + 1 for i, c in enumerate(columnas_limpias) if "pres" in c or "imp" in c or "prec" in c or "val" in c), None)

        # Añadimos la columna de revisión a la derecha
        col_ia_destino = ws.max_column + 1
        ws.cell(row=1, column=col_ia_destino, value="REVISIÓN IA").font = openpyxl.styles.Font(bold=True, color="0000FF")

        resultados_vista = []

        for row_idx in range(2, ws.max_row + 1):
            codigo = str(ws.cell(row=row_idx, column=col_codigo_idx).value or '').strip()
            resumen = str(ws.cell(row=row_idx, column=col_resumen_idx).value or '').strip()
            valor_comercial = str(ws.cell(row=row_idx, column=col_comercial_idx).value or '').strip().lower()
            
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
            resumen_lower = resumen.lower()

            # --- PRIORIDAD 1: CASO 3 (Elementos comerciales con rastreo de marca) ---
            if valor_comercial != "":
                marca_detectada = None
                precio_mercado_estimado = 0.0
                
                # Si la rama existe en nuestra base de datos, buscamos la marca en la descripción (Columna B)
                if valor_comercial in db_marcas_comerciales:
                    for item in db_marcas_comerciales[valor_comercial]:
                        if any(kw in resumen_lower for kw in item["keywords"]):
                            marca_detectada = item["marca_real"]
                            precio_mercado_estimado = item["precio_est"]
                            break
                
                if marca_detectada:
                    # Margen de tolerancia: si tu precio es superior al estimado de mercado, avisa
                    if precio_presu > (precio_mercado_estimado * 1.15):
                        status = "❌ Precio Alto"
                    else:
                        status = "🟢 Precio Correcto"
                        
                    val_texto = f"Mismo equipo o equivalente ({marca_detectada}). Precio aprox mercado: {precio_mercado_estimado}€ | Estado: {status}."
                else:
                    # Si pusiste la rama pero en el texto no se menciona una marca conocida de la lista
                    val_texto = f"Mismo equipo o equivalente (Marca no identificada en texto). Revisar presupuesto estimado para rama: {valor_comercial.upper()}."

            # --- PRIORIDAD 2: ANÁLISIS DE CÓDIGOS (Si la columna C está vacía) ---
            else:
                if any(ive_ref in codigo_upper for ive_ref in codigos_ive_referencia) or (len(codigo) >= 6 and codigo[0].isdigit() and codigo[1].isalpha()):
                    val_texto = "Código IVE Para precios se deberían de usar de la base de precios del IVE actual."

                elif any(c in codigo_upper for c in [".", "-", "_"]) or len(codigo) >= 5:
                    val_texto = "Código CYPE Para precios se deberían de usar de la base de precios del IVE, son superiores y más acorde al mercado"
                
                else:
                    val_texto = "❌ CODIGO ERRONEO / INVENTADO | Cotejar con IVE"

            # Guardamos el resultado en el Excel
            ws.cell(row=row_idx, column=col_ia_destino, value=val_texto)
            
            resultados_vista.append({
                "Código": codigo,
                "Descripción": resumen[:35] + "...",
                "Datos Comerciales": valor_comercial if valor_comercial != "" else "—",
                "Precio Presu": f"{precio_presu} €",
                "Resultado REVISIÓN IA": val_texto
            })

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("✅ ¡Ajustado! El sistema ahora cruza la descripción buscando la marca exacta y calcula su precio aproximado.")
        st.dataframe(pd.DataFrame(resultados_vista), use_container_width=True)
        
        st.download_button(
            label="📥 DESCARGAR EXCEL CON PRECIOS EQUIVALENTES (.XLSX)",
            data=output.getvalue(),
            file_name=f"{uploaded_file.name.split('.')[0]}_Revision_Equivalentes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
            
    except Exception as e:
        st.error(f"Error en la ejecución del script: {e}")
