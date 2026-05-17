import streamlit as st
import pandas as pd
import io
import openpyxl

st.set_page_config(page_title="Revisor IVE BDC25 - Valencia", layout="wide")
st.title("🛠️ Revisor de Precios Pro - Cruce por Casos (Bugarra)")
st.caption("Configuración activa: Flujo estricto por Casos (Etiqueta fija 'CODIGO CYPE' integrada en el Caso 2)")

# --- DICCIONARIO MAESTRO IVE (Para detectar si el código pertenece a esta base) ---
codigos_ive_referencia = ["0AF010", "EIEB20", "DRT030", "EIEC", "PIBB", "DAISA"]

# --- BANCO DE PRECIOS E IDENTIFICADORES CYPE FISS ---
precios_cype_fijos = {
    "0AE010": {"precio": 292.54, "codigo_oficial": "0AE010"},
    "0AS010": {"precio": 203.04, "codigo_oficial": "0AS010"},
    "DPT020": {"precio": 5.84, "codigo_oficial": "DPT020"},
    "IEEL.1db": {"precio": 1.45, "codigo_oficial": "IEEL.1db"},
    "IFA005": {"precio": 36.94, "codigo_oficial": "IFA005"}
}

# --- RADAR COMERCIAL AVANZADO (Indexado por tus nuevas etiquetas de cajón) ---
cat_comercial = {
    "iluminación": {"marca": "Philips / Ledvance / Jiso", "precio": "35€ - 120€ / ud"},
    "fotovoltaica": {"marca": "Huawei FusionSolar / Fronius / Longi", "precio": "Inversores: 1.150€ - 4.200€"},
    "aerotermia": {"marca": "Mitsubishi Electric / Daikin / Vaillant", "precio": "850€ - 3.400€ / ud"},
    "clima": {"marca": "Mitsubishi Electric / Daikin", "precio": "850€ - 3.400€"},
    "bomba": {"marca": "Grundfos / Ebara / Wilo", "precio": "180€ - 700€ / ud"},
    "extractor": {"marca": "S&P (Soler & Palau) / Casals", "precio": "110€ - 420€ / ud"},
    "mecanismos": {"marca": "Simon 27-100 / Schneider", "precio": "12€ - 40€ / ud"}
}

# --- MOTOR AUXILIAR ESTIMADOR CYPE ---
def mapear_y_estimar_cype(codigo, descripcion, precio_presu):
    codigo_clean = codigo.upper().strip()
    desc_clean = descripcion.lower()
    
    if codigo_clean in precios_cype_fijos:
        return precios_cype_fijos[codigo_clean]["precio"], True # Código existe en base
        
    if codigo_clean.startswith("DDDI") or "desmontado" in desc_clean:
        if "saneamiento" in desc_clean: return 610.50, True
        if "fontanería" in desc_clean: return 545.20, True
        return 450.00, True

    if codigo_clean.startswith("DIE") or "eléctrica" in desc_clean: return 685.00, True
    if "acometida" in desc_clean and "agua" in desc_clean: return 36.94, True
    if codigo_clean.startswith("DSM") or "sanitario" in desc_clean: return 31.50, True
    if codigo_clean.startswith("DPT") or "demolición" in desc_clean: return 5.50, True

    # Si es una estructura CYPE válida detectada por caracteres pero inventada o modificada
    if any(c in codigo_clean for c in [".", "-"]) or len(codigo_clean) > 6:
        return round(precio_presu * 0.95, 2), False # Código dudoso/mal estructurado
        
    return None, False

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
        df_temp.columns = [str(c).strip() for c in df_temp.columns]
        
        col_codigo_idx = next((i for i, c in enumerate(df_temp.columns) if "cod" in c.lower() or c == "Presupuesto" or "unnamed: 0" in c.lower()), 0)
        col_resumen_idx = next((i for i, c in enumerate(df_temp.columns) if "res" in c.lower() or "desc" in c.lower() or "unnamed: 3" in c.lower()), 2)
        col_pres_idx = next((i for i, c in enumerate(df_temp.columns) if ("pres" in c.lower() or "imp" in c.lower() or "prec" in c.lower()) and "can" not in c.lower() or "unnamed: 5" in c.lower()), 4)

        col_ia_destino = ws.max_column + 1
        ws.cell(row=1, column=col_ia_destino, value="REVISIÓN IA").font = openpyxl.styles.Font(bold=True, color="0000FF")

        resultados_vista = []

        for row_idx in range(2, ws.max_row + 1):
            codigo = str(ws.cell(row=row_idx, column=col_codigo_idx + 1).value or '').strip()
            resumen = str(ws.cell(row=row_idx, column=col_resumen_idx + 1).value or '').strip()
            
            if codigo == "" or codigo.lower() == "none" or codigo.lower() == "código" or "capítulo" in codigo.lower() or "total" in resumen.lower():
                continue
                
            try:
                precio_presu = float(ws.cell(row=row_idx, column=col_pres_idx + 1).value or 0.0)
            except:
                precio_presu = 0.0

            val_texto = ""
            codigo_upper = codigo.upper()
            resumen_lower = resumen.lower()

            # --- CASO 3: DETECCIÓN DE TU ETIQUETA COMERCIAL "DE CAJÓN" ---
            if "comercial_" in resumen_lower:
                sub_rama = resumen_lower.split("comercial_")[1].split()[0].strip()
                
                if sub_rama in cat_comercial:
                    info = cat_comercial[sub_rama]
                    val_texto = f"🟣 EQUIPO COMERCIAL (Etiqueta detectada: {sub_rama}). Marcas aconsejadas: {info['marca']} | Coste estimado: {info['precio']}."
                else:
                    val_texto = f"🟣 EQUIPO COMERCIAL (Rama: {sub_rama}). Revisar marcas autorizadas en el proyecto."

            # --- CASO 1: RECONOCIMIENTO DE CÓDIGO NATIVO IVE ---
            elif any(ive_ref in codigo_upper for ive_ref in codigos_ive_referencia) or (len(codigo) >= 6 and codigo[0].isdigit() and codigo[1].isalpha()):
                val_texto = "🔍 CODIGO IVE REVISAR"

            # --- CASO 2: RECONOCIMIENTO DE CÓDIGO CYPE ---
            else:
                precio_cype_est, codigo_existe_en_base = mapear_y_estimar_cype(codigo, resumen, precio_presu)
                
                if precio_cype_est is not None:
                    # Siempre añadimos el prefijo fijo 'CODIGO CYPE' solicitado
                    if precio_presu >= precio_cype_est:
                        val_texto = "CODIGO CYPE | 🟢 CYPE OK"
                    else:
                        if codigo_existe_en_base:
                            val_texto = "CODIGO CYPE | ❌ Precio mal"
                        else:
                            if precio_presu == 0.0:
                                val_texto = "CODIGO CYPE | ❌ Código mal, Precio mal o Ambas"
                            else:
                                val_texto = "CODIGO CYPE | ❌ Código mal"
                                
                        val_texto += " | Cotejar con IVE"
                else:
                    val_texto = "CODIGO CYPE | 🔍 REVISAR MANUALMENTE | Cotejar con IVE"

            # Inyectamos el dictamen final en la celda del Excel original
            ws.cell(row=row_idx, column=col_ia_destino, value=val_texto)
            
            resultados_vista.append({
                "Código Original": codigo,
                "Descripción Corta": resumen[:50] + "...",
                "Precio Original": f"{precio_presu} €",
                "Resultado Columna IA": val_texto
            })

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("✅ ¡Actualizado! Las partidas de CYPE ahora se clasifican de forma homogénea con la etiqueta 'CODIGO CYPE'.")
        st.dataframe(pd.DataFrame(resultados_vista), use_container_width=True)
        
        st.download_button(
            label="📥 DESCARGAR EXCEL AJUSTADO (.XLSX)",
            data=output.getvalue(),
            file_name=f"{uploaded_file.name.split('.')[0]}_Casos_IA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
            
    except Exception as e:
        st.error(f"Error en la ejecución del script: {e}")
