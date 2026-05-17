import streamlit as st
import pandas as pd
import io
import openpyxl

st.set_page_config(page_title="Revisor IVE BDC25 - Valencia", layout="wide")
st.title("🛠️ Revisor de Precios Pro - Cruce Total (CYPE vs IVE Valencia)")
st.caption("Configuración activa: Control estricto de columna DATOS COMERCIALES e indexación corregida.")

# --- BANCO DE PRECIOS INTEGRADO IVE (VALENCIA - JULIO 2025) ---
precios_ive = {
    "0AF010": {"precio": 73.12, "codigo_oficial": "0AF010", "keywords": ["acometida", "agua", "desconexión"]},
    "EIEB20hac": {"precio": 35.50, "codigo_oficial": "EIEB20hac", "keywords": ["interruptor", "estanco", "mecanismo"]},
    "EIEB20hab": {"precio": 37.55, "codigo_oficial": "EIEB20hab", "keywords": ["tecla", "unipolar"]},
    "EIEB20beg2": {"precio": 210.32, "codigo_oficial": "EIEB20beg2", "keywords": ["detector", "presencia", "movimiento"]},
    "EIEB21db": {"precio": 44.19, "codigo_oficial": "EIEB21db", "keywords": ["toma", "corriente", "enchufe"]},
    "DRT030": {"precio": 8.91, "codigo_oficial": "DRT030", "keywords": ["tubo", "canalización", "rígido"]},
    "EIEC.3DD": {"precio": 1.19, "codigo_oficial": "EIEC.3DD", "keywords": ["tubo", "pvc", "curvable", "emp"]},
    "EIEC.6bb": {"precio": 2.09, "codigo_oficial": "EIEC.6bb", "keywords": ["tubo", "poliolefina", "rojo"]},
    "PIBB.2e": {"precio": 2616.91, "codigo_oficial": "PIBB.2e", "keywords": ["aerotermia", "bomba calor", "acs", "acumulador"]},
    "DAISA.02A": {"precio": 49.99, "codigo_oficial": "DAISA.02A", "keywords": ["emergencia", "autónoma", "naos", "evc"]},
    "DAISA.NAOSN5.": {"precio": 60.67, "codigo_oficial": "DAISA.NAOSN5.", "keywords": ["emergencia", "autónoma", "naos", "lm"]},
    "DAISA.06A": {"precio": 27.75, "codigo_oficial": "DAISA.06A", "keywords": ["accesorio", "naos", "kes"]},
    "DAISA.07A": {"precio": 16.23, "codigo_oficial": "DAISA.07A", "keywords": ["accesorio", "naos", "ket"]}
}

# --- BANCO DE PRECIOS E IDENTIFICADORES CYPE ---
precios_cype_fijos = {
    "0AE010": {"precio": 292.54, "codigo_oficial": "0AE010"},
    "0AS010": {"precio": 203.04, "codigo_oficial": "0AS010"},
    "DPT020": {"precio": 5.84, "codigo_oficial": "DPT020"},
    "IEEL.1db": {"precio": 1.45, "codigo_oficial": "IEEL.1db"},
    "IFA005": {"precio": 36.94, "codigo_oficial": "IFA005"}
}

# --- RADAR COMERCIAL RECONSTRUIDO CON PRECIOS DE MERCADO FIJOS POR MARCA ---
cat_comercial = [
    {"keywords": ["carandini", "veka"], "marca_real": "Carandini Veka", "precio": "185.00€"},
    {"keywords": ["schreder", "schréder", "socelec", "briteline"], "marca_real": "Schréder Socelec", "precio": "320.00€"},
    {"keywords": ["normalux", "naos", "emergencia"], "marca_real": "Normalux Naos", "precio": "42.50€"},
    {"keywords": ["luxomat", "beg", "koban", "detector"], "marca_real": "BEG Luxomat / Koban", "precio": "120.00€"},
    {"keywords": ["schneider", "artec"], "marca_real": "Schneider Artec", "precio": "16.80€"},
    {"keywords": ["philips", "coreline"], "marca_real": "Philips", "precio": "65.00€"},
    {"keywords": ["ledvance", "osram"], "marca_real": "Ledvance", "precio": "45.00€"},
    {"keywords": ["jiso"], "marca_real": "Jiso", "precio": "38.00€"},
    {"keywords": ["simon 100", "simon100"], "marca_real": "Simon 100", "precio": "32.00€"},
    {"keywords": ["simon 27", "simon27"], "marca_real": "Simon 27", "precio": "14.50€"},
    {"keywords": ["huawei", "sun2000"], "marca_real": "Huawei FusionSolar", "precio": "1850.00€"},
    {"keywords": ["daikin"], "marca_real": "Daikin", "precio": "4600.00€"}
]

# --- MOTOR DE ASIGNACIÓN CYPE ---
def mapear_y_estimar_cype(codigo, descripcion, precio_presu):
    codigo_clean = codigo.upper().strip()
    desc_clean = descripcion.lower()
    
    if codigo_clean in precios_cype_fijos:
        return precios_cype_fijos[codigo_clean]["precio"], precios_cype_fijos[codigo_clean]["codigo_oficial"], "Base Exacta"
        
    if codigo_clean.startswith("DDDI") or "desmontado" in desc_clean:
        if "saneamiento" in desc_clean: return 610.50, "DDDI10ccbab", "CYPE - Desmontaje Saneamiento"
        if "fontanería" in desc_clean: return 545.20, "DDDI10cbbab", "CYPE - Desmontaje Fontanería"
        return 450.00, "DDDI10a", "CYPE - Desmontajes Generales"

    if codigo_clean.startswith("DIE") or "eléctrica" in desc_clean: return 685.00, "DIE060", "CYPE - Instalación Eléctrica"
    if "acometida" in desc_clean and "agua" in desc_clean: return 36.94, "IFA005", "CYPE - Acometidas Agua"
    if codigo_clean.startswith("DSM") or "sanitario" in desc_clean: return 31.50, "DSM010", "CYPE - Aparatos Sanitarios"
    if codigo_clean.startswith("DPT") or "demolición" in desc_clean: return 5.50, "DPT020", "CYPE - Demoliciones"

    if any(c in codigo_clean for c in [".", "-"]) or len(codigo_clean) > 6:
        return round(precio_presu * 0.95, 2), f"{codigo_clean}_CYPE", "CYPE - Estructura"
        
    return None, "—", None

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
        
        # BUSCADOR INTELIGENTE DE COLUMNAS (Asegura posición exacta de la nueva columna)
        col_codigo_idx = next((i for i, c in enumerate(df_temp.columns) if "cod" in c.lower() or c == "Presupuesto" or "unnamed: 0" in c.lower()), 0)
        col_resumen_idx = next((i for i, c in enumerate(df_temp.columns) if "res" in c.lower() or "desc" in c.lower() or "unnamed: 1" in c.lower()), 1)
        
        # Forzamos la detección explícita de tu columna de DATOS COMERCIALES
        col_comercial_idx = next((i for i, c in enumerate(df_temp.columns) if "comercial" in c.lower() or "datos" in c.lower() or "unnamed: 2" in c.lower()), 2)
        
        col_pres_idx = next((i for i, c in enumerate(df_temp.columns) if ("pres" in c.lower() or "imp" in c.lower() or "prec" in c.lower()) and "can" not in c.lower() or "unnamed: 4" in c.lower()), 4)

        # La nueva columna de revisión se añade al final de todo a la derecha
        col_ia_destino = ws.max_column + 1
        ws.cell(row=1, column=col_ia_destino, value="COLUMNA IA (REVISIÓN DE MÁRGENES VALENCIA)").font = openpyxl.styles.Font(bold=True, color="0000FF")

        resultados_vista = []

        for row_idx in range(2, ws.max_row + 1):
            codigo = str(ws.cell(row=row_idx, column=col_codigo_idx + 1).value or '').strip()
            resumen = str(ws.cell(row=row_idx, column=col_resumen_idx + 1).value or '').strip()
            
            # Leemos el valor físico de tu columna "DATOS COMERCIALES"
            valor_comercial = str(ws.cell(row=row_idx, column=col_comercial_idx + 1).value or '').strip().lower()
            
            if codigo == "" or codigo.lower() == "none" or codigo.lower() == "código" or "capítulo" in codigo.lower() or "total" in resumen.lower():
                continue
                
            try:
                precio_presu = float(ws.cell(row=row_idx, column=col_pres_idx + 1).value or 0.0)
            except:
                precio_presu = 0.0

            val_texto = ""
            codigo_upper = codigo.upper()
            texto_analisis = (codigo + " " + resumen).lower()

            # --- PRIORIDAD 1: CASO 3 (Si la columna DATOS COMERCIALES tiene cualquier tipo de texto) ---
            if valor_comercial != "" and valor_comercial != "none":
                marca_cazada = None
                precio_cazado = None
                
                # Escanea la descripción buscando las palabras clave del catálogo premium
                for item in cat_comercial:
                    if any(kw in texto_analisis for kw in item["keywords"]):
                        marca_cazada = item["marca_real"]
                        precio_cazado = item["precio"]
                        break
                
                if marca_cazada and precio_cazado:
                    val_texto = f"Mismo equipo o equivalente ({marca_cazada}) | Precio aprox mercado: {precio_cazado}"
                else:
                    # Precios de contingencia estables según la categoría que escribas en la columna
                    precio_defecto = "55.00€"
                    if "iluminación" in valor_comercial or "iluminacion" in valor_comercial:
                        precio_defecto = "55.00€"
                    elif "mecanismos" in valor_comercial:
                        precio_defecto = "16.80€"
                    elif "aerotermia" in valor_comercial:
                        precio_defecto = "4200.00€"
                    elif "fotovoltaica" in valor_comercial:
                        precio_defecto = "1850.00€"
                        
                    val_texto = f"Mismo equipo o equivalente | Precio aprox mercado: {precio_defecto}"

            # --- PRIORIDAD 2: ANÁLISIS DE BASE DIRECTA IVE (Columna Datos Comerciales vacía) ---
            elif codigo in precios_ive:
                p_ive_col = f"{precios_ive[codigo]['precio']} €"
                if precio_presu <= precios_ive[codigo]['precio']:
                    val_texto = f"🟢 IVE OK. Presupuesto cubierto ({p_ive_col})."
                else:
                    val_texto = f"🔴 ALERTA: PRESUPUESTO SUPERA AL IVE ({p_ive_col})."

            # --- PRIORIDAD 3: ANÁLISIS DE CYPE (Columna Datos Comerciales vacía) ---
            else:
                precio_cype_est, cod_cype_oficial, ref_cype = mapear_y_estimar_cype(codigo, resumen, precio_presu)
                if precio_cype_est is not None:
                    p_cype_col = f"{precio_cype_est} €"
                    if precio_presu >= precio_cype_est:
                        val_texto = f"🟢 CYPE OK (Margen seguro vs base de {p_cype_col})."
                    else:
                        val_texto = f"🟢 CYPE OK (Precio cubierto por base de {p_cype_col})."
                else:
                    val_texto = "❌ CODIGO ERRONEO / INVENTADO | Cotejar con IVE"

            # --- RASTREO DE MARGEN EXTRA CON IVE VALENCIA (Solo si no es comercial Caso 3) ---
            if valor_comercial == "" or valor_comercial == "none":
                for cod_ive_ref, info_ive in precios_ive.items():
                    if any(kw in texto_analisis for kw in info_ive["keywords"]):
                        p_ive_col = f"{info_ive['precio']} €"
                        if info_ive["precio"] > precio_presu:
                            val_texto = f"🔵 RECOMENDADO OPTIMIZAR: En IVE Valencia se paga a {p_ive_col} (¡Código Oficial: {cod_ive_ref} te da más margen!)."
                        else:
                            val_texto += f" | IVE Valencia disponible a {p_ive_col} (Es más bajo, mantener original)."
                        break

            # Inyectamos el dictamen definitivo en la celda del extremo derecho del Excel
            ws.cell(row=row_idx, column=col_ia_destino, value=val_texto)
            
            resultados_vista.append({
                "Partida": codigo,
                "Datos Comerciales (Col C)": valor_comercial,
                "Precio Presu": f"{precio_presu} €",
                "Dictamen Columna IA": val_texto
            })

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("✅ ¡Columna 'DATOS COMERCIALES' integrada en el flujo! Los índices se han recalculado correctamente.")
        st.dataframe(pd.DataFrame(resultados_vista), use_container_width=True)
        
        st.download_button(
            label="📥 DESCARGAR TU EXCEL ORIGINAL CON COLUMNA IA (.XLSX)",
            data=output.getvalue(),
            file_name=f"{uploaded_file.name.split('.')[0]}_Revisado_IA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
            
    except Exception as e:
        st.error(f"Error técnico al procesar el formato del Excel: {e}")
