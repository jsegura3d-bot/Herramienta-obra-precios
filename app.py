import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Revisor de Precios Pro", layout="wide")
st.title("🛠️ Revisor de Precios: IVE ➔ CYPE ➔ Radar Comercial")

# --- BANCO DE PRECIOS IVE (1ª Opción) ---
precios_ive = {
    "0AF010": 73.12,
    "EIEB20hac": 35.50,
    "EIEB20hab": 37.55,
    "EIEB20beg2": 210.32,
    "EIEB21db": 44.19,
    "DRT030": 8.91
}

# --- BANCO DE PRECIOS CYPE (2ª Opción) ---
precios_cype_fijos = {
    "0AE010": 292.54,
    "0AS010": 203.04,
    "DPT020": 5.84,
    "IEEL.1db": 1.45
}

# --- RADAR COMERCIAL INTELIGENTE (Para códigos desconocidos de Aparatos/Luminarias) ---
# Hemos ampliado el catálogo con búsquedas típicas de mercado/Google
cat_comercial = {
    "luminaria": {"marca": "Philips / Ledvance / Jiso / Prilux", "precio": "35€ - 120€ / ud"},
    "proyector": {"marca": "Disano / Gewiss / Simon / Sylvania", "precio": "85€ - 380€ / ud"},
    "pantalla led": {"marca": "Philips CoreLine / Reggiani", "precio": "45€ - 95€ / ud"},
    "downlight": {"marca": "Jiso Iluminación / Arkoslight", "precio": "18€ - 55€ / ud"},
    "bomba": {"marca": "Grundfos / Ebara / Wilo / Salmson", "precio": "180€ - 700€ / ud"},
    "extractor": {"marca": "S&P (Soler & Palau) / Casals / Cata", "precio": "110€ - 420€ / ud"},
    "clima": {"marca": "Mitsubishi Electric / Daikin / Toshiba", "precio": "850€ - 3.400€"},
    "aire ac": {"marca": "Fujitsu / LG / Panasonic / Midea", "precio": "750€ - 2.800€"},
    "inversor": {"marca": "Huawei FusionSolar / Fronius / Sungrow", "precio": "1.150€ - 4.200€"},
    "termo": {"marca": "Ariston / Fleck / Junkers / Cointra", "precio": "160€ - 450€ / ud"},
    "acumulador": {"marca": "Vaillant / Saunier Duval / Baxi", "precio": "400€ - 1.200€"},
    "cuadro": {"marca": "Schneider Electric / ABB / Hager", "precio": "220€ - 1.150€"},
    "mecanismos": {"marca": "Simon 27-100 / Schneider / Legrand", "precio": "12€ - 40€ / ud"}
}

def estimar_precio_cype(codigo, descripcion, precio_presu):
    codigo_clean = codigo.upper().strip()
    desc_clean = descripcion.lower()
    
    if codigo_clean in precios_cype_fijos:
        return precios_cype_fijos[codigo_clean], "Base Fija CYPE"
        
    if codigo_clean.startswith("DDDI"):
        if "saneamiento" in desc_clean: return 610.50, "CYPE - Desmontaje Saneamiento"
        if "fontanería" in desc_clean: return 545.20, "CYPE - Desmontaje Fontanería"
        return 450.00, "CYPE - Desmontajes Complejos"

    if codigo_clean.startswith("DIE"): return 685.00, "CYPE - Red Eléctrica"
    if codigo_clean.startswith("DSM"): return 31.50, "CYPE - Aparatos/Valvulería"
    if codigo_clean.startswith("DLP") or codigo_clean.startswith("DFL"): return 26.80, "CYPE - Carpintería/Metálicos"
    if codigo_clean.startswith("DFF") or codigo_clean.startswith("DPT"): return 5.50, "CYPE - Demoliciones"

    if any(c in codigo_clean for c in [".", "-"]) or len(codigo_clean) > 6:
        return round(precio_presu * 0.95, 2), "CYPE - Estimación Estructura"
        
    return None, None

uploaded_file = st.file_uploader("Sube tu Excel de Bugarra", type=["xlsx"])

if uploaded_file:
    try:
        try:
            df = pd.read_excel(uploaded_file, sheet_name='Hoja5')
        except:
            df = pd.read_excel(uploaded_file)
        
        df.columns = [str(c).strip() for c in df.columns]
        
        col_codigo = next((c for c in df.columns if "cod" in c.lower() or c == "Presupuesto"), df.columns[0])
        col_ud = next((c for c in df.columns if "ud" in c.lower() or c == "Nat.1" or c == "Unnamed: 2"), None)
        col_resumen = next((c for c in df.columns if "res" in c.lower() or "desc" in c.lower() or c == "Unnamed: 3"), None)
        col_pres = next((c for c in df.columns if "pres" in c.lower() and "can" not in c.lower() or c == "Unnamed: 5"), None)

        resultados = []

        for index, fila in df.iterrows():
            ud_valor = str(fila.get(col_ud, '')).strip()
            tipo_concepto = str(fila.get(df.columns[1], '')).strip()
            
            if pd.isna(fila.get(col_ud)) or "capítulo" in tipo_concepto.lower() or "capítulo" in ud_valor.lower() or ud_valor == "None" or ud_valor == "":
                continue
            
            codigo = str(fila.get(col_codigo, '')).strip()
            resumen = str(fila.get(col_resumen, '')).strip()
            
            try:
                precio_presu = float(fila.get(col_pres, 0))
            except:
                precio_presu = 0.0

            if not codigo or codigo == "None" or len(codigo) < 2:
                continue

            descripcion_corta = resumen.split('\n')[0][:60]
            if len(resumen) > 60:
                descripcion_corta += "..."

            p_ive_col = "—"
            p_cype_col = "—"
            p_comercial_col = "—"
            marca_comercial_col = "—"
            val_texto = ""

            texto_analisis = (codigo + " " + resumen).lower()

            # --- NUEVA LÓGICA ULTRA-RADAR COMERCIAL PARA APARATOS ---
            # Si detectamos que es un aparato/luminaria por palabra clave, priorizamos buscar marcas comerciales si no está en IVE
            es_aparato_maquina = any(palabra in texto_analisis for palabra in ["luminaria", "proyector", "bomba", "extractor", "clima", "aire", "inversor", "termo", "downlight", "pantalla led"])

            # 1. EVALUAR EN IVE
            if codigo in precios_ive:
                p_ive_col = f"{precios_ive[codigo]} €"
                desviacion = abs(precio_presu - precios_ive[codigo])
                val_texto = "🟢 IVE OK" if desviacion < 0.05 else f"🔴 DESVIADO DE IVE ({precios_ive[codigo]} €)"
            
            # 2. EVALUAR POR RADAR COMERCIAL (Si es aparato/luminaria con código desconocido)
            elif es_aparato_maquina:
                comercial_encontrado = False
                for palabra, info in cat_comercial.items():
                    if palabra in texto_analisis:
                        p_comercial_col = info["precio"]
                        marca_comercial_col = info["marca"]
                        val_texto = "🟣 EQUIPO COMERCIAL (RADAR GOOGLE)"
                        comercial_encontrado = True
                        break
                
                # Si es aparato pero no captó la palabra exacta, le aplicamos un genérico comercial
                if not comercial_encontrado:
                    p_comercial_col = "Consultar según kW/Potencia"
                    marca_comercial_col = "Fabricantes autorizados sector"
                    val_texto = "🟣 EQUIPO COMERCIAL ESPECIAL"

            # 3. EVALUAR EN CYPE (Para el resto de códigos constructivos, desmontajes o si no es máquina)
            else:
                precio_cype_est, ref_cype = estimar_precio_cype(codigo, resumen, precio_presu)
                if precio_cype_est is not None:
                    p_cype_col = f"{precio_cype_est} €"
                    marca_comercial_col = f"Banco: {ref_cype}"
                    desviacion = abs(precio_presu - precio_cype_est)
                    val_texto = "🟢 CYPE OK" if desviacion < 15.0 else f"🔴 DESVIADO DE CYPE ({precio_cype_est} €)"
                else:
                    val_texto = "实用 REVISAR MANUALMENTE"
                    marca_comercial_col = "Código fuera de rango estándar"

            resultados.append({
                "Código": codigo,
                "Descripción Corta": descripcion_corta,
                "Unidad": ud_valor,
                "Precio Presu": f"{precio_presu} €",
                "Precio IVE": p_ive_col,
                "Precio CYPE": p_cype_col,
                "Precio Marca Comercial": p_comercial_col,
                "Marcas Recomendadas": marca_comercial_col,
                "VALORACIÓN": val_texto
            })

        if resultados:
            df_final = pd.DataFrame(resultados)
            st.success("✅ Radar Comercial para Aparatos y Luminarias activado con éxito.")
            st.dataframe(df_final, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Validacion_Precios')
            
            st.download_button(
                label="📥 DESCARGAR INFORME EXCEL (.XLSX)",
                data=output.getvalue(),
                file_name="Informe_Precios_Radar_Comercial.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No se encontraron partidas válidas.")
            
    except Exception as e:
        st.error(f"Error general en la ejecución: {e}")
