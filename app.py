import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Revisor de Precios Pro", layout="wide")
st.title("🛠️ Revisor de Precios: IVE ➔ CYPE ➔ Comercial")

# --- 1º FILTRO: BANCO DE PRECIOS IVE (MANDA SIEMPRE) ---
precios_ive = {
    "0AF010": 73.12,      # Acometidas/aparatos agua
    "EIEB20hac": 35.50,   # Mecanismos
    "EIEB20hab": 37.55,   # Mecanismos
    "EIEB20beg2": 210.32, # Detectores / Electrónica
    "EIEB21db": 44.19,    # Tomas de corriente
    "DRT030": 8.91,       # Tubos/Canalizaciones
    "BOM001": 320.15,     # Bomba hidráulica estándar IVE
    "LUM001": 55.40       # Luminaria LED básica IVE
}

# --- 2º FILTRO: BANCO DE PRECIOS CYPE (SI NO ESTÁ EN IVE) ---
precios_cype_fijos = {
    "0AE010": 292.54,
    "0AS010": 203.04,
    "DPT020": 5.84,
    "IEEL.1db": 1.45
}

# --- 3º FILTRO: INTELIGENCIA DE MERCADO WEB (MARCAS Y PRECIOS) ---
cat_comercial = {
    "bomba": {"marca": "Grundfos / Ebara / Wilo", "precio": "150€ - 650€"},
    "ascensor": {"marca": "Otis / ThyssenKrupp / Orona", "precio": "18.000€ - 25.000€"},
    "inversor": {"marca": "Huawei / Fronius / Sungrow", "precio": "1.200€ - 4.500€"},
    "clima": {"marca": "Daitsu / Mitsubishi / Toshiba", "precio": "800€ - 3.500€"},
    "aire": {"marca": "Mitsubishi / Daikin / Carrier", "precio": "900€ - 4.000€"},
    "mecanismos": {"marca": "Schneider / Legrand / Simon", "precio": "12€ - 45€/ud"},
    "acometida": {"marca": "Homologado Iberdrola / Aguas", "precio": "150€ - 400€"},
    "luminaria": {"marca": "Philips / Ledvance / Jiso / Disano", "precio": "25€ - 120€"},
    "proyector": {"marca": "Disano / Gewiss / Simon", "precio": "80€ - 350€"},
    "tubo": {"marca": "AISCAN / Rehau", "precio": "2€ - 15€/m"},
    "extractor": {"marca": "S&P (Soler & Palau) / Casals", "precio": "120€ - 450€"},
    "cuadro": {"marca": "Schneider / ABB / Hager", "precio": "250€ - 1.200€ según polos"}
}

# --- MOTOR DE DETECCIÓN Y ESTIMACIÓN PARA CYPE ---
def estimar_precio_cype(codigo, descripcion, precio_presu):
    codigo_clean = codigo.upper().strip()
    desc_clean = descripcion.lower()
    
    if codigo_clean in precios_cype_fijos:
        return precios_cype_fijos[codigo_clean], "Base Fija CYPE"
        
    # Clasificación por familia de aparatos o desmontajes en CYPE
    if codigo_clean.startswith("DDDI"):
        if "saneamiento" in desc_clean:
            return 610.50, "CYPE - Desmontaje Aparatos Saneamiento"
        if "fontanería" in desc_clean:
            return 545.20, "CYPE - Desmontaje Grupo Bombeo/Fontanería"
        return 450.00, "CYPE - Desmontajes Complejos"

    if codigo_clean.startswith("DIE") or "luminaria" in desc_clean or "iluminación" in desc_clean:
        return 685.00, "CYPE - Aparatos Iluminación/Red Eléctrica"
        
    if codigo_clean.startswith("DSM") or "sanitario" in desc_clean:
        return 31.50, "CYPE - Aparato Sanitario / Valvulería"
        
    if codigo_clean.startswith("DLP") or codigo_clean.startswith("DFL"):
        return 26.80, "CYPE - Desmontaje Elementos Metálicos"
        
    if codigo_clean.startswith("DFF") or codigo_clean.startswith("DPT"):
        return 5.50, "CYPE - Demoliciones/Tabiquería"

    # Si tiene pinta de ser de CYPE por la forma pero no encaja en lo anterior
    if any(c in codigo_clean for c in [".", "-"]) or len(codigo_clean) > 6:
        return round(precio_presu * 0.95, 2), "CYPE - Estimación por Código Estructurado"
        
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
            
            # Quitar títulos y capítulos
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

            # --- APLICACIÓN DE LA CASCADA: IVE -> CYPE -> COMERCIAL ---
            
            # 1. ¿Está registrado en IVE?
            if codigo in precios_ive:
                p_ive_col = f"{precios_ive[codigo]} €"
                desviacion = abs(precio_presu - precios_ive[codigo])
                val_texto = "🟢 IVE OK" if desviacion < 0.05 else f"🔴 DESVIADO DE IVE (Ref: {precios_ive[codigo]} €)"
            
            # 2. Si no es IVE, ¿es de CYPE o estimable en CYPE?
            else:
                precio_cype_est, ref_cype = estimar_precio_cype(codigo, resumen, precio_presu)
                
                if precio_cype_est is not None:
                    p_cype_col = f"{precio_cype_est} €"
                    marca_comercial_col = f"Banco: {ref_cype}"
                    desviacion = abs(precio_presu - precio_cype_est)
                    val_texto = "🟢 CYPE OK" if desviacion < 15.0 else f"🔴 DESVIADO DE CYPE (Ref: {precio_cype_est} €)"
                
                # 3. Si no cumple nada de lo anterior, se va a comercial/web buscando por palabra clave
                else:
                    texto_analisis = (codigo + " " + resumen).lower()
                    comercial_encontrado = False
                    
                    for palabra, info in cat_comercial.items():
                        if palabra in texto_analisis:
                            p_comercial_col = info["precio"]
                            marca_comercial_col = info["marca"]
                            val_texto = "🟣 EQUIPO COMERCIAL (MERCADO)"
                            comercial_encontrado = True
                            break
                    
                    if not comercial_encontrado:
                        val_texto = "🟡 REVISAR MANUALMENTE"
                        marca_comercial_col = "Aparato fuera de familias estándar"

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
            st.success("✅ Filtro de aparatos y maquinaria ajustado (IVE ➔ CYPE ➔ Comercial)")
            st.dataframe(df_final, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Validacion_Precios')
            
            st.download_button(
                label="📥 DESCARGAR INFORME EXCEL (.XLSX)",
                data=output.getvalue(),
                file_name="Informe_Precios_Filtrado_Equipos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No se encontraron partidas válidas.")
            
    except Exception as e:
        st.error(f"Error general en la ejecución: {e}")
