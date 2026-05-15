import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Revisor de Precios Pro", layout="wide")
st.title("🛠️ Revisor de Precios: IVE ➔ CYPE ➔ Comercial")

# --- BASE DE DATOS EXTENDIDA (IVE) ---
precios_ive = {
    "0AF010": 73.12,
    "EIEB20hac": 35.50,
    "EIEB20hab": 37.55,
    "EIEB20beg2": 210.32,
    "EIEB21db": 44.19,
    "DRT030": 8.91
}

# --- BASE DE DATOS DE RESPALDO (CYPE) ---
precios_cype = {
    "0AE010": 292.54,
    "0AS010": 203.04,
    "DPT020": 5.84,
    "IEEL.1db": 1.45
}

# --- DICCIONARIO PARA BÚSQUEDA COMERCIAL ---
cat_comercial = {
    "bomba": {"marca": "Grundfos / Ebara / Wilo", "precio": "150€ - 650€"},
    "ascensor": {"marca": "Otis / ThyssenKrupp / Orona", "precio": "18.000€ - 25.000€"},
    "inversor": {"marca": "Huawei / Fronius / Sungrow", "precio": "1.200€ - 4.500€"},
    "clima": {"marca": "Daitsu / Mitsubishi / Toshiba", "precio": "800€ - 3.500€"},
    "mecanismos": {"marca": "Schneider / Legrand / Simon", "precio": "12€ - 45€/ud"},
    "acometida": {"marca": "Homologado Iberdrola / Aguas", "precio": "150€ - 400€"},
    "luminaria": {"marca": "Philips / Ledvance / Jiso", "precio": "25€ - 120€"},
    "proyector": {"marca": "Disano / Gewiss / Simon", "precio": "80€ - 350€"},
    "tubo": {"marca": "AISCAN / Rehau", "precio": "2€ - 15€/m"}
}

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
            
            # Filtrar capítulos y títulos
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

            # Columnas limpias que me pediste
            p_ive_col = "—"
            p_cype_col = "—"
            p_comercial_col = "—"
            marca_comercial_col = "—"
            val_texto = ""

            # 1. PASO: BUSCAR EN IVE
            if codigo in precios_ive:
                p_ive_col = f"{precios_ive[codigo]} €"
                desviacion = abs(precio_presu - precios_ive[codigo])
                val_texto = "🟢 IVE OK" if desviacion < 0.05 else "🔴 DESVIADO DE IVE"
            
            # 2. PASO: SI NO ESTÁ EN IVE, BUSCAR EN CYPE
            elif codigo in precios_cype:
                p_cype_col = f"{precios_cype[codigo]} €"
                desviacion = abs(precio_presu - precios_cype[codigo])
                val_texto = "🟢 CYPE OK" if desviacion < 0.05 else "🔴 DESVIADO DE CYPE"
            
            # 3. PASO: CASO 4 - MERCADO COMERCIAL
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
                    val_texto = "🟡 CÓDIGO DESCONOCIDO / REVISAR MANUALMENTE"
                    marca_comercial_col = "Revisar catálogo comercial específico"

            resultados.append({
                "Código": codigo,
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
            st.success("✅ Análisis finalizado sin errores de estructura.")
            st.dataframe(df_final, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Validacion_Precios')
            
            st.download_button(
                label="📥 DESCARGAR INFORME EXCEL (.XLSX)",
                data=output.getvalue(),
                file_name="Informe_Precios_IVE_CYPE_Comercial.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No se encontraron partidas con unidades válidas para analizar.")
            
    except Exception as e:
        st.error(f"Error general en la ejecución: {e}")
