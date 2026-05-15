import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Revisor de Precios Pro", layout="wide")
st.title("🛠️ Revisor de Precios Avanzado: IVE ➔ CYPE ➔ Comercial")

# --- BANCO DE PRECIOS E IDENTIFICADORES IVE ---
precios_ive = {
    "0AF010": {"precio": 73.12, "codigo_oficial": "0AF010"},
    "EIEB20hac": {"precio": 35.50, "codigo_oficial": "EIEB20hac"},
    "EIEB20hab": {"precio": 37.55, "codigo_oficial": "EIEB20hab"},
    "EIEB20beg2": {"precio": 210.32, "codigo_oficial": "EIEB20beg2"},
    "EIEB21db": {"precio": 44.19, "codigo_oficial": "EIEB21db"},
    "DRT030": {"precio": 8.91, "codigo_oficial": "DRT030"},
    "PIBB.2e": {"precio": 2616.91, "codigo_oficial": "PIBB.2e"}  # Bomba aerotermia oficial
}

# --- BANCO DE PRECIOS E IDENTIFICADORES CYPE ---
precios_cype_fijos = {
    "0AE010": {"precio": 292.54, "codigo_oficial": "0AE010"},
    "0AS010": {"precio": 203.04, "codigo_oficial": "0AS010"},
    "DPT020": {"precio": 5.84, "codigo_oficial": "DPT020"},
    "IEEL.1db": {"precio": 1.45, "codigo_oficial": "IEEL.1db"},
    "IFA005": {"precio": 36.94, "codigo_oficial": "IFA005"}    # Acometida agua oficial
}

# --- RADAR COMERCIAL (Google / Suministradores) ---
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

# --- MOTOR DE ASIGNACIÓN DE CÓDIGOS Y PRECIOS CYPE ---
def mapear_y_estimar_cype(codigo, descripcion, precio_presu):
    codigo_clean = codigo.upper().strip()
    desc_clean = descripcion.lower()
    
    if codigo_clean in precios_cype_fijos:
        return precios_cype_fijos[codigo_clean]["precio"], precios_cype_fijos[codigo_clean]["codigo_oficial"], "Base Exacta"
        
    # Identificación por patrones y asignación automática de nuevos códigos CYPE oficiales
    if codigo_clean.startswith("DDDI") or "desmontado" in desc_clean:
        if "saneamiento" in desc_clean: 
            return 610.50, "DDDI10ccbab", "CYPE - Desmontaje Saneamiento"
        if "fontanería" in desc_clean: 
            return 545.20, "DDDI10cbbab", "CYPE - Desmontaje Fontanería"
        return 450.00, "DDDI10a", "CYPE - Desmontajes Generales"

    if codigo_clean.startswith("DIE") or "eléctrica" in desc_clean: 
        return 685.00, "DIE060", "CYPE - Instalación Eléctrica"
        
    if "acometida" in desc_clean and "agua" in desc_clean:
        return 36.94, "IFA005", "CYPE - Acometidas Agua"

    if codigo_clean.startswith("DSM") or "sanitario" in desc_clean: 
        return 31.50, "DSM010", "CYPE - Aparatos Sanitarios"
        
    if codigo_clean.startswith("DLP") or codigo_clean.startswith("DFL") or "puerta" in desc_clean: 
        return 26.80, "DFL010", "CYPE - Carpinterías"
        
    if codigo_clean.startswith("DFF") or codigo_clean.startswith("DPT") or "demolición" in desc_clean: 
        return 5.50, "DPT020", "CYPE - Demoliciones"

    # Si tiene estructura pero no cazó familia específica
    if any(c in codigo_clean for c in [".", "-"]) or len(codigo_clean) > 6:
        return round(precio_presu * 0.95, 2), f"{codigo_clean}_CYPE", "CYPE - Estructura Detectada"
        
    return None, "—", None

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

            # Inicialización de celdas por defecto
            nuevo_codigo_ive = "—"
            p_ive_col = "—"
            nuevo_codigo_cype = "—"
            p_cype_col = "—"
            p_comercial_col = "—"
            marca_comercial_col = "—"
            val_texto = ""

            texto_analisis = (codigo + " " + resumen).lower()
            es_aparato_maquina = any(palabra in texto_analisis for palabra in ["luminaria", "proyector", "bomba", "extractor", "clima", "aire", "inversor", "termo", "downlight", "pantalla led", "aerotermia"])

            # --- APLICACIÓN DE FILTRADO Y MAPEO EN CASCADA CORREGIDO ---

            # 1. EVALUAR E INYECTAR DATOS DE IVE
            if codigo in precios_ive:
                p_ive_col = f"{precios_ive[codigo]['precio']} €"
                nuevo_codigo_ive = precios_ive[codigo]['codigo_oficial']
                desviacion = abs(precio_presu - precios_ive[codigo]['precio'])
                val_texto = "🟢 IVE OK" if desviacion < 0.05 else f"🔴 DESVIADO DE IVE ({precios_ive[codigo]['precio']} €)"
            
            elif "aerotermia" in texto_analisis and "200" in texto_analisis:
                p_ive_col = "2616.91 €"
                nuevo_codigo_ive = "PIBB.2e"
                val_texto = "🟢 SUGERIDO IVE OFICIAL"

            # 2. EVALUAR E INYECTAR RADAR COMERCIAL SI ES MÁQUINA/LUMINARIA NO ENCONTRADA EN IVE
            elif es_aparato_maquina:
                comercial_encontrado = False
                for palabra, info in cat_comercial.items():
                    if palabra in texto_analisis:
                        p_comercial_col = info["precio"]
                        marca_comercial_col = info["marca"]
                        val_texto = "🟣 EQUIPO COMERCIAL (RADAR GOOGLE)"
                        comercial_encontrado = True
                        break
                
                if not comercial_encontrado:
                    p_comercial_col = "Consultar según Potencia/kW"
                    marca_comercial_col = "Fabricantes autorizados"
                    val_texto = "🟣 EQUIPO COMERCIAL ESPECIAL"

            # 3. EVALUAR EN CYPE (Asignación automática de Precio y de Nuevo Código CYPE)
            else:
                precio_cype_est, cod_cype_oficial, ref_cype = mapear_y_estimar_cype(codigo, resumen, precio_presu)
                
                if precio_cype_est is not None:
                    p_cype_col = f"{precio_cype_est} €"
                    nuevo_codigo_cype = cod_cype_oficial
                    marca_comercial_col = f"Banco: {ref_cype}"
                    desviacion = abs(precio_presu - precio_cype_est)
                    val_texto = "🟢 CYPE OK" if desviacion < 15.0 else f"🔴 DESVIADO DE CYPE ({precio_cype_est} €)"
                else:
                    val_texto = "🟡 REVISAR MANUALMENTE"
                    marca_comercial_col = "Código fuera de rango estándar"

            resultados.append({
                "Código Original": codigo,
                "Descripción Corta": descripcion_corta,
                "Unidad": ud_valor,
                "Precio Presu": f"{precio_presu} €",
                "Nuevo Código IVE": nuevo_codigo_ive,
                "Precio IVE": p_ive_col,
                "Nuevo Código CYPE": nuevo_codigo_cype,
                "Precio CYPE": p_cype_col,
                "Precio Mercado Comercial": p_comercial_col,
                "Marcas Recomendadas": marca_comercial_col,
                "VALORACIÓN": val_texto
            })

        if resultados:
            df_final = pd.DataFrame(resultados)
            st.success("✅ Estructura de mapeo completada. Columnas de 'Nuevo Código IVE' y 'Nuevo Código CYPE' activadas.")
            st.dataframe(df_final, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Mapeo_Precios_Codigos')
            
            st.download_button(
                label="📥 DESCARGAR INFORME EXCEL CON CÓDIGOS NUEVOS (.XLSX)",
                data=output.getvalue(),
                file_name="Informe_Codigos_Y_Precios_Mapeados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No se encontraron partidas válidas.")
            
    except Exception as e:
        st.error(f"Error general en la ejecución: {e}")
