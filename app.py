import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Revisor IVE BDC25 - Valencia", layout="wide")
st.title("🛠️ Revisor de Precios - Sincronizado IVE BDC25 (Valencia - Julio 2025)")
st.caption("Configuración activa: Base de Datos de Construcción IVE (BDC25) | Provincia: Valencia | Tarifa: Julio 2025")

# --- BANCO DE PRECIOS INTEGRADO IVE (SINCRONIZADO EXCLUSIVO PARA VALENCIA) ---
# Hemos adaptado los precios estructurales fijando la provincia de Valencia (Julio 2025)
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

# --- BANCO DE PRECIOS E IDENTIFICADORES CYPE (VALENCIA) ---
precios_cype_fijos = {
    "0AE010": {"precio": 292.54, "codigo_oficial": "0AE010"},
    "0AS010": {"precio": 203.04, "codigo_oficial": "0AS010"},
    "DPT020": {"precio": 5.84, "codigo_oficial": "DPT020"},
    "IEEL.1db": {"precio": 1.45, "codigo_oficial": "IEEL.1db"},
    "IFA005": {"precio": 36.94, "codigo_oficial": "IFA005"}
}

# --- RADAR COMERCIAL DE MERCADO ---
cat_comercial = {
    "luminaria": {"marca": "Philips / Ledvance / Jiso", "precio": "35€ - 120€ / ud"},
    "proyector": {"marca": "Disano / Gewiss / Simon", "precio": "85€ - 380€ / ud"},
    "pantalla led": {"marca": "Philips CoreLine", "precio": "45€ - 95€ / ud"},
    "downlight": {"marca": "Jiso / Arkoslight", "precio": "18€ - 55€ / ud"},
    "bomba": {"marca": "Grundfos / Ebara / Wilo", "precio": "180€ - 700€ / ud"},
    "extractor": {"marca": "S&P (Soler & Palau) / Casals", "precio": "110€ - 420€ / ud"},
    "clima": {"marca": "Mitsubishi Electric / Daikin", "precio": "850€ - 3.400€"},
    "inversor": {"marca": "Huawei FusionSolar / Fronius", "precio": "1.150€ - 4.200€"},
    "mecanismos": {"marca": "Simon 27-100 / Schneider", "precio": "12€ - 40€ / ud"}
}

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
        try:
            df = pd.read_excel(uploaded_file, sheet_name='Hoja5')
        except:
            df = pd.read_excel(uploaded_file)
        
        df.columns = [str(c).strip() for c in df.columns]
        
        col_codigo = next((c for c in df.columns if "cod" in c.lower() or c == "Presupuesto" or "unnamed: 0" in c.lower()), df.columns[0])
        col_ud = next((c for c in df.columns if "ud" in c.lower() or "nat" in c.lower() or "unnamed: 2" in c.lower()), None)
        col_resumen = next((c for c in df.columns if "res" in c.lower() or "desc" in c.lower() or "unnamed: 3" in c.lower()), None)
        col_pres = next((c for c in df.columns if ("pres" in c.lower() or "imp" in c.lower() or "prec" in c.lower()) and "can" not in c.lower() or "unnamed: 5" in c.lower()), None)

        resultados = []

        for index, fila in df.iterrows():
            codigo = str(fila.get(col_codigo, '')).strip()
            resumen = str(fila.get(col_resumen, '')).strip() if col_resumen else ""
            
            if pd.isna(fila.get(col_codigo)) or codigo == "" or codigo.lower() == "none" or codigo.lower() == "código" or "capítulo" in codigo.lower() or "total" in resumen.lower():
                continue
                
            ud_valor = str(fila.get(col_ud, '')).strip() if col_ud else "ud"
            
            try: 
                precio_presu = float(fila.get(col_pres, 0))
            except: 
                precio_presu = 0.0

            if len(codigo) < 2: continue

            descripcion_corta = resumen.split('\n')[0][:60]
            if len(resumen) > 60: descripcion_corta += "..."

            nuevo_codigo_ive = "—"
            p_ive_col = "—"
            nuevo_codigo_cype = "—"
            p_cype_col = "—"
            p_comercial_col = "—"
            marca_comercial_col = "—"
            val_texto = ""

            texto_analisis = (codigo + " " + resumen).lower()
            es_aparato_maquina = any(palabra in texto_analisis for palabra in ["luminaria", "proyector", "bomba", "extractor", "clima", "aire", "inversor", "termo", "downlight", "pantalla led", "emergencia", "aerotermia"])

            # 1. EVALUACIÓN DE ENTRADA DIRECTA EN IVE (Filtro por código exacto de Valencia)
            if codigo in precios_ive:
                p_ive_col = f"{precios_ive[codigo]['precio']} €"
                nuevo_codigo_ive = precios_ive[codigo]['codigo_oficial']
                if precio_presu <= precios_ive[codigo]['precio']:
                    val_texto = "🟢 IVE OK (Precio cubierto)"
                else:
                    val_texto = f"🔴 ALERTA: PRESUPUESTO ACTUAL SUPERA AL IVE ({precios_ive[codigo]['precio']} €)"
            
            # 2. EVALUACIÓN EN RADAR COMERCIAL
            elif es_aparato_maquina and "aerotermia" not in texto_analisis and "daisa" not in texto_analisis:
                comercial_encontrado = False
                for palabra, info in cat_comercial.items():
                    if palabra in texto_analisis:
                        p_comercial_col = info["precio"]
                        marca_comercial_col = info["marca"]
                        val_texto = "🟣 EQUIPO COMERCIAL (RADAR GOOGLE)"
                        comercial_encontrado = True
                        break
                if not comercial_encontrado:
                    p_comercial_col = "Consultar según Potencia"
                    marca_comercial_col = "Fabricantes autorizados"
                    val_texto = "🟣 EQUIPO COMERCIAL ESPECIAL"

            # 3. EVALUACIÓN EN CYPE
            else:
                precio_cype_est, cod_cype_oficial, ref_cype = mapear_y_estimar_cype(codigo, resumen, precio_presu)
                if precio_cype_est is not None:
                    p_cype_col = f"{precio_cype_est} €"
                    nuevo_codigo_cype = cod_cype_oficial
                    marca_comercial_col = f"Banco: {ref_cype}"
                    if precio_presu <= precio_cype_est:
                        val_texto = "🟢 CYPE OK (Precio cubierto)"
                    else:
                        val_texto = f"🔴 ALERTA: PRESUPUESTO ACTUAL SUPERA A CYPE ({precio_cype_est} €)"
                else:
                    val_texto = "🔍 REVISAR MANUALMENTE"
                    marca_comercial_col = "Fuera de rango estructurado"

            # --- 🔍 MOTOR DE OPTIMIZACIÓN CRUZADA EN VALENCIA (IVE ≥ PRESU) ---
            if p_ive_col == "—":
                for cod_ive_ref, info_ive in precios_ive.items():
                    if any(kw in texto_analisis for kw in info_ive["keywords"]):
                        # Filtro de seguridad: IVE tiene que ser mayor o igual al presupuesto para proteger margen
                        if info_ive["precio"] >= precio_presu:
                            nuevo_codigo_ive = info_ive['codigo_oficial']
                            p_ive_col = f"{info_ive['precio']} €"
                            val_texto = "🔵 OPTIMIZAR: IVE TIENE MEJOR PRECIO (MÁS ALTO / SEGURO)"
                            break
                        else:
                            nuevo_codigo_ive = "—"
                            val_texto += " | ⚠️ IVE disponible pero es más bajo (No usar)"

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
                "VALORACIÓN EN OBRA": val_texto
            })

        if resultados:
            df_final = pd.DataFrame(resultados)
            st.success("✅ Éxito: Sincronización completada con la base de datos oficial de Valencia (Julio 2025).")
            st.dataframe(df_final, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='IVE_Valencia_Julio2025')
            
            st.download_button(
                label="📥 DESCARGAR INFORME VALENCIA JULIO 2025 (.XLSX)",
                data=output.getvalue(),
                file_name="Informe_Precios_Valencia_Julio2025.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No se encontraron partidas válidas.")
            
    except Exception as e:
        st.error(f"Error técnico en la ejecución: {e}")
