import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Revisor de Precios Pro", layout="wide")
st.title("🛠️ Revisor de Precios con Filtro de Seguridad de Presupuesto (IVE ≥ Presu)")

# --- BANCO DE PRECIOS E IDENTIFICADORES IVE (MANDA SIEMPRE) ---
precios_ive = {
    "0AF010": {"precio": 73.12, "codigo_oficial": "0AF010", "keywords": ["acometida", "agua", "desconexión"]},
    "EIEM.1ac": {"precio": 35.50, "codigo_oficial": "EIEM.1ac", "keywords": ["interruptor", "estanco", "mecanismo"]},
    "EIEM.1ab": {"precio": 37.55, "codigo_oficial": "EIEM.1ab", "keywords": ["tecla", "unipolar"]},
    "EIEM.5eg2": {"precio": 210.32, "codigo_oficial": "EIEM.5eg2", "keywords": ["detector", "presencia", "movimiento"]},
    "EIEM.2db": {"precio": 44.19, "codigo_oficial": "EIEM.2db", "keywords": ["toma", "corriente", "enchufe"]},
    "EIEC.3dd": {"precio": 8.91, "codigo_oficial": "EIEC.3dd", "keywords": ["tubo", "canalización", "rígido"]},
    "PIBB.2e": {"precio": 2616.91, "codigo_oficial": "PIBB.2e", "keywords": ["aerotermia", "bomba calor", "clima", "acumulador"]},
    "EIEL.led": {"precio": 125.50, "codigo_oficial": "EIEL.led", "keywords": ["luminaria", "led", "proyector", "pantalla"]}
}

# --- BANCO DE PRECIOS E IDENTIFICADORES CYPE ---
precios_cype_fijos = {
    "0AE010": {"precio": 292.54, "codigo_oficial": "0AE010"},
    "0AS010": {"precio": 203.04, "codigo_oficial": "0AS010"},
    "DPT020": {"precio": 5.84, "codigo_oficial": "DPT020"},
    "IFA005": {"precio": 36.94, "codigo_oficial": "IFA005"}
}

# --- RADAR COMERCIAL (Mercado / Google) ---
cat_comercial = {
    "luminaria": {"marca": "Philips / Ledvance / Jiso", "precio": "35€ - 120€ / ud"},
    "proyector": {"marca": "Disano / Gewiss / Simon", "precio": "85€ - 380€ / ud"},
    "pantalla led": {"marca": "Philips CoreLine", "precio": "45€ - 95€ / ud"},
    "bomba": {"marca": "Grundfos / Ebara / Wilo", "precio": "180€ - 700€ / ud"},
    "inversor": {"marca": "Huawei FusionSolar / Fronius", "precio": "1.150€ - 4.200€"}
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
            es_aparato_maquina = any(palabra in texto_analisis for palabra in ["luminaria", "proyector", "bomba", "extractor", "clima", "aire", "inversor", "pantalla led", "emergencia"])

            # 1. EVALUACIÓN DE ENTRADA DIRECTA EN IVE
            if codigo in precios_ive:
                p_ive_col = f"{precios_ive[codigo]['precio']} €"
                nuevo_codigo_ive =
