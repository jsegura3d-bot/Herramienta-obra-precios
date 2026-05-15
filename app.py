import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Revisor de Precios Pro", layout="wide")
st.title("🛠️ Revisor Avanzado: IVE, CYPE y Mercado Comercial")

# --- BASE DE DATOS DE PRECIOS OFICIALES (IVE) ---
precios_ive = {
    "EIEB20hac": {"precio": 35.50, "ref": "IVE 2025 - Mecanismos Estancos"},
    "EIEB20hab": {"precio": 37.55, "ref": "IVE 2025 - Tecla unipolar grn"},
    "EIEB20beg2": {"precio": 210.32, "ref": "IVE 2025 - Detector presencia"},
    "EIEB21db": {"precio": 44.19, "ref": "IVE 2025 - Toma corriente 16A"},
    "0AF010": {"precio": 73.12, "ref": "IVE 2025 - Conductor cobre"},
    "DRT030": {"precio": 8.91, "ref": "IVE 2025 - Canalización rígida"}
}

# --- BASE DE DATOS DE CYPE (Para cuando no encuentra en IVE) ---
precios_cype = {
    "IEEL.1db": {"precio": 1.45, "ref": "CYPE - Pequeño material"},
    "EIEB20hac-CY": {"precio": 33.10, "ref": "CYPE - Alternativa Mecanismo"}
}

# --- DICCIONARIO DE INTELIGENCIA DE MERCADO (Caso 4) ---
inteligencia_mercado = {
    "bomba": {"marcas": "Grundfos, Ebara, Wilo", "rango": "150€ - 600€ según caudal"},
    "ascensor": {"marcas": "Otis, ThyssenKrupp, Orona", "rango": "18.000€ - 25.000€ (6 paradas)"},
    "inversor": {"marcas": "Huawei, Fronius, Sungrow", "rango": "1.200€ - 4.500€ según kW"},
    "clima": {"marcas": "Daitsu, Mitsubishi, Toshiba", "rango": "800€ - 3.500€ VRF/Splits"},
    "mecanismos": {"marcas": "Schneider, Legrand, Simon", "rango": "12€ - 45€ gama media/alta"}
}

uploaded_file = st.file_uploader("Sube tu Excel de Bugarra", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, sheet_name='Hoja5')
        resultados = []
        
        for index, fila in df.iterrows():
            valores_fila = [str(v).strip() for v in fila.values if pd.notna(v)]
            texto_completo_fila = " ".join(valores_fila).lower()
            
            # Identificar códigos en la fila
            codigo_detectado = None
            origen_base = None
            precio_oficial = None
            referencia_codigo = "N/A"
            
            # 1. PASO: BUSCAR EN IVE
            for v in valores_fila:
                if v in precios_ive:
                    codigo_detectado = v
                    precio_oficial = precios_ive[v]["precio"]
                    referencia_codigo = precios_ive[v]["ref"]
                    origen_base = "IVE"
                    break
            
            # 2. PASO: SI NO ESTÁ EN IVE, BUSCAR EN CYPE
            if not codigo_detectado:
                for v in valores_fila:
                    if v in precios_cype or "cy" in v.lower() or len(v) > 10:
                        codigo_detectado = v if v in precios_cype else "Código CYPE"
                        precio_oficial = precios_cype.get(v, {"precio": 0.0})["precio"]
                        referencia_codigo = precios_cype.get(v, {"ref": "Generado por CYPE Ingenieros"})["ref"]
                        origen_base = "CYPE"
                        break

            # Extraer los precios numéricos de la fila para comparar
            precios_en_fila = []
            for v in fila.values:
                try:
                    precios_en_fila.append(float(v))
                except:
                    pass

            # Procesar el resultado según lo encontrado
            if codigo_detectado and origen_base:
                precio_presupuesto = precios_en_fila[0] if precios_en_fila else 0.0
                # Buscar el precio que más se aproxime si hay varios números
                if len(precios_en_fila) > 1:
                    for p in precios_en_fila:
                        if precio_oficial > 0 and abs(p - precio_oficial) < (precio_oficial * 0.5):
                            precio_presupuesto = p
                            break

                desviacion = abs(precio_presupuesto - precio_oficial)
                
                if desviacion < 0.05:
                    val_texto = f"🟢 CORRECTO ({origen_base} OK)"
                else:
                    val_texto = f"🔴 DESVIADO (Dif: {round(precio_presupuesto - precio_oficial, 2)} €)"

                resultados.append({
                    "Código Detectado": codigo_detectado,
                    "Base de Datos": origen_base,
                    "Código Referencia Origen": referencia_codigo,
                    "Precio Presu": f"{precio_presupuesto} €",
                    "Precio Oficial Banco": f"{precio_oficial} €",
                    "VALORACIÓN": val_texto,
                    "Info Mercado / Alternativas": "Usando precio de banco oficial"
                })
                
            # 3. PASO: CASO 4 - NO ESTÁ EN NINGÚN BANCO (MERCADO COMERCIAL)
            else:
                mercado_encontrado = False
                for palabra, info in inteligencia_mercado.items():
                    if palabra in texto_completo_fila:
                        resultados.append({
                            "Código Detectado": "N/A (Comercial)",
                            "Base de Datos": "WEB / MERCADO",
                            "Código Referencia Origen": "No disponible en bancos",
                            "Precio Presu": "Ver original",
                            "Precio Oficial Banco": "—",
                            "VALORACIÓN": "🟣 CASO 4: EQUIPO COMERCIAL",
                            "Info Mercado / Alternativas": f"Marcas: {info['marcas']} | Rango aprox: {info['rango']}"
                        })
                        mercado_encontrado = True
                        break
                
                if not mercado_encontrado:
                    # Fila de texto normal, capítulos o celdas vacías significativas
                    if len(valores_fila) > 2:
                        resultados.append({
                            "Código Detectado": "N/A",
                            "Base de Datos": "REVISIÓN MANUAL",
                            "Código Referencia Origen": "Revisar partida descrita",
                            "Precio Presu": "—",
                            "Precio Oficial Banco": "—",
                            "VALORACIÓN": "⚪ Analizar descripción o texto combinado",
                            "Info Mercado / Alternativas": "Sin coincidencias de palabras clave"
                        })

        if resultados:
            df_informe = pd.DataFrame(resultados)
            st.success("✅ Análisis de ingeniería completado con cruce de tres niveles.")
            st.dataframe(df_informe, use_container_width=True)
            
            # Descarga en Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_informe.to_excel(writer, index=False, sheet_name='Validacion_Precios')
            
            st.download_button(
                label="📥 DESCARGAR INFORME EXCEL COMPLETO (.XLSX)",
                data=output.getvalue(),
                file_name="Informe_Precios_IVE_CYPE_Mercado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Error: {e}")
