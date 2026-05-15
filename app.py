import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Revisor de Precios Pro", layout="wide")
st.title("🛠️ Revisor Avanzado de Instalaciones (Filtro por Partidas Reales)")

# --- BASE DE DATOS EXTENDIDA (IVE) ---
precios_ive = {
    "0AF010": {"precio": 73.12, "ref": "IVE 2025 - Desconexión acometida agua"},
    "EIEB20hac": {"precio": 35.50, "ref": "IVE 2025 - Interruptor unipolar estanco"},
    "EIEB20hab": {"precio": 37.55, "ref": "IVE 2025 - Interruptor unipolar tecla grn"},
    "EIEB20beg2": {"precio": 210.32, "ref": "IVE 2025 - Detector presencia y control"},
    "EIEB21db": {"precio": 44.19, "ref": "IVE 2025 - Toma corriente 10/16 A estanca"},
    "DRT030": {"precio": 8.91, "ref": "IVE 2025 - Canalización rígida"}
}

# --- BASE DE DATOS DE RESPALDO (CYPE) ---
precios_cype = {
    "0AE010": {"precio": 292.54, "ref": "CYPE - Desconexión acometida eléctrica"},
    "0AS010": {"precio": 203.04, "ref": "CYPE - Desconexión acometida saneamiento"},
    "DPT020": {"precio": 5.84, "ref": "CYPE - Democión partición ladrillo hueco"},
    "IEEL.1db": {"precio": 1.45, "ref": "CYPE - Pequeño material"}
}

# --- INTELIGENCIA DE MERCADO (Caso 4) ---
inteligencia_mercado = {
    "bomba": {"marcas": "Grundfos, Ebara, Wilo", "rango": "150€ - 600€ según caudal"},
    "ascensor": {"marcas": "Otis, ThyssenKrupp, Orona", "rango": "18.000€ - 25.000€"},
    "inversor": {"marcas": "Huawei, Fronius, Sungrow", "rango": "1.200€ - 4.500€ según kW"},
    "clima": {"marcas": "Daitsu, Mitsubishi, Toshiba", "rango": "800€ - 3.500€ VRF/Splits"},
    "mecanismos": {"marcas": "Schneider, Legrand, Simon", "rango": "12€ - 45€ gama media"},
    "acometida": {"marcas": "Material certificado homologado Iberdrola/Aguas", "rango": "Aprox 150€ - 400€ por actuación"}
}

uploaded_file = st.file_uploader("Sube tu Excel de Bugarra", type=["xlsx"])

if uploaded_file:
    try:
        # Cargamos el archivo buscando la Hoja5 o la que esté disponible
        try:
            df = pd.read_excel(uploaded_file, sheet_name='Hoja5')
        except:
            df = pd.read_excel(uploaded_file)
        
        # Forzar nombres limpios temporalmente para localizar datos de forma inteligente
        df.columns = [str(c).strip() for c in df.columns]
        
        # Encontrar las columnas dinámicamente por aproximación de nombre
        col_codigo = next((c for c in df.columns if "cod" in c.lower() or c == "Presupuesto"), df.columns[0])
        col_ud = next((c for c in df.columns if "ud" in c.lower() or c == "Nat.1" or c == "Unnamed: 2"), None)
        col_resumen = next((c for c in df.columns if "res" in c.lower() or "desc" in c.lower() or c == "Unnamed: 3"), None)
        col_pres = next((c for c in df.columns if "pres" in c.lower() and "can" not in c.lower() or c == "Unnamed: 5"), None)

        resultados = []

        for index, fila in df.iterrows():
            # FILTRO CRÍTICO: Solo actuamos si la celda tiene una unidad válida (Ud, m², m, kg...)
            # E ignoramos explícitamente si dice "Capítulo"
            ud_valor = str(fila.get(col_ud, '')).strip()
            tipo_concepto = str(fila.get(df.columns[1], '')).strip() # Suele ser la columna 'Nat' o tipo de partida
            
            if pd.isna(fila.get(col_ud)) or "capítulo" in tipo_concepto.lower() or "capítulo" in ud_valor.lower():
                continue # Saltamos la fila de títulos o capítulos automáticamente
            
            codigo = str(fila.get(col_codigo, '')).strip()
            resumen = str(fila.get(col_resumen, '')).strip()
            
            # Extraer precio unitario del presupuesto
            try:
                precio_presu = float(fila.get(col_pres, 0))
            except:
                precio_presu = 0.0

            if not codigo or codigo == "None" or len(codigo) < 2:
                continue

            # --- CASCADA DE DECISIÓN DE INGENIERÍA ---
            codigo_ref = "N/A"
            precio_oficial = "—"
            val_texto = ""
            base_origen = ""
            info_comercial = "Partida tipificada en banco de precios"

            # 1. ¿Está en IVE? (Manda el IVE)
            if codigo in precios_ive:
                base_origen = "IVE"
                precio_oficial = precios_ive[codigo]["precio"]
                codigo_ref = precios_ive[codigo]["ref"]
                desviacion = abs(precio_presu - precio_oficial)
                val_texto = "🟢 PRECIO CORRECTO (IVE OK)" if desviacion < 0.05 else f"🔴 DESVIADO (Dif: {round(precio_presu - precio_oficial, 2)} €)"
            
            # 2. ¿No está en IVE pero está en CYPE?
            elif codigo in precios_cype:
                base_origen = "CYPE"
                precio_oficial = precios_cype[codigo]["precio"]
                codigo_ref = precios_cype[codigo]["ref"]
                desviacion = abs(precio_presu - precio_oficial)
                val_texto = "🟢 PRECIO CORRECTO (CYPE OK)" if desviacion < 0.05 else f"🔴 DESVIADO (Dif: {round(precio_presu - precio_oficial, 2)} €)"
            
            # 3. No está en ningún banco oficial -> Mercado Comercial (Caso 4) o Inventado
            else:
                mercado_detectado = False
                texto_analisis = (codigo + " " + resumen).lower()
                for palabra, info in inteligencia_mercado.items():
                    if palabra in texto_analisis:
                        base_origen = "WEB / MERCADO"
                        codigo_ref = "No tipificado en bancos oficiales"
                        val_texto = "🟣 CASO 4: EQUIPO COMERCIAL"
                        info_comercial = f"Marcas: {info['marcas']} | Rango aprox mercado: {info['rango']}"
                        mercado_detectado = True
                        break
                
                if not mercado_detectado:
                    base_origen = "CYPE / INVENTADO"
                    codigo_ref = "Código no encontrado en maestros estándar"
                    val_texto = "🟡 CÓDIGO INVENTADO O REVISAR DESCRIPCIÓN"
                    info_comercial = "Verificar procedencia con el proyectista"

            resultados.append({
                "Código": codigo,
                "Unidad": ud_valor,
                "Base Evaluada": base_origen,
                "Código Ref. Origen": codigo_ref,
                "Precio Presu": f"{precio_presu} €" if precio_presu > 0 else "0.00 €",
                "Precio Oficial Banco": f"{precio_oficial} €" if isinstance(precio_oficial, (int, float)) else precio_oficial,
                "VALORACIÓN": val_texto,
                "Ejemplos y Marcas Comerciales": info_comercial
            })

        if resultados:
            df_final = pd.DataFrame(resultados)
            st.success("✅ Análisis optimizado aplicado exclusivamente a partidas con unidades.")
            st.dataframe(df_final, use_container_width=True)
            
            # Generador del Excel limpio
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Validacion_Final_Obra')
            
            st.download_button(
                label="📥 DESCARGAR INFORME EXCEL LIMPIO (.XLSX)",
                data=output.getvalue(),
                file_name="Informe_Precios_Filtrado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No se encontraron filas válidas con unidades de medida para analizar.")
            
    except Exception as e:
        st.error(f"Error en el procesado: {e}")
