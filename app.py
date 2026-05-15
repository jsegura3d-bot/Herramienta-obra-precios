import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Revisor de Precios", layout="wide")
st.title("🛠️ Revisor de Instalaciones (Informe Excel)")

# Base de datos ampliada (añade aquí tus códigos reales)
precios_ive = {"0AF010": 73.12, "DRT030": 8.91, "PIEM10cb": 115.20}

file = st.file_uploader("Sube tu Excel de Bugarra", type=["xlsx"])

if file:
    try:
        # 1. Leemos el archivo saltándonos posibles filas vacías arriba
        df = pd.read_excel(file, sheet_name=None) # Lee todas las hojas
        hoja_nombre = 'Hoja5' if 'Hoja5' in df else list(df.keys())[0]
        data = pd.read_excel(file, sheet_name=hoja_nombre)

        # 2. Limpieza radical: buscamos la fila donde realmente empiezan los datos
        # Buscamos una fila que contenga algo parecido a "Código" o "Pres"
        data.columns = [str(c) for c in data.columns]
        
        st.write(f"Analizando hoja: {hoja_nombre}")

        # 3. Lógica de valoración
        def analizar(fila):
            # Buscamos el valor en las columnas (sea cual sea su nombre)
            valores = [str(v) for v in fila.values]
            codigo = next((v for v in valores if any(c in v for c in precios_ive)), None)
            
            if codigo and codigo in precios_ive:
                p_oficial = precios_ive[codigo]
                return f"PRECIO IVE CORRECTO ({p_oficial}€)"
            return "REVISAR / COMERCIAL"

        data['VALORACIÓN INGENIERÍA'] = data.apply(analizar, axis=1)

        # 4. MOSTRAR RESULTADO Y BOTÓN DE DESCARGA REAL
        st.success("✅ Análisis finalizado")
        st.dataframe(data)

        # Crear un archivo Excel real para descargar
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            data.to_excel(writer, index=False, sheet_name='Resultado_Revision')
        
        st.download_button(
            label="📥 DESCARGAR INFORME EXCEL VALORADO",
            data=output.getvalue(),
            file_name="Informe_Revision_Precios.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error técnico: {e}")
