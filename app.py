import streamlit as st
import pandas as pd

st.title("🛠️ Revisor de Instalaciones (IVE BDC25)")
st.write("Sube tu Excel y añadiré la valoración a la derecha.")

# Base de datos de prueba
precios_ive = {"0AF010": 73.12, "DRT030": 8.91, "PIEM10cb": 115.20}

uploaded_file = st.file_uploader("Elige tu archivo Excel", type=["xlsx"])

if uploaded_file:
    try:
        # Intentamos leer la Hoja5, si no existe, leemos la primera
        try:
            df = pd.read_excel(uploaded_file, sheet_name='Hoja5')
        except:
            df = pd.read_excel(uploaded_file)
        
        # BUSCADOR DE COLUMNAS (Para que no de error si cambia una letra o tilde)
        col_codigo = next((c for c in df.columns if "cod" in c.lower()), None)
        col_pres = next((c for c in df.columns if "pres" in c.lower()), None)
        col_resumen = next((c for c in df.columns if "res" in c.lower() or "desc" in c.lower()), None)

        if col_codigo and col_pres:
            def valorar(fila):
                cod = str(fila[col_codigo]).strip()
                p_presu = fila[col_pres]
                if cod in precios_ive:
                    p_oficial = precios_ive[cod]
                    return f"(ok) 🟢" if abs(p_presu - p_oficial) < 0.1 else f"(IVE: {p_oficial}€) 🔴"
                return "(Revisar / Comercial) ⚪"

            df['VALORACIÓN FINAL'] = df.apply(valorar, axis=1)
            
            st.success("¡Análisis listo!")
            st.dataframe(df)
            
            # EL INFORME EN EXCEL (Lo que querías)
            # Usamos formato CSV para que sea descarga directa y fácil
            csv = df.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="📥 DESCARGAR EXCEL CON VALORACIONES",
                data=csv,
                file_name="informe_precios_revisado.csv",
                mime="text/csv",
            )
        else:
            st.error(f"No encuentro las columnas necesarias. Tu Excel tiene: {list(df.columns)}")
            
    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")
