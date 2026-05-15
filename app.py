import streamlit as st
import pandas as pd
import urllib.parse

# --- SIMULACIÓN DE BASES DE DATOS (Aquí cargarás tus archivos reales luego) ---
BDC25_IVE = {
    "0AF010": 73.12,
    "DRT030": 8.91,
    "PIEM10cb": 115.20
}

CYPE_BASE = {
    "IEEL.1db": 1.45
}

def analizar_presupuesto(df):
    def aplicar_logica(row):
        codigo = str(row['Código']).strip()
        resumen = str(row['Resumen']).lower()
        precio_pres = row['Pres']

        # CASO 4: CÓDIGOS COMBINADOS (con guion)
        if "-" in codigo:
            partes = codigo.split("-")
            suma_ive = 0
            encontrados = 0
            for p in partes:
                if p in BDC25_IVE:
                    suma_ive += BDC25_IVE[p]
                    encontrados += 1
            if encontrados > 0:
                return f"(ok: SUMA IVE {suma_ive}€)" if abs(precio_pres - suma_ive) < 0.5 else f"(PRECIO IVE SUMADO: {suma_ive}€)"

        # CASO 1 y 2: IVE
        if codigo in BDC25_IVE:
            p_ive = BDC25_IVE[codigo]
            if abs(precio_pres - p_ive) < 0.1: # Margen de 10 céntimos
                return "(ok) 🟢"
            else:
                return f"(PRECIO IVE: {p_ive}€) 🔴"

        # CASO 3: CYPE
        if codigo in CYPE_BASE:
            p_cype = CYPE_BASE[codigo]
            p_ive_sim = BDC25_IVE.get(codigo, "No en IVE")
            return f"[CYPE: {p_cype}€] [IVE: {p_ive_sim}€] 🟠"

        # CASO 4: COMERCIAL / INVENTADO / APARATOS
        palabras_aparatos = ["bomba", "ascensor", "inversor", "clima", "equipo", "grupo"]
        if any(palabra in resumen for palabra in palabras_aparatos):
            query = urllib.parse.quote(f"precio comercial {resumen}")
            link = f"https://www.google.com/search?q={query}"
            return f"(BUSCAR WEB: {resumen}) 🟣"

        return "(Revisar manualmente) ⚪"

    # Aplicar la función a cada fila
    df['Valoración'] = df.apply(aplicar_logica, axis=1)
    return df

# --- INTERFAZ DE USUARIO (Streamlit) ---
st.title("🛠️ Revisor de Instalaciones (IVE BDC25)")
st.write("Sube tu Excel (Hoja5) y obtén la valoración automática.")

uploaded_file = st.file_uploader("Elige tu archivo Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name='Hoja5')
    # Limpiamos filas vacías de códigos
    df = df.dropna(subset=['Código'])
    
    with st.spinner('Analizando precios...'):
        df_final = analizar_presupuesto(df)
        st.success("¡Análisis completado!")
        st.dataframe(df_final) # Muestra la tabla en la web
        
        # Botón para descargar
        csv = df_final.to_csv(index=False).encode('utf-8')
        st.download_button("Descargar Excel Valorado", csv, "presupuesto_revisado.csv", "text/csv")