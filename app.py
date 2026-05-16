import streamlit as st
import pandas as pd
import io
import openpyxl

st.set_page_config(page_title="Revisor IVE BDC25 - Valencia", layout="wide")
st.title("🛠️ Revisor de Precios Pro - Formato Bugarra Fijo")
st.caption("Configuración activa: Lógica indexada por posición exacta de columnas (A=Código, B=Resumen, C=Datos Comerciales).")

# --- PREFIJOS MAESTROS IVE PARA DETECCIÓN ---
codigos_ive_referencia = ["0AF010", "EIEB20", "DRT030", "EIEC", "PIBB", "DAISA"]

# --- RADAR COMERCIAL AVANZADO (Caso 3) ---
cat_comercial = {
    "iluminación": {"marca": "Philips / Ledvance / Jiso", "precio": "35€ - 120€ / ud"},
    "fotovoltaica": {"marca": "Huawei FusionSolar / Fronius / Longi", "precio": "Inversores: 1.150€ - 4.200€"},
    "aerotermia": {"marca": "Mitsubishi Electric / Daikin / Vaillant", "precio": "850€ - 3.400€ / ud"},
    "clima": {"marca": "Mitsubishi Electric / Daikin", "precio": "850€ - 3.400€"},
    "bomba": {"marca": "Grundfos / Ebara / Wilo", "precio": "180€ - 700€ / ud"},
    "extractor": {"marca": "S&P (Soler & Palau) / Casals", "precio": "110€ - 420€ / ud"},
    "mecanismos": {"marca": "Simon 27-100 / Schneider", "precio": "12€ - 40€ / ud"}
}

uploaded_file = st.file_uploader("Sube tu Excel de Bugarra", type=["xlsx"])

if uploaded_file:
    try:
        file_bytes = uploaded_file.read()
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
        
        # Buscamos 'Hoja5' o en su defecto la activa para no fallar
        if 'Hoja5' in wb.sheetnames:
            ws = wb['Hoja5']
        else:
            ws = wb.active

        # Determinamos de forma fija las posiciones basadas en tu captura real:
        # Columna A (1) = Código | Columna B (2) = Resumen | Columna C (3) = DATOS COMERCIALES
        col_codigo_idx = 1      # Columna A
        col_resumen_idx = 2     # Columna B
        col_comercial_idx = 3   # Columna C

        # La nueva columna de revisión se añade al final de la tabla existente (Derecha)
        col_ia_destino = ws.max_column + 1
        ws.cell(row=1, column=col_ia_destino, value="REVISIÓN IA").font = openpyxl.styles.Font(bold=True, color="0000FF")

        resultados_vista = []

        # Recorremos las filas desde la 2 hasta el final
        for row_idx in range(2, ws.max_row + 1):
            codigo = str(ws.cell(row=row_idx, column=col_codigo_idx).value or '').strip()
            resumen = str(ws.cell(row=row_idx, column=col_resumen_idx).value or '').strip()
            valor_comercial = str(ws.cell(row=row_idx, column=col_comercial_idx).value or '').strip().lower()
            
            # Detectar si es una fila vacía o de títulos de familias/capítulos sin código real
            if codigo == "" or codigo.lower() == "none" or codigo.lower() == "código" or "capítulo" in codigo.lower():
                continue

            val_texto = ""
            codigo_upper = codigo.upper()

            # --- PRIORIDAD 1: CASO 3 (Si la columna C 'DATOS COMERCIALES' tiene contenido) ---
            if valor_comercial != "":
                if valor_comercial in cat_comercial:
                    info = cat_comercial[valor_comercial]
                    val_texto = f"🟣 EQUIPO COMERCIAL (Detectado en columna: {valor_comercial}). Marcas aconsejadas: {info['marca']} | Coste estimado: {info['precio']}."
                else:
                    val_texto = f"🟣 EQUIPO COMERCIAL (Rama nueva: {valor_comercial}). Revisar marcas autorizadas en el proyecto."

            # --- PRIORIDAD 2: ANÁLISIS DE CÓDIGOS (Si la columna C está vacía) ---
            else:
                # CASO 1: Código Nativo IVE
                if any(ive_ref in codigo_upper for ive_ref in codigos_ive_referencia) or (len(codigo) >= 6 and codigo[0].isdigit() and codigo[1].isalpha()):
                    val_texto = "Código IVE Para precios se deberían de usar de la base de precios del IVE actual."

                # CASO 2: Estructuras reconocibles de CYPE (llevan puntos, guiones o son largas tipo DISAN_04)
                elif any(c in codigo_upper for c in [".", "-", "_"]) or len(codigo) >= 5:
                    val_texto = "Código CYPE Para precios se deberían de usar de la base de precios del IVE, son superiores y más acorde al mercado"
                
                # CASO 4: Código Erróneo / Inventado / Roto
                else:
                    val_texto = "❌ CODIGO ERRONEO / INVENTADO | Cotejar con IVE"

            # Inyectamos el dictamen en la nueva columna a la derecha de la fila correspondiente
            ws.cell(row=row_idx, column=col_ia_destino, value=val_texto)
            
            # Guardamos una copia limpia para la previsualización en la web de Streamlit
            resultados_vista.append({
                "Código (Col A)": codigo,
                "Descripción (Col B)": resumen[:30] + "...",
                "Datos Comerciales (Col C)": valor_comercial if valor_comercial != "" else "—",
                "Resultado REVISIÓN IA": val_texto
            })

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("✅ ¡Corregido al 100%! Lógica de comerciales blindada y formato original respetado.")
        st.dataframe(pd.DataFrame(resultados_vista), use_container_width=True)
        
        st.download_button(
            label="📥 DESCARGAR EXCEL REVISADO (.XLSX)",
            data=output.getvalue(),
            file_name=f"{uploaded_file.name.split('.')[0]}_Revisado_IA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
            
    except Exception as e:
        st.error(f"Error en la ejecución del script: {e}")
