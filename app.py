import streamlit as st
import pandas as pd
import io
import openpyxl

st.set_page_config(page_title="Revisor IVE BDC25 - Valencia", layout="wide")
st.title("🛠️ Revisor de Precios Pro - Columna 'datos comerciales' (Bugarra)")
st.caption("Configuración activa: Mapeo directo desde la columna 'datos comerciales' para el Caso 3.")

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
        
        if 'Hoja5' in wb.sheetnames:
            ws = wb['Hoja5']
        else:
            ws = wb.active
            
        df_temp = pd.read_excel(io.BytesIO(file_bytes), sheet_name=ws.title)
        df_temp.columns = [str(c).strip().lower() for c in df_temp.columns]
        
        # --- DETECCIÓN DINÁMICA DE COLUMNAS ---
        col_codigo_idx = next((i for i, c in enumerate(df_temp.columns) if "cod" in c or "presupuesto" in c or "unnamed: 0" in c), 0)
        col_resumen_idx = next((i for i, c in enumerate(df_temp.columns) if "res" in c or "desc" in c or "unnamed: 3" in c), 2)
        
        # Busca exactamente tu columna "datos comerciales"
        col_comercial_idx = next((i for i, c in enumerate(df_temp.columns) if "datos comerciales" in c or "comercial" in c), None)

        col_ia_destino = ws.max_column + 1
        ws.cell(row=1, column=col_ia_destino, value="REVISIÓN IA").font = openpyxl.styles.Font(bold=True, color="0000FF")

        resultados_vista = []

        for row_idx in range(2, ws.max_row + 1):
            codigo = str(ws.cell(row=row_idx, column=col_codigo_idx + 1).value or '').strip()
            resumen = str(ws.cell(row=row_idx, column=col_resumen_idx + 1).value or '').strip()
            
            # Leer de la columna "datos comerciales"
            valor_comercial = ""
            if col_comercial_idx is not None:
                valor_comercial = str(ws.cell(row=row_idx, column=col_comercial_idx + 1).value or '').strip().lower()
            
            if codigo == "" or codigo.lower() == "none" or codigo.lower() == "código" or "capítulo" in codigo.lower() or "total" in resumen.lower():
                continue

            val_texto = ""
            codigo_upper = codigo.upper()

            # --- CASO 3: ELEMENTOS COMERCIALES ---
            if valor_comercial != "":
                if valor_comercial in cat_comercial:
                    info = cat_comercial[valor_comercial]
                    val_texto = f"🟣 EQUIPO COMERCIAL (Detectado en columna: {valor_comercial}). Marcas aconsejadas: {info['marca']} | Coste estimado: {info['precio']}."
                else:
                    val_texto = f"🟣 EQUIPO COMERCIAL (Rama nueva: {valor_comercial}). Revisar marcas autorizadas en el proyecto."

            # --- CASO 1: CÓDIGO NATIVO IVE ---
            elif any(ive_ref in codigo_upper for ive_ref in codigos_ive_referencia) or (len(codigo) >= 6 and codigo[0].isdigit() and codigo[1].isalpha()):
                val_texto = "Código IVE Para precios se deberían de usar de la base de precios del IVE actual."

            # --- CASOS 2 y 4: CYPE VS INVENTADO ---
            else:
                if any(c in codigo_upper for c in [".", "-"]) or len(codigo) >= 5:
                    val_texto = "Código CYPE Para precios se deberían de usar de la base de precios del IVE, son superiores y más acorde al mercado"
                else:
                    val_texto = "❌ CODIGO ERRONEO / INVENTADO | Cotejar con IVE"

            # Guardar celda en el Excel
            ws.cell(row=row_idx, column=col_ia_destino, value=val_texto)
            
            resultados_vista.append({
                "Código Original": codigo,
                "Descripción": resumen[:30] + "...",
                "Datos Comerciales": valor_comercial if valor_comercial != "" else "—",
                "Resultado Columna IA": val_texto
            })

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("✅ ¡Hecho! El motor está completamente sincronizado con tu columna 'datos comerciales'.")
        st.dataframe(pd.DataFrame(resultados_vista), use_container_width=True)
        
        st.download_button(
            label="📥 DESCARGAR EXCEL ACTUALIZADO (.XLSX)",
            data=output.getvalue(),
            file_name=f"{uploaded_file.name.split('.')[0]}_Revisado_IA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
            
    except Exception as e:
        st.error(f"Error en la ejecución del script: {e}")
