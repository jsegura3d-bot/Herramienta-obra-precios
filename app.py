import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Revisor IVE BDC25", layout="wide")
st.title("🛠️ Revisor de Precios - Sincronizado con IVE BDC25")
st.write("Cruce automático de presupuestos (CYPE / Genéricos) con la base oficial del IVE.")

# --- BANCO DE PRECIOS OFICIAL IVE BDC25 ---
# Precios y códigos reales extraídos directamente de la plataforma bdc.f-ive.es
precios_ive = {
    "0AF010": {"precio": 73.12, "codigo_oficial": "0AF010", "keywords": ["acometida", "agua", "potable"]},
    "PIBB.2e": {"precio": 2616.91, "codigo_oficial": "PIBB.2e", "keywords": ["aerotermia", "bomba calor", "acumulador"]},
    "EIEM.1ac": {"precio": 35.50, "codigo_oficial": "EIEM.1ac", "keywords": ["interruptor", "estanco", "mecanismo"]},
    "EIEM.1ab": {"precio": 37.55, "codigo_oficial": "EIEM.1ab", "keywords": ["tecla", "unipolar"]},
    "EIEM.5eg2": {"precio": 210.32, "codigo_oficial": "EIEM.5eg2", "keywords": ["detector", "presencia", "movimiento"]},
    "EIEM.2db": {"precio": 44.19, "codigo_oficial": "EIEM.2db", "keywords": ["toma", "corriente", "enchufe"]},
    "EIEC.3dd": {"precio": 8.91, "codigo_oficial": "EIEC.3dd", "keywords": ["tubo", "canalización", "rígido"]}
}

# --- BANCO DE DATOS CYPE GENERADOR DE PRECIOS ---
precios_cype_fijos = {
    "0AE010": {"precio": 292.54, "codigo_oficial": "0AE010", "desc": "Desconexión de acometida eléctrica"},
    "0AS010": {"precio": 203.04, "codigo_oficial": "0AS010", "desc": "Desconexión de acometida de saneamiento"},
    "DPT020": {"precio": 5.84, "codigo_oficial": "DPT020", "desc": "Demolición de partición interior"},
    "IFA005": {"precio": 36.94, "codigo_oficial": "IFA005", "desc": "Acometida de abastecimiento de agua potable"}
}

uploaded_file = st.file_uploader("Sube tu Excel de Bugarra", type=["xlsx"])

if uploaded_file:
    try:
        # Intento de lectura flexible de pestañas
        try:
            df = pd.read_excel(uploaded_file, sheet_name='Hoja5')
        except:
            df = pd.read_excel(uploaded_file)
        
        # Limpieza de espacios en los nombres de las columnas
        df.columns = [str(c).strip() for c in df.columns]
        
        # Localización dinámica de las columnas del presupuesto
        col_codigo = next((c for c in df.columns if "cod" in c.lower() or c == "Presupuesto" or "unnamed" in c.lower()), df.columns[0])
        col_ud = next((c for c in df.columns if "ud" in c.lower() or "nat" in c.lower() or "unnamed: 2" in c.lower()), None)
        col_resumen = next((c for c in df.columns if "res" in c.lower() or "desc" in c.lower() or "unnamed: 3" in c.lower()), None)
        col_pres = next((c for c in df.columns if ("pres" in c.lower() or "imp" in c.lower() or "prec" in c.lower()) and "can" not in c.lower() or "unnamed: 5" in c.lower()), None)

        resultados = []

        for index, fila in df.iterrows():
            codigo = str(fila.get(col_codigo, '')).strip()
            resumen = str(fila.get(col_resumen, '')).strip() if col_resumen else ""
            
            # Saltarse filas de títulos, celdas repetidas de cabecera o vacías
            if pd.isna(fila.get(col_codigo)) or codigo == "" or codigo.lower() == "none" or codigo.lower() == "código" or "capítulo" in codigo.lower() or "total" in resumen.lower():
                continue
                
            ud_valor = str(fila.get(col_ud, '')).strip() if col_ud else "ud"
            
            try:
                precio_presu = float(fila.get(col_pres, 0))
            except:
                precio_presu = 0.0

            descripcion_corta = resumen.split('\n')[0][:60]
            if len(resumen) > 60: descripcion_corta += "..."

            nuevo_codigo_ive = "—"
            p_ive_col = "—"
            nuevo_codigo_cype = "—"
            p_cype_col = "—"
            val_texto = ""

            texto_analisis = (codigo + " " + resumen).lower()

            # 1. COMPROBACIÓN EN BANCO IVE
            if codigo in precios_ive:
                p_ive_col = f"{precios_ive[codigo]['precio']} €"
                nuevo_codigo_ive = precios_ive[codigo]['codigo_oficial']
                if precio_presu <= precios_ive[codigo]['precio']:
                    val_texto = "🟢 IVE OK (Precio cubierto)"
                else:
                    val_texto = f"🔴 ALERTA: PRECIO PRESUPUESTO SUPERA AL IVE ({precios_ive[codigo]['precio']} €)"
            
            # 2. COMPROBACIÓN EN BANCO CYPE
            elif codigo in precios_cype_fijos:
                p_cype_col = f"{precios_cype_fijos[codigo]['precio']} €"
                nuevo_codigo_cype = precios_cype_fijos[codigo]['codigo_oficial']
                if precio_presu <= precios_cype_fijos[codigo]['precio']:
                    val_texto = "🟢 CYPE OK (Precio cubierto)"
                else:
                    val_texto = f"🔴 ALERTA: PRECIO PRESUPUESTO SUPERA A CYPE ({precios_cype_fijos[codigo]['precio']} €)"
            
            # 3. IDENTIFICACIÓN EN BASE A DESCRIPCIONES (DESMONTAJES / DEMOLICIONES)
            else:
                if "desconexión" in texto_analisis and "saneamiento" in texto_analisis:
                    nuevo_codigo_cype = "0AS010"
                    p_cype_col = f"{precios_cype_fijos['0AS010']['precio']} €"
                    val_texto = "🟡 SUGERENCIA CYPE: ASIGNAR 0AS010"
                elif "demolición" in texto_analisis and "partición" in texto_analisis:
                    nuevo_codigo_cype = "DPT020"
                    p_cype_col = f"{precios_cype_fijos['DPT020']['precio']} €"
                    val_texto = "🟡 SUGERENCIA CYPE: ASIGNAR DPT020"
                else:
                    val_texto = "🔍 REVISAR MANUALMENTE"

            # --- MOTOR DE MEJORA CRUZADA (IVE ≥ PRESU) ---
            # Si la partida no viene con código IVE de origen, buscamos si hay equivalente que mejore el precio
            if p_ive_col == "—":
                for cod_ive_ref, info_ive in precios_ive.items():
                    if any(kw in texto_analisis for kw in info_ive["keywords"]):
                        if info_ive["precio"] >= precio_presu:
                            nuevo_codigo_ive = info_ive['codigo_oficial']
                            p_ive_col = f"{info_ive['precio']} €"
                            val_texto = f"🔵 OPTIMIZACIÓN IVE: CAMBIAR A {nuevo_codigo_ive} (PRECIO MÁS SEGURO)"
                            break

            resultados.append({
                "Código Original": codigo,
                "Descripción Corta": descripcion_corta,
                "Unidad": ud_valor,
                "Precio Presu": f"{precio_presu} €",
                "Nuevo Código IVE (Real)": nuevo_codigo_ive,
                "Precio IVE": p_ive_col,
                "Nuevo Código CYPE": nuevo_codigo_cype,
                "Precio CYPE": p_cype_col,
                "VALORACIÓN EN OBRA": val_texto
            })

        if resultados:
            df_final = pd.DataFrame(resultados)
            st.success("✅ Estructura corregida y sincronizada correctamente.")
            st.dataframe(df_final, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Validacion_IVE')
            
            st.download_button(
                label="📥 DESCARGAR INFORME CORREGIDO IVE (.XLSX)",
                data=output.getvalue(),
                file_name="Informe_Precios_Validado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No se encontraron registros válidos en las filas procesadas.")
            
    except Exception as e:
        st.error(f"Error en la ejecución del script: {e}")
