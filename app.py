import streamlit as st
import pandas as pd
import io
import openpyxl
import urllib.parse

st.set_page_config(page_title="Revisor IVE BDC25 - Valencia", layout="wide")
st.title("🛠️ Revisor de Precios Pro - Cruce Total (CYPE vs IVE Valencia)")
st.caption("Configuración activa: Caso 3 con Multibuscador Profesional (Matmax, Obramat y Google).")

# --- BANCO DE PRECIOS INTEGRADO IVE (VALENCIA - JULIO 2025) ---
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

# --- MULTIBUSCADOR COMERCIAL AVANZADO (Caso 3) ---
def generar_multibuscador_comercial(resumen, info_comercial):
    """
    Rastrea el catálogo interno o genera pasarelas de enlaces directos a las 
    principales webs de distribución y tarifas profesionales de España.
    """
    texto_busqueda = f"{info_comercial} {resumen}".lower()
    
    # 1. Catálogo interno directo por si es una marca prémium exacta
    catalogo_web = {
        "carandini": {"precio": "185.00 €", "url": "https://www.carandini.com/es/productos/"},
        "veka": {"precio": "185.00 €", "url": "https://www.veka.es/ventanas-veka/"},
        "schreder": {"precio": "320.00 €", "url": "https://www.schreder.com/es-es/productos"},
        "schréder": {"precio": "320.00 €", "url": "https://www.schreder.com/es-es/productos"},
        "socelec": {"precio": "320.00 €", "url": "https://www.schreder.com/es-es/productos"},
        "normalux": {"precio": "42.50 €", "url": "https://normagrup.com/es/marcas/normalux/"},
        "naos": {"precio": "42.50 €", "url": "https://normagrup.com/es/marcas/normalux/"},
        "luxomat": {"precio": "120.00 €", "url": "https://www.beg-luxomat.com/es/productos/"},
        "beg": {"precio": "120.00 €", "url": "https://www.beg-luxomat.com/es/productos/"},
        "schneider": {"precio": "16.80 €", "url": "https://www.se.com/es/es/product-subcategory/5100-mecanismos/"},
        "artec": {"precio": "16.80 €", "url": "https://www.se.com/es/es/product-subcategory/5100-mecanismos/"},
        "philips": {"precio": "65.00 €", "url": "https://www.lighting.philips.es/prof/luminarias-de-interior/"},
        "coreline": {"precio": "65.00 €", "url": "https://www.lighting.philips.es/prof/luminarias-de-interior/"},
        "ledvance": {"precio": "45.00 €", "url": "https://www.ledvance.es/profesional/productos/"},
        "osram": {"precio": "45.00 €", "url": "https://www.ledvance.es/profesional/productos/"},
        "jiso": {"precio": "38.00 €", "url": "https://jisoiluminacion.com/productos/"},
        "simon 100": {"precio": "32.00 €", "url": "https://www.simonelectric.com/simon100"},
        "simon 27": {"precio": "14.50 €", "url": "https://www.simonelectric.com/simon27"},
        "huawei": {"precio": "1850.00 €", "url": "https://solar.huawei.com/es/productos/"},
        "sun2000": {"precio": "1850.00 €", "url": "https://solar.huawei.com/es/productos/"},
        "daikin": {"precio": "4600.00 €", "url": "https://www.daikin.es/es_es/lineas-de-producto/aerotermia.html"}
    }
    
    for clave, datos in catalogo_web.items():
        if clave in texto_busqueda:
            return f"Precio catálogo: {datos['precio']} | Web oficial: {datos['url']}"
            
    # 2. Si es una referencia genérica o no indexada, preparamos las queries seguras
    query_comercial = f"{info_comercial} {resumen}"
    query_encoded = urllib.parse.quote(query_comercial)
    
    # Construcción de pasarelas de enlaces profesionales
    url_matmax = f"https://www.matmax.es/search?q={query_encoded}"
    url_obramat = f"https://www.google.com/search?q=site:obramat.es+{query_encoded}"
    url_shopping = f"https://www.google.com/search?q={query_encoded}&tbm=shop"
    
    # Texto limpio estructurado que se inyectará en la Columna IA del Excel
    return (
        f"Buscar precio profesional en: "
        f"[MATMAX]({url_matmax}) | "
        f"[OBRAMAT]({url_obramat}) | "
        f"[GOOGLE SHOPPING]({url_shopping})"
    )


uploaded_file = st.file_uploader("Sube tu Excel de Bugarra", type=["xlsx"])

if uploaded_file:
    try:
        file_bytes = uploaded_file.read()
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
        
        if 'Hoja5' in wb.sheetnames:
            ws = wb['Hoja5']
        else:
            ws = wb.active

        # --- ARQUITECTURA FIJA ABSOLUTA (Base 1 para openpyxl) ---
        col_codigo_idx = 1      # Columna A
        col_resumen_idx = 2     # Columna B
        col_comercial_idx = 3   # Columna C (DATOS COMERCIALES)
        col_pres_idx = 5        # Columna E (Precio Presupuesto)

        # Inyección de la columna de revisión al final de las celdas de la fila 1
        col_ia_destino = ws.max_column + 1
        ws.cell(row=1, column=col_ia_destino, value="COLUMNA IA (REVISIÓN DE MÁRGENES VALENCIA)").font = openpyxl.styles.Font(bold=True, color="0000FF")

        resultados_vista = []

        for row_idx in range(2, ws.max_row + 1):
            codigo = str(ws.cell(row=row_idx, column=col_codigo_idx).value or '').strip()
            resumen = str(ws.cell(row=row_idx, column=col_resumen_idx).value or '').strip()
            valor_comercial = str(ws.cell(row=row_idx, column=col_comercial_idx).value or '').strip()
            
            if codigo == "" or codigo.lower() == "none" or codigo.lower() == "código" or "capítulo" in codigo.lower() or "total" in resumen.lower():
                continue
                
            try:
                precio_presu = float(ws.cell(row=row_idx, column=col_pres_idx).value or 0.0)
            except:
                precio_presu = 0.0

            val_texto = ""
            codigo_upper = codigo.upper()

            tiene_info_comercial = valor_comercial != "" and valor_comercial.lower() != "none"

            # --- CASO 3: PRIORIDAD ABSOLUTA SI LA COLUMNA C TIENE INFO ---
            if tiene_info_comercial:
                val_texto = generar_multibuscador_comercial(resumen, valor_comercial)

            # --- SI LA COLUMNA C ESTÁ VACÍA, EJECUTA LOS CASOS RESTANTES ---
            else:
                # --- CASO 1: ANÁLISIS DE BASE DIRECTA IVE ---
                if codigo in precios_ive:
                    val_texto = "Código IVE revisar si el precio es actual."
                
                # --- CASO 2: SI DETECTA QUE ES UN CYPE ---
                elif any(c in codigo_upper for c in [".", "-", "_"]) or len(codigo_upper) > 6:
                    val_texto = "Código CYPE Para precios se deberían de usar de la base de precios del IVE, son superiores y más acorde al mercado."
                
                # --- CASO 4: CÓDIGO ERRÓNEO O NO IDENTIFICADO ---
                else:
                    val_texto = "🔍 REVISAR MANUALMENTE."

            # Guardamos el resultado en la celda del Excel original
            ws.cell(row=row_idx, column=col_ia_destino, value=val_texto)
            
            resultados_vista.append({
                "Partida": codigo,
                "Datos Comerciales (Col C)": valor_comercial if tiene_info_comercial else "Vacío",
                "Dictamen / Enlaces Generados": val_texto
            })

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("🚀 ¡Brutal! Despliegue completado. El Caso 3 ahora es un Multibuscador Profesional.")
        st.dataframe(pd.DataFrame(resultados_vista), use_container_width=True)
        
        st.download_button(
            label="📥 DESCARGAR TU EXCEL CON ACCESO A PORTALES PROFESIONALES (.XLSX)",
            data=output.getvalue(),
            file_name=f"{uploaded_file.name.split('.')[0]}_Revisado_IA_Buscador.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
            
    except Exception as e:
        st.error(f"Error técnico al procesar el formato del Excel: {e}")
