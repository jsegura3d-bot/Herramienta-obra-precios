import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import re
from io import BytesIO

st.set_page_config(page_title="Auditor de presupuestos de obra", layout="wide")

# Intentar importar PyPDF2 y docx opcionalmente
try:
    import PyPDF2
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False

try:
    import docx
    DOCX_SUPPORT = True
except ImportError:
    DOCX_SUPPORT = False


# =========================
# PARSER UNIVERSAL BC3
# =========================

def detectar_formato_bc3(text: str) -> str:
    lines = text.splitlines()
    sample = lines[:200]

    if any(l.startswith("~V|") for l in sample) and any(l.startswith("~C|") for l in sample):
        return "presto"
    if any(l.startswith("C;") or l.startswith("L;") for l in sample):
        return "arquimedes"
    if any("|P|" in l for l in sample):
        return "cype"
    if any("<ITEM>" in l or "<BC3" in l for l in sample):
        return "xml"
    return "desconocido"


def extraer_subpartida(codigo: str) -> tuple:
    """Extrae la subpartida (primeros niveles del código)"""
    if not codigo:
        return "", ""
    
    codigo_str = str(codigo)
    
    # Para códigos como 001, 001001, 001001001
    niveles = re.findall(r'\d{3}', codigo_str)
    
    if len(niveles) >= 2:
        subpartida = niveles[0]
        subpartida_detalle = '.'.join(niveles[:2])
    elif len(niveles) == 1:
        subpartida = niveles[0]
        subpartida_detalle = niveles[0]
    else:
        subpartida = codigo_str[:3] if len(codigo_str) >= 3 else codigo_str
        subpartida_detalle = codigo_str[:6] if len(codigo_str) >= 6 else codigo_str
    
    return subpartida, subpartida_detalle


def parse_presto(text: str) -> pd.DataFrame:
    partidas = {}
    lines = text.splitlines()

    for line in lines:
        if line.startswith("~C|"):
            parts = line.split("|")
            if len(parts) >= 5:
                codigo = parts[1].strip()
                unidad = parts[2].strip()
                texto = parts[3].strip()
                try:
                    precio = float(parts[4].replace(",", "."))
                except ValueError:
                    precio = 0.0

                cantidad = 1.0
                importe = precio
                subpartida, subpartida_detalle = extraer_subpartida(codigo)

                partidas[codigo] = {
                    "codigo": codigo,
                    "subpartida": subpartida,
                    "subpartida_detalle": subpartida_detalle,
                    "texto": texto,
                    "unidad": unidad,
                    "cantidad": cantidad,
                    "precio": precio,
                    "importe": importe,
                    "descomp_materiales": [],
                    "descomp_mano_obra": [],
                    "descomp_maquinaria": [],
                }

        elif line.startswith("~T|"):
            parts = line.split("|")
            if len(parts) >= 3:
                codigo = parts[1].strip()
                texto_largo = parts[2].strip()
                if codigo in partidas:
                    base = partidas[codigo]["texto"]
                    partidas[codigo]["texto"] = base + " " + texto_largo if base else texto_largo

    df = pd.DataFrame(list(partidas.values()))
    return df


def parse_arquimedes(text: str) -> pd.DataFrame:
    partidas = []
    for line in text.splitlines():
        parts = line.split(";")
        if len(parts) < 6:
            continue
        tipo = parts[0].strip().upper()
        if tipo == "L":
            try:
                cantidad = float(parts[4].replace(",", "."))
            except ValueError:
                cantidad = 0.0
            try:
                precio = float(parts[5].replace(",", "."))
            except ValueError:
                precio = 0.0
            
            codigo = parts[1].strip()
            subpartida, subpartida_detalle = extraer_subpartida(codigo)
            
            partidas.append({
                "codigo": codigo,
                "subpartida": subpartida,
                "subpartida_detalle": subpartida_detalle,
                "texto": parts[2].strip(),
                "unidad": parts[3].strip(),
                "cantidad": cantidad,
                "precio": precio,
                "importe": cantidad * precio,
                "descomp_materiales": [],
                "descomp_mano_obra": [],
                "descomp_maquinaria": [],
            })
    return pd.DataFrame(partidas)


def parse_cype(text: str) -> pd.DataFrame:
    partidas = []
    for line in text.splitlines():
        parts = line.split("|")
        if len(parts) < 8:
            continue
        tipo = parts[1].strip().upper()
        if tipo == "P":
            try:
                cantidad = float(parts[5].replace(",", "."))
            except ValueError:
                cantidad = 0.0
            try:
                precio = float(parts[6].replace(",", "."))
            except ValueError:
                precio = 0.0
            try:
                importe = float(parts[7].replace(",", "."))
            except ValueError:
                importe = cantidad * precio
            
            codigo = parts[2].strip()
            subpartida, subpartida_detalle = extraer_subpartida(codigo)
            
            partidas.append({
                "codigo": codigo,
                "subpartida": subpartida,
                "subpartida_detalle": subpartida_detalle,
                "texto": parts[3].strip(),
                "unidad": parts[4].strip(),
                "cantidad": cantidad,
                "precio": precio,
                "importe": importe,
                "descomp_materiales": [],
                "descomp_mano_obra": [],
                "descomp_maquinaria": [],
            })
    return pd.DataFrame(partidas)


def parse_xml(text: str) -> pd.DataFrame:
    partidas = []
    try:
        root = ET.fromstring(text)
    except Exception:
        return pd.DataFrame()

    for item in root.findall(".//ITEM"):
        codigo = item.findtext("CODE", "").strip()
        texto = item.findtext("DESC", "").strip()
        unidad = item.findtext("UNIT", "").strip()
        try:
            cantidad = float(item.findtext("QTY", "0").replace(",", "."))
        except ValueError:
            cantidad = 0.0
        try:
            precio = float(item.findtext("PRICE", "0").replace(",", "."))
        except ValueError:
            precio = 0.0
        importe = cantidad * precio
        
        subpartida, subpartida_detalle = extraer_subpartida(codigo)
        
        partidas.append({
            "codigo": codigo,
            "subpartida": subpartida,
            "subpartida_detalle": subpartida_detalle,
            "texto": texto,
            "unidad": unidad,
            "cantidad": cantidad,
            "precio": precio,
            "importe": importe,
            "descomp_materiales": [],
            "descomp_mano_obra": [],
            "descomp_maquinaria": [],
        })
    return pd.DataFrame(partidas)


def parse_bc3_auto(file_bytes: bytes) -> pd.DataFrame:
    text = file_bytes.decode("latin-1", errors="ignore")
    fmt = detectar_formato_bc3(text)

    if fmt == "presto":
        return parse_presto(text)
    if fmt == "arquimedes":
        return parse_arquimedes(text)
    if fmt == "cype":
        return parse_cype(text)
    if fmt == "xml":
        return parse_xml(text)

    raise ValueError("Formato BC3 no reconocido.")


# =========================
# LECTURA DE ACTUACIONES
# =========================

def leer_actuaciones(file) -> str:
    """Lee actuaciones desde PDF o DOCX"""
    texto = ""
    
    if file.type == "application/pdf":
        if not PDF_SUPPORT:
            return "⚠️ PyPDF2 no instalado. Ejecuta: pip install PyPDF2"
        try:
            pdf_reader = PyPDF2.PdfReader(file)
            for page in pdf_reader.pages:
                texto += page.extract_text()
        except Exception as e:
            return f"Error al leer PDF: {e}"
    
    elif file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        if not DOCX_SUPPORT:
            return "⚠️ python-docx no instalado. Ejecuta: pip install python-docx"
        try:
            doc = docx.Document(file)
            texto = "\n".join(p.text for p in doc.paragraphs)
        except Exception as e:
            return f"Error al leer DOCX: {e}"
    
    return texto


# =========================
# DETECTORES
# =========================

def es_texto_pobre(texto: str) -> bool:
    if not texto or not isinstance(texto, str):
        return True
    t = texto.strip()
    if len(t) < 10:
        return True
    genericas = ["varios", "completo", "instalación completa", "según proyecto", "trabajos varios"]
    return any(g.lower() in t.lower() for g in genericas)


def detectar_partidas_alzadas(row) -> bool:
    if row["cantidad"] == 0:
        return True
    if str(row["unidad"]).lower() in ["alzada", "global", "servicio", "trabajo"]:
        return True
    if "partida alzada" in str(row["texto"]).lower():
        return True
    return False


def detectar_partidas_incompletas(row) -> bool:
    texto = str(row["texto"]).lower()
    mo = row.get("descomp_mano_obra", [])
    if any(k in texto for k in ["instalación", "montaje", "suministro e instalación"]):
        if len(mo) == 0:
            return True
    return False


def detectar_elementos_sobrantes(row) -> bool:
    texto = str(row["texto"]).lower()
    comps = row.get("descomp_mano_obra", []) + row.get("descomp_materiales", [])
    if "tubería" in texto and any("roza" in c.get("descripcion", "").lower() for c in comps):
        return True
    if "cuadro" in texto and any("cable" in c.get("descripcion", "").lower() for c in comps):
        return True
    return False


def detectar_debe_dividirse(row) -> bool:
    texto = str(row["texto"]).lower()
    patrones = [
        ("cuadro", "cable"),
        ("cuadro", "mecanismo"),
        ("tubería", "aislamiento"),
        ("tubería", "soporte"),
        ("equipo", "instalación"),
        ("bie", "señalización"),
    ]
    for a, b in patrones:
        if a in texto and b in texto:
            return True
    return False


def detectar_sin_medicion(row) -> bool:
    return float(row["cantidad"]) == 0.0


def detectar_sin_descomposicion(row) -> bool:
    return (
        len(row.get("descomp_materiales", [])) == 0
        and len(row.get("descomp_mano_obra", [])) == 0
        and len(row.get("descomp_maquinaria", [])) == 0
    )


def detectar_sin_texto(row) -> bool:
    return es_texto_pobre(row["texto"])


def detectar_duplicadas(df: pd.DataFrame) -> pd.Series:
    clave = (
        df["texto"].fillna("").str.lower()
        + "|"
        + df["unidad"].fillna("").str.lower()
        + "|"
        + df["precio"].fillna(0).astype(str)
    )
    return clave.duplicated(keep=False)


def detectar_contradictorias(df: pd.DataFrame) -> pd.Series:
    contrad = pd.Series(False, index=df.index)
    textos = df["texto"].fillna("").str.lower()
    multicapa = textos.str.contains("multicapa")
    cobre = textos.str.contains("cobre")
    tuberia = textos.str.contains("tubería")
    mask_multi = tuberia & multicapa
    mask_cobre = tuberia & cobre
    if mask_multi.any() and mask_cobre.any():
        contrad = mask_multi | mask_cobre
    return contrad


def extraer_sistemas_desde_actuaciones(doc_text: str) -> list:
    sistemas = []
    claves = ["electricidad", "clima", "climatización", "pci", "fontanería", "saneamiento", "ventilación"]
    for c in claves:
        if c in doc_text.lower():
            sistemas.append(c)
    return list(set(sistemas))


def detectar_faltantes_segun_actuaciones(df: pd.DataFrame, doc_text: str):
    sistemas = extraer_sistemas_desde_actuaciones(doc_text)
    textos = df["texto"].fillna("").str.lower()
    faltantes = []
    for s in sistemas:
        if not textos.str.contains(s).any():
            faltantes.append(s)
    serie = pd.Series(False, index=df.index)
    return serie, faltantes


def detectar_precios_bajos(df: pd.DataFrame) -> pd.Series:
    precios = df["precio"].fillna(0)
    sospechoso = precios < 5

    familias = {
        "tubería": df["texto"].str.contains("tubería", case=False, na=False),
        "cable": df["texto"].str.contains("cable", case=False, na=False),
        "detector": df["texto"].str.contains("detector", case=False, na=False),
        "cuadro": df["texto"].str.contains("cuadro", case=False, na=False),
    }
    for fam, mask in familias.items():
        if mask.any():
            mediana = df.loc[mask, "precio"].median()
            sospechoso |= (mask & (df["precio"] < 0.5 * mediana))

    return sospechoso


def detectar_mano_obra_irreal(row) -> bool:
    texto = str(row["texto"]).lower()
    mo = row.get("descomp_mano_obra", [])

    if any(k in texto for k in ["instalación", "montaje", "suministro e instalación"]):
        if len(mo) == 0:
            return True

    for c in mo:
        if c.get("cantidad", 0) < 0.01:
            return True

    return False


def detectar_partidas_criticas_coste(df: pd.DataFrame) -> pd.Series:
    importe = df["importe"].fillna(df["cantidad"] * df["precio"])
    total = float(importe.sum()) if len(df) else 0.0
    crit = pd.Series(False, index=df.index)

    crit |= importe > 5000
    if total > 0:
        crit |= importe > 0.1 * total
    crit |= (df["precio"] < 5) & (df["cantidad"] > 100)
    return crit


# =========================
# VALORACIONES TEXTUALES
# =========================

def generar_valoracion(row) -> str:
    """Genera una valoración textual basada en las alertas de la partida"""
    alertas = []
    
    if row.get("alzada", False):
        alertas.append("⚠️ Partida alzada: Sin medición desglosada")
    
    if row.get("incompleta", False):
        alertas.append("🔧 Incompleta: Requiere descomposición de mano de obra")
    
    if row.get("sobran_elementos", False):
        alertas.append("🧩 Elementos sobrantes: Revisar duplicidades")
    
    if row.get("debe_dividirse", False):
        alertas.append("📋 Debe dividirse: Mezcla conceptos independientes")
    
    if row.get("sin_medicion", False):
        alertas.append("📏 Sin medición: Cantidad = 0")
    
    if row.get("sin_descomposicion", False):
        alertas.append("🔍 Sin descomposición: No se puede auditar coste real")
    
    if row.get("sin_texto", False):
        alertas.append("📝 Sin texto informativo: Descripción insuficiente")
    
    if row.get("duplicada", False):
        alertas.append("🔄 Duplicada: Posible error")
    
    if row.get("contradictoria", False):
        alertas.append("⚡ Contradictoria: Materiales incompatibles")
    
    if row.get("precio_bajo", False):
        alertas.append("💰 Precio bajo: Revisar costes")
    
    if row.get("mo_irreal", False):
        alertas.append("👷 Mano de obra irreal")
    
    if row.get("critica_coste", False):
        alertas.append("💸 Crítica por coste: Impacto >5000€ o >10% del total")
    
    if row.get("faltante_segun_actuaciones", False):
        alertas.append("📄 Faltante según actuaciones")
    
    if not alertas:
        return "✅ Correcta"
    
    return " | ".join(alertas)


# =========================
# INTERFAZ STREAMLIT
# =========================

st.title("🔍 Auditor de presupuestos de obra (instalaciones)")

col1, col2 = st.columns(2)

with col1:
    bc3_file = st.file_uploader("📁 Sube el archivo BC3", type=["bc3", "BC3", "txt"])

with col2:
    actuaciones_file = st.file_uploader(
        "📄 Sube actuaciones (PDF o DOCX) - opcional",
        type=["pdf", "docx"]
    )

doc_text = ""
faltantes_global = []

if actuaciones_file is not None:
    with st.spinner("Leyendo documento de actuaciones..."):
        doc_text = leer_actuaciones(actuaciones_file)
        if doc_text and not doc_text.startswith("⚠️") and not doc_text.startswith("Error"):
            st.success(f"✅ Documento cargado ({len(doc_text)} caracteres)")
        elif doc_text:
            st.warning(doc_text)

if bc3_file is None:
    st.info("📂 Sube un archivo BC3 para comenzar")
    st.stop()

try:
    with st.spinner("Procesando BC3..."):
        df = parse_bc3_auto(bc3_file.read())
except ValueError as e:
    st.error(str(e))
    st.stop()

if df.empty:
    st.warning("No se detectaron partidas en el BC3")
    st.stop()

# Asegurar columnas
for col in ["codigo", "subpartida", "subpartida_detalle", "texto", "unidad", "cantidad", "precio", "importe"]:
    if col not in df.columns:
        if col in ["cantidad", "precio", "importe"]:
            df[col] = 0.0
        else:
            df[col] = ""

# Ejecutar detecciones
df["alzada"] = df.apply(detectar_partidas_alzadas, axis=1)
df["incompleta"] = df.apply(detectar_partidas_incompletas, axis=1)
df["sobran_elementos"] = df.apply(detectar_elementos_sobrantes, axis=1)
df["debe_dividirse"] = df.apply(detectar_debe_dividirse, axis=1)
df["sin_medicion"] = df.apply(detectar_sin_medicion, axis=1)
df["sin_descomposicion"] = df.apply(detectar_sin_descomposicion, axis=1)
df["sin_texto"] = df.apply(detectar_sin_texto, axis=1)
df["precio_bajo"] = detectar_precios_bajos(df)
df["mo_irreal"] = df.apply(detectar_mano_obra_irreal, axis=1)
df["critica_coste"] = detectar_partidas_criticas_coste(df)
df["duplicada"] = detectar_duplicadas(df)
df["contradictoria"] = detectar_contradictorias(df)

if doc_text and not doc_text.startswith("⚠️") and not doc_text.startswith("Error"):
    serie_faltantes, faltantes_global = detectar_faltantes_segun_actuaciones(df, doc_text)
    df["faltante_segun_actuaciones"] = serie_faltantes
else:
    df["faltante_segun_actuaciones"] = False

# Generar valoraciones
df["valoracion"] = df.apply(generar_valoracion, axis=1)

# Sidebar
with st.sidebar:
    st.header("📊 Resumen")
    
    alertas_count = {
        "Alzadas": df["alzada"].sum(),
        "Incompletas": df["incompleta"].sum(),
        "Sobran elementos": df["sobran_elementos"].sum(),
        "Deben dividirse": df["debe_dividirse"].sum(),
        "Sin medición": df["sin_medicion"].sum(),
        "Sin descomposición": df["sin_descomposicion"].sum(),
        "Sin texto": df["sin_texto"].sum(),
        "Duplicadas": df["duplicada"].sum(),
        "Contradictorias": df["contradictoria"].sum(),
        "Precios bajos": df["precio_bajo"].sum(),
        "MO irreal": df["mo_irreal"].sum(),
        "Críticas coste": df["critica_coste"].sum(),
    }
    
    for alerta, count in alertas_count.items():
        if count > 0:
            st.write(f"🔴 {alerta}: **{int(count)}**")
    
    st.divider()
    st.metric("Total incidencias", int(sum(alertas_count.values())))
    st.metric("Partidas con problemas", int(df[list(alertas_count.keys())].any(axis=1).sum()))
    
    if faltantes_global:
        st.divider()
        st.markdown("### 🚨 Sistemas faltantes:")
        for s in faltantes_global:
            st.write(f"- {s}")

# Filtros
st.subheader("🎯 Filtros")

cols_alertas = ["alzada", "incompleta", "sobran_elementos", "debe_dividirse",
                "sin_medicion", "sin_descomposicion", "sin_texto", "duplicada",
                "contradictoria", "faltante_segun_actuaciones", "precio_bajo",
                "mo_irreal", "critica_coste"]

nombres_alertas = {
    "alzada": "Partidas alzadas", "incompleta": "Incompletas",
    "sobran_elementos": "Elementos que sobran", "debe_dividirse": "Deben dividirse",
    "sin_medicion": "Sin medición", "sin_descomposicion": "Sin descomposición",
    "sin_texto": "Sin texto", "duplicada": "Duplicadas",
    "contradictoria": "Contradictorias", "faltante_segun_actuaciones": "Faltantes",
    "precio_bajo": "Precios bajos", "mo_irreal": "Mano de obra irreal",
    "critica_coste": "Críticas por coste"
}

col_f1, col_f2 = st.columns(2)

with col_f1:
    filtro_tipo = st.multiselect("Alertas:", options=cols_alertas,
                                 format_func=lambda x: nombres_alertas.get(x, x), default=[])

with col_f2:
    subpartidas = sorted(df["subpartida"].dropna().unique())
    filtro_subpartida = st.multiselect("Capítulo:", options=subpartidas, default=[])

# Aplicar filtros
df_view = df.copy()
if filtro_tipo:
    mask = pd.Series(False, index=df_view.index)
    for c in filtro_tipo:
        mask |= df_view[c]
    df_view = df_view[mask]

if filtro_subpartida:
    df_view = df_view[df_view["subpartida"].isin(filtro_subpartida)]

# Mostrar resultados
st.subheader("📋 Partidas analizadas")

cols_mostrar = ["codigo", "subpartida", "subpartida_detalle", "texto", "unidad", "cantidad", "precio", "importe", "valoracion"]

st.dataframe(df_view[cols_mostrar].reset_index(drop=True), use_container_width=True,
             column_config={
                 "codigo": "Código", "subpartida": "Capítulo", "subpartida_detalle": "Subpartida",
                 "texto": "Descripción", "unidad": "Ud",
                 "cantidad": st.column_config.NumberColumn("Cantidad", format="%.2f"),
                 "precio": st.column_config.NumberColumn("Precio (€)", format="%.2f €"),
                 "importe": st.column_config.NumberColumn("Importe (€)", format="%.2f €"),
                 "valoracion": "Valoración"
             })

# Exportar
st.divider()
if st.button("📥 Exportar a CSV"):
    export_df = df_view[cols_mostrar].copy()
    csv = export_df.to_csv(index=False, encoding='utf-8-sig')
    st.download_button("Descargar CSV", data=csv, file_name="auditoria.csv", mime="text/csv")
