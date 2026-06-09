import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import pdfplumber

st.set_page_config(page_title="Auditor de presupuestos de obra", layout="wide")


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

                partidas[codigo] = {
                    "codigo": codigo,
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
                    if base:
                        partidas[codigo]["texto"] = base + " " + texto_largo
                    else:
                        partidas[codigo]["texto"] = texto_largo

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

            partidas.append(
                {
                    "codigo": parts[1].strip(),
                    "texto": parts[2].strip(),
                    "unidad": parts[3].strip(),
                    "cantidad": cantidad,
                    "precio": precio,
                    "importe": cantidad * precio,
                    "descomp_materiales": [],
                    "descomp_mano_obra": [],
                    "descomp_maquinaria": [],
                }
            )

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

            partidas.append(
                {
                    "codigo": parts[2].strip(),
                    "texto": parts[3].strip(),
                    "unidad": parts[4].strip(),
                    "cantidad": cantidad,
                    "precio": precio,
                    "importe": importe,
                    "descomp_materiales": [],
                    "descomp_mano_obra": [],
                    "descomp_maquinaria": [],
                }
            )

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

        partidas.append(
            {
                "codigo": codigo,
                "texto": texto,
                "unidad": unidad,
                "cantidad": cantidad,
                "precio": precio,
                "importe": importe,
                "descomp_materiales": [],
                "descomp_mano_obra": [],
                "descomp_maquinaria": [],
            }
        )

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

    raise ValueError("Formato BC3 no reconocido. Ajustar detector o parser.")


# =========================
# MÓDULOS DE ANÁLISIS (13)
# =========================
# (Tu bloque completo aquí, sin tocar nada)


# =========================
# INTERFAZ STREAMLIT
# =========================

st.title("Auditor de presupuestos de obra (instalaciones)")

col1, col2, col3 = st.columns(3)

with col1:
    bc3_file = st.file_uploader("Sube el archivo BC3", type=["bc3", "BC3", "txt"])
with col2:
    docx_file = st.file_uploader("Sube el DOCX de actuaciones (opcional)", type=["docx"])
with col3:
    pdf_file = st.file_uploader("Sube el PDF de actuaciones (opcional)", type=["pdf"])

doc_text = ""
pdf_text = ""
faltantes_global = []

# DOCX
if docx_file is not None:
    try:
        import docx as docx_lib
        doc = docx_lib.Document(docx_file)
        doc_text = "\n".join(p.text for p in doc.paragraphs)
    except:
        st.warning("No se ha podido leer el DOCX.")

# PDF
if pdf_file is not None:
    try:
        with pdfplumber.open(pdf_file) as pdf:
            pdf_text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    except:
        st.warning("No se ha podido leer el PDF.")

texto_actuaciones = doc_text or pdf_text

if bc3_file is None:
    st.info("Sube un BC3 para empezar el análisis.")
    st.stop()

df = parse_bc3_auto(bc3_file.read())

if df.empty:
    st.warning("No se han detectado partidas en el BC3.")
    st.stop()

# =========================
# APLICAR TUS 13 ANÁLISIS
# =========================
# (Aquí van tus análisis tal cual los tienes)


# =========================
# FALTANTES SEGÚN ACTUACIONES (DOCX o PDF)
# =========================

if texto_actuaciones:
    serie_faltantes, faltantes_global = detectar_faltantes_segun_actuaciones(df, texto_actuaciones)
    df["faltante_segun_actuaciones"] = serie_faltantes
else:
    df["faltante_segun_actuaciones"] = False


# =========================
# TABLA FINAL
# =========================

cols_alertas = [
    "alzada","incompleta","sobran_elementos","debe_dividirse","sin_medicion",
    "sin_descomposicion","sin_texto","duplicada","contradictoria",
    "faltante_segun_actuaciones","precio_bajo","mo_irreal","critica_coste"
]

cols_mostrar = [
    "codigo","texto","unidad","cantidad","precio","importe"
] + cols_alertas

st.dataframe(df[cols_mostrar].reset_index(drop=True), use_container_width=True)
