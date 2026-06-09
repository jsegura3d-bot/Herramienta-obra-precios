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

        elif line.startswith("~D|"):
            parts = line.split("|")
            if len(parts) < 3:
                continue

            padre = parts[1].strip()
            if padre not in partidas:
                continue

            trozos = parts[2].split("\\")
            comps = []

            for i in range(0, len(trozos) - 1, 2):
                cod = trozos[i].strip()
                try:
                    coef = float(trozos[i+1].replace(",", "."))
                except:
                    coef = 0.0
                comps.append({"codigo": cod, "cantidad": coef})

            partidas[padre]["descomp_materiales"] = comps

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

    raise ValueError("Formato BC3 no reconocido.")


# =========================
# LECTURA DOCX + PDF
# =========================

def leer_docx(docx_file):
    try:
        import docx as docx_lib
        doc = docx_lib.Document(docx_file)
        return "\n".join(p.text for p in doc.paragraphs)
    except:
        return ""


def leer_pdf(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            return "\n".join(page.extract_text() or "" for page in pdf.pages)
    except:
        return ""


# =========================
# INTERFAZ
# =========================

st.title("Auditor de presupuestos de obra (instalaciones)")

col1, col2, col3 = st.columns(3)

with col1:
    bc3_file = st.file_uploader("Sube el archivo BC3", type=["bc3", "txt"])
with col2:
    docx_file = st.file_uploader("Sube DOCX (opcional)", type=["docx"])
with col3:
    pdf_file = st.file_uploader("Sube PDF (opcional)", type=["pdf"])

texto_actuaciones = ""

if docx_file:
    texto_actuaciones = leer_docx(docx_file)

if not texto_actuaciones and pdf_file:
    texto_actuaciones = leer_pdf(pdf_file)

if bc3_file is None:
    st.stop()

df = parse_bc3_auto(bc3_file.read())

# =========================
# (AQUÍ VAN TUS 13 ANÁLISIS ORIGINALES)
# =========================

# … (NO LOS REPITO, TÚ YA LOS TIENES)

# =========================
# TABLA FINAL
# =========================

st.dataframe(df)
