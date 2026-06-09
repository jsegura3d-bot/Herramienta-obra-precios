import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET

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
# CAPÍTULOS Y JERARQUÍA
# =========================

def detectar_capitulos(text):
    capitulos = {}
    for line in text.splitlines():
        if line.startswith("~C|"):
            parts = line.split("|")
            codigo = parts[1].strip()
            texto = parts[3].strip() if len(parts) > 3 else ""
            if codigo.endswith("#"):
                capitulos[codigo] = texto
    return capitulos


def asignar_capitulo_a_partidas(df, capitulos):
    for idx, row in df.iterrows():
        codigo = row["codigo"]
        padre = codigo[:3] + "#"
        df.at[idx, "capitulo"] = padre
        df.at[idx, "capitulo_nombre"] = capitulos.get(padre, "SIN CAPÍTULO")
    return df


def detectar_sistema(texto):
    t = texto.lower()
    reglas = {
        "BT": ["cuadro", "magnetotérmico", "diferencial", "eléctrica", "mecanismo"],
        "SANEAMIENTO": ["saneamiento", "pozo", "arqueta", "tubería pvc", "drenaje"],
        "FONTANERÍA": ["fontanería", "tubería", "grifo", "lavabo"],
        "PCI": ["bie", "extintor", "detector", "rociador"],
        "CLIMA": ["clima", "climatización", "split", "conducto"],
        "VENTILACIÓN": ["ventilación", "extractor"],
    }
    for sistema, palabras in reglas.items():
        if any(p in t for p in palabras):
            return sistema
    return "GENERAL"


def expandir_componentes(df):
    filas = []

    for _, row in df.iterrows():
        filas.append({**row, "tipo": "partida", "padre": None})

        for comp in row["descomp_materiales"]:
            filas.append({
                "codigo": comp["codigo"],
                "texto": f"Componente de {row['codigo']}",
                "unidad": "",
                "cantidad": comp["cantidad"],
                "precio": 0.0,
                "importe": 0.0,
                "descomp_materiales": [],
                "descomp_mano_obra": [],
                "descomp_maquinaria": [],
                "capitulo": row["capitulo"],
                "capitulo_nombre": row["capitulo_nombre"],
                "sistema": row.get("sistema", "GENERAL"),
                "tipo": "componente",
                "padre": row["codigo"],
            })

    return pd.DataFrame(filas)


def numeracion_tipo_B(df):
    numeraciones = {}

    for cap in df["capitulo"].unique():
        subdf = df[(df["capitulo"] == cap) & (df["tipo"] == "partida")]
        subdf = subdf.sort_values("codigo")
        for i, (idx, row) in enumerate(subdf.iterrows(), start=1):
            numeraciones[row["codigo"]] = f"{int(cap.replace('#',''))}.{i}"

    for padre in df["padre"].dropna().unique():
        hijos = df[df["padre"] == padre]
        base = numeraciones.get(padre, "0")
        for j, (idx, row) in enumerate(hijos.iterrows(), start=1):
            numeraciones[row["codigo"]] = f"{base}.{j}"

    df["numeracion"] = df["codigo"].apply(lambda c: numeraciones.get(c, ""))
    return df


# =========================
# MÓDULOS DE ANÁLISIS (13)
# =========================

# (TU BLOQUE COMPLETO DE ANÁLISIS SE MANTIENE IGUAL)
# NO LO REPITO AQUÍ PORQUE YA LO PEGASTE ENTERO
# Y NO HAY QUE TOCAR NADA


# =========================
# INTERFAZ STREAMLIT
# =========================

st.title("Auditor de presupuestos de obra (instalaciones)")

col1, col2 = st.columns(2)

with col1:
    bc3_file = st.file_uploader("Sube el archivo BC3", type=["bc3", "BC3", "txt"])
with col2:
    docx_file = st.file_uploader("Sube el DOCX de actuaciones (opcional)", type=["docx"])

doc_text = ""
faltantes_global = []
if docx_file is not None:
    try:
        import docx as docx_lib
        doc = docx_lib.Document(docx_file)
        doc_text = "\n".join(p.text for p in doc.paragraphs)
    except Exception:
        st.warning("No se ha podido leer el DOCX.")

if bc3_file is None:
    st.info("Sube un BC3 para empezar.")
    st.stop()

raw_text = bc3_file.read().decode("latin-1", errors="ignore")

df = parse_bc3_auto(raw_text.encode("latin-1"))

capitulos = detectar_capitulos(raw_text)
df = asignar_capitulo_a_partidas(df, capitulos)

df["sistema"] = df["texto"].apply(detectar_sistema)

df = expandir_componentes(df)
df = numeracion_tipo_B(df)

# (AQUÍ SIGUE TU BLOQUE DE ANÁLISIS, SIN CAMBIOS)

# =========================
# TABLA FINAL
# =========================

cols_mostrar = [
    "numeracion",
    "codigo",
    "texto",
    "capitulo_nombre",
    "unidad",
    "cantidad",
    "precio",
    "importe",
] + cols_alertas

st.dataframe(df_view[cols_mostrar].reset_index(drop=True), use_container_width=True)
