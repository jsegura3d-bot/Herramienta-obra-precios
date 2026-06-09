import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET

st.set_page_config(page_title="Auditor de presupuestos", layout="wide")

# ============================================================
# 1. DETECTOR DE FORMATO BC3
# ============================================================

def detectar_formato_bc3(text):
    if "~V|" in text and "~C|" in text:
        return "presto"
    return "desconocido"

# ============================================================
# 2. PARSER PRESTO ADAPTADO A TU BC3 REAL
# ============================================================

def parse_presto(text):
    partidas = {}
    capitulos = {}

    for line in text.splitlines():

        # -------------------------------
        # CAPÍTULOS Y PARTIDAS
        # -------------------------------
        if line.startswith("~C|"):
            parts = line.split("|")
            if len(parts) < 4:
                continue

            codigo = parts[1].strip()
            unidad = parts[2].strip()
            texto = parts[3].strip()

            # Capítulo
            if codigo.endswith("#"):
                capitulos[codigo] = texto
                continue

            # Partida
            try:
                precio = float(parts[4].replace(",", ".")) if len(parts) > 4 else 0.0
            except:
                precio = 0.0

            partidas[codigo] = {
                "codigo": codigo,
                "texto": texto,
                "unidad": unidad,
                "cantidad": 1.0,
                "precio": precio,
                "importe": precio,
                "descomp_materiales": [],
                "capitulo": None,
                "capitulo_nombre": None,
                "sistema": None,
            }

        # -------------------------------
        # DESCOMPOSICIÓN
        # -------------------------------
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
    return df, capitulos

# ============================================================
# 3. ASIGNAR CAPÍTULO A CADA PARTIDA
# ============================================================

def asignar_capitulos(df, capitulos):
    for idx, row in df.iterrows():
        codigo = row["codigo"]
        padre = codigo[:3] + "#"
        df.at[idx, "capitulo"] = padre
        df.at[idx, "capitulo_nombre"] = capitulos.get(padre, "SIN CAPÍTULO")
    return df

# ============================================================
# 4. DETECTAR SISTEMA (BT, SANEAMIENTO, PCI…)
# ============================================================

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

# ============================================================
# 5. EXPANDIR COMPONENTES COMO SUBPARTIDAS
# ============================================================

def expandir_componentes(df):
    filas = []

    for _, row in df.iterrows():
        # Partida original
        filas.append({**row, "tipo": "partida", "padre": None})

        # Componentes
        for comp in row["descomp_materiales"]:
            filas.append({
                "codigo": comp["codigo"],
                "texto": f"Componente de {row['codigo']}",
                "unidad": "",
                "cantidad": comp["cantidad"],
                "precio": 0.0,
                "importe": 0.0,
                "descomp_materiales": [],
                "capitulo": row["capitulo"],
                "capitulo_nombre": row["capitulo_nombre"],
                "sistema": row["sistema"],
                "tipo": "componente",
                "padre": row["codigo"],
            })

    return pd.DataFrame(filas)

# ============================================================
# 6. NUMERACIÓN TIPO B (ORDEN REAL)
# ============================================================

def numeracion_tipo_B(df):
    numeraciones = {}

    # Partidas
    for cap in df["capitulo"].unique():
        subdf = df[(df["capitulo"] == cap) & (df["tipo"] == "partida")]
        subdf = subdf.sort_values("codigo")

        for i, (idx, row) in enumerate(subdf.iterrows(), start=1):
            numeraciones[row["codigo"]] = f"{int(cap.replace('#',''))}.{i}"

    # Componentes
    for padre in df["padre"].dropna().unique():
        hijos = df[df["padre"] == padre]
        base = numeraciones.get(padre, "0")

        for j, (idx, row) in enumerate(hijos.iterrows(), start=1):
            numeraciones[row["codigo"]] = f"{base}.{j}"

    df["numeracion"] = df["codigo"].apply(lambda c: numeraciones.get(c, ""))
    return df

# ============================================================
# 7. INFORME FINAL
# ============================================================

def generar_linea_informe(row):
    return f"{row['numeracion']} / {row['codigo']} / {row['texto']} / {row['capitulo_nombre']} / {row['sistema']}"

# ============================================================
# 8. INTERFAZ STREAMLIT
# ============================================================

st.title("Auditor de presupuestos con jerarquía completa")

bc3_file = st.file_uploader("Sube tu BC3", type=["bc3", "txt"])

if bc3_file is None:
    st.stop()

raw = bc3_file.read().decode("latin-1", errors="ignore")

fmt = detectar_formato_bc3(raw)
if fmt != "presto":
    st.error("Formato no reconocido como Presto.")
    st.stop()

df, capitulos = parse_presto(raw)
df = asignar_capitulos(df, capitulos)

df["sistema"] = df["texto"].apply(detectar_sistema)

df = expandir_componentes(df)
df = numeracion_tipo_B(df)

st.subheader("Informe final")

df["informe"] = df.apply(generar_linea_informe, axis=1)

for linea in df["informe"]:
    st.write(linea)
