import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET

st.set_page_config(page_title="Auditor de presupuestos de obra", layout="wide")

# =========================
# PARSER UNIVERSAL BC3
# =========================
# (todo igual que tu código original)
# ...

# =========================
# MÓDULOS DE ANÁLISIS (13)
# =========================
# (todo igual que tu código original)
# ...

# ============================================================
# MÉTODOS B, C y D PARA DETECTAR SISTEMAS SI EL DOCX NO LOS DA
# ============================================================

def detectar_sistemas_en_bc3_palabras(df):
    textos = df["texto"].fillna("").str.lower()
    sistemas = set()

    if textos.str.contains("cuadro|cable|mecanismo|enchufe|luminaria", regex=True).any():
        sistemas.add("electricidad")

    if textos.str.contains("tubería|válvula|bomba|sanitario|grifo", regex=True).any():
        sistemas.add("fontanería")

    if textos.str.contains("conducto|rejilla|unidad interior|unidad exterior|clima", regex=True).any():
        sistemas.add("climatización")

    if textos.str.contains("bie|extintor|detector|rociador|pci", regex=True).any():
        sistemas.add("pci")

    return list(sistemas)


def detectar_sistemas_en_bc3_familias(df):
    textos = df["texto"].fillna("").str.lower()
    sistemas = set()

    familias = {
        "electricidad": ["cuadro", "cable", "mecanismo", "luminaria", "magnetotérmico"],
        "fontanería": ["tubería", "válvula", "bomba", "desagüe", "sanitario"],
        "climatización": ["conducto", "rejilla", "unidad interior", "unidad exterior", "split"],
        "pci": ["bie", "extintor", "detector", "rociador", "central pci"],
    }

    for sistema, palabras in familias.items():
        for p in palabras:
            if textos.str.contains(p, regex=False).any():
                sistemas.add(sistema)

    return list(sistemas)


def detectar_sistemas_por_capitulos(df):
    sistemas = set()

    for codigo in df["codigo"].astype(str):
        if codigo.startswith("2"):
            sistemas.add("electricidad")
        if codigo.startswith("3"):
            sistemas.add("fontanería")
        if codigo.startswith("4"):
            sistemas.add("climatización")
        if codigo.startswith("5"):
            sistemas.add("pci")

    return list(sistemas)

# =========================
# INTERFAZ STREAMLIT
# =========================

st.title("Auditor de presupuestos de obra (instalaciones)")

# (todo igual que tu código original)
# ...

# =========================
# APLICAR TUS 13 ANÁLISIS
# =========================
# (todo igual que tu código original)
# ...

# ============================================================
# MÉTODO A + B + C + D PARA FALTANTES SEGÚN ACTUACIONES
# ============================================================

# Método A → DOCX
sistemas = extraer_sistemas_desde_actuaciones(doc_text)

# Método B → palabras del BC3
if not sistemas:
    sistemas = detectar_sistemas_en_bc3_palabras(df)

# Método C → familias técnicas
if not sistemas:
    sistemas = detectar_sistemas_en_bc3_familias(df)

# Método D → capítulos del BC3
if not sistemas:
    sistemas = detectar_sistemas_por_capitulos(df)

# Comparativa final
if sistemas:
    texto_actuaciones = " ".join(sistemas)
    serie_faltantes, faltantes_global = detectar_faltantes_segun_actuaciones(df, texto_actuaciones)
    df["faltante_segun_actuaciones"] = serie_faltantes
else:
    df["faltante_segun_actuaciones"] = False

# =========================
# RESUMEN DE ALERTAS (AÑADIDO FALTANTES)
# =========================

with st.sidebar:
    st.header("Resumen de alertas")

    def count(col):
        return int(df[col].sum())

    st.write(f"🔴 Partidas alzadas: **{count('alzada')}**")
    st.write(f"🔴 Incompletas: **{count('incompleta')}**")
    st.write(f"🔴 Con elementos que sobran: **{count('sobran_elementos')}**")
    st.write(f"🔴 Deben dividirse: **{count('debe_dividirse')}**")
    st.write(f"🔴 Sin medición: **{count('sin_medicion')}**")
    st.write(f"🔴 Sin descomposición: **{count('sin_descomposicion')}**")
    st.write(f"🔴 Sin texto informativo: **{count('sin_texto')}**")
    st.write(f"🔴 Duplicadas: **{count('duplicada')}**")
    st.write(f"🔴 Contradictorias: **{count('contradictoria')}**")
    st.write(f"🔴 Precios bajos: **{count('precio_bajo')}**")
    st.write(f"🔴 Mano de obra inexistente/irreal: **{count('mo_irreal')}**")
    st.write(f"🔴 Críticas por coste: **{count('critica_coste')}**")

    # NUEVO → contador de faltantes
    st.write(f"🔴 Faltantes según actuaciones: **{len(faltantes_global)}**")

    if faltantes_global:
        st.markdown("### Sistemas faltantes según actuaciones:")
        for s in faltantes_global:
            st.write(f"- {s}")

# =========================
# TABLA FINAL (sin tocar)
# =========================

st.subheader("Filtros de visualización")

cols_alertas = [
    "alzada","incompleta","sobran_elementos","debe_dividirse","sin_medicion",
    "sin_descomposicion","sin_texto","duplicada","contradictoria",
    "faltante_segun_actuaciones","precio_bajo","mo_irreal","critica_coste"
]

alerta_sel = st.multiselect("Mostrar solo partidas con estas alertas:", options=cols_alertas, default=[])

df_view = df.copy()
if alerta_sel:
    mask = False
    for c in alerta_sel:
        mask |= df_view[c]
    df_view = df_view[mask]

st.subheader("Partidas analizadas")

cols_mostrar = ["codigo","texto","unidad","cantidad","precio","importe"] + cols_alertas

st.dataframe(df_view[cols_mostrar].reset_index(drop=True), use_container_width=True)
