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
# NUEVO → MÉTODO B: detectar sistemas según el contenido del BC3
# ============================================================

def detectar_sistemas_en_bc3(df):
    textos = df["texto"].fillna("").str.lower()
    sistemas = set()

    # Electricidad
    if textos.str.contains("cuadro|cable|mecanismo|enchufe|luminaria", regex=True).any():
        sistemas.add("electricidad")

    # Fontanería
    if textos.str.contains("tubería|válvula|bomba|sanitario|grifo", regex=True).any():
        sistemas.add("fontanería")

    # Climatización
    if textos.str.contains("conducto|rejilla|unidad interior|unidad exterior|clima", regex=True).any():
        sistemas.add("climatización")

    # PCI
    if textos.str.contains("bie|extintor|detector|rociador|pci", regex=True).any():
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
# NUEVO → MÉTODO A + MÉTODO B para faltantes según actuaciones
# ============================================================

# Método A: detectar sistemas desde DOCX
sistemas_doc = extraer_sistemas_desde_actuaciones(doc_text)

# Método B: si A no detecta nada, usar BC3
if not sistemas_doc:
    sistemas_doc = detectar_sistemas_en_bc3(df)

# Comparativa final
if sistemas_doc:
    texto_actuaciones = " ".join(sistemas_doc)
    serie_faltantes, faltantes_global = detectar_faltantes_segun_actuaciones(df, texto_actuaciones)
    df["faltante_segun_actuaciones"] = serie_faltantes
else:
    df["faltante_segun_actuaciones"] = False

# =========================
# RESTO DE TU APP (sin tocar)
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

    if faltantes_global:
        st.markdown("### Sistemas faltantes según actuaciones:")
        for s in faltantes_global:
            st.write(f"- {s}")

st.subheader("Filtros de visualización")

cols_alertas = [
    "alzada",
    "incompleta",
    "sobran_elementos",
    "debe_dividirse",
    "sin_medicion",
    "sin_descomposicion",
    "sin_texto",
    "duplicada",
    "contradictoria",
    "faltante_segun_actuaciones",
    "precio_bajo",
    "mo_irreal",
    "critica_coste",
]

alerta_sel = st.multiselect(
    "Mostrar solo partidas con estas alertas:",
    options=cols_alertas,
    default=[],
)

df_view = df.copy()
if alerta_sel:
    mask = False
    for c in alerta_sel:
        mask |= df_view[c]
    df_view = df_view[mask]

st.subheader("Partidas analizadas")

cols_mostrar = [
    "codigo",
    "texto",
    "unidad",
    "cantidad",
    "precio",
    "importe",
] + cols_alertas

st.dataframe(df_view[cols_mostrar].reset_index(drop=True), use_container_width=True)
