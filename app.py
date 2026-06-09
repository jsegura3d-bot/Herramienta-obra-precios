import io
import re
from collections import defaultdict

import streamlit as st
import pandas as pd

st.set_page_config(page_title="Auditor de presupuestos de obra", layout="wide")


# =========================
# PARSER BC3 (MVP COMPLETO)
# =========================

def parse_bc3(file_bytes: bytes) -> pd.DataFrame:
    """
    Parser BC3 muy simplificado orientado a MVP.
    Asume:
      - Líneas de partidas con tipo 'C' (capítulos) y 'L' (líneas/partidas).
      - Formato típico: TIPO;CODIGO;TEXTO;UNIDAD;CANTIDAD;PRECIO;IMPORTE;...
      - Líneas de descomposición con referencia a código padre.
    ADÁPTALO a tu formato real de BC3 si difiere.
    """
    text = file_bytes.decode("latin-1", errors="ignore")
    lines = text.splitlines()

    partidas = []
    descomp_por_codigo = defaultdict(list)

    # Ejemplo de patrones muy genéricos (ajustar a tu BC3 real)
    # Supongamos:
    # P;CODIGO;TEXTO;UNIDAD;CANTIDAD;PRECIO;IMPORTE
    # D;COD_PADRE;TIPO(M,MdO,Ma);DESCRIPCION;CANTIDAD;PRECIO
    for raw in lines:
        parts = raw.split(";")
        if len(parts) < 2:
            continue

        tipo = parts[0].strip().upper()

        # Partida principal
        if tipo == "P" and len(parts) >= 7:
            codigo = parts[1].strip()
            texto = parts[2].strip()
            unidad = parts[3].strip()
            try:
                cantidad = float(parts[4].replace(",", ".") or 0)
            except ValueError:
                cantidad = 0.0
            try:
                precio = float(parts[5].replace(",", ".") or 0)
            except ValueError:
                precio = 0.0
            try:
                importe = float(parts[6].replace(",", ".") or 0)
            except ValueError:
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

        # Descomposición
        elif tipo == "D" and len(parts) >= 6:
            cod_padre = parts[1].strip()
            tipo_comp = parts[2].strip().upper()  # M, MO, MAQ...
            desc = parts[3].strip()
            try:
                cant = float(parts[4].replace(",", ".") or 0)
            except ValueError:
                cant = 0.0
            try:
                prec = float(parts[5].replace(",", ".") or 0)
            except ValueError:
                prec = 0.0

            descomp_por_codigo[cod_padre].append(
                {
                    "tipo": tipo_comp,
                    "descripcion": desc,
                    "cantidad": cant,
                    "precio": prec,
                }
            )

    # Vincular descomposición a partidas
    partidas_por_codigo = {p["codigo"]: p for p in partidas}
    for cod, comps in descomp_por_codigo.items():
        if cod not in partidas_por_codigo:
            continue
        mat = []
        mo = []
        maq = []
        for c in comps:
            if c["tipo"] in ("M", "MAT", "MATERIAL"):
                mat.append(c)
            elif c["tipo"] in ("MO", "MANO", "MANO_OBRA"):
                mo.append(c)
            elif c["tipo"] in ("MAQ", "MAQUINARIA"):
                maq.append(c)
        partidas_por_codigo[cod]["descomp_materiales"] = mat
        partidas_por_codigo[cod]["descomp_mano_obra"] = mo
        partidas_por_codigo[cod]["descomp_maquinaria"] = maq

    df = pd.DataFrame(partidas)
    if df.empty:
        st.warning("No se han detectado partidas en el BC3. Revisa el formato o ajusta el parser.")
    return df


# =========================
# MÓDULOS DE ANÁLISIS (13)
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
    if "completa" in str(row["texto"]).lower():
        return True
    return False


def detectar_partidas_incompletas(row) -> bool:
    # MVP: si no hay mano de obra ni maquinaria y el tipo de partida lo requiere
    texto = str(row["texto"]).lower()
    mo = row["descomp_mano_obra"]
    maq = row["descomp_maquinaria"]
    if any(k in texto for k in ["instalación", "montaje", "suministro e instalación"]):
        if len(mo) == 0:
            return True
    # Podrías añadir reglas por tipo (tubería, cuadro, split, etc.)
    return False


def detectar_elementos_sobrantes(row) -> bool:
    texto = str(row["texto"]).lower()
    # Ejemplos típicos: rozas dentro de tubería, cableado dentro de cuadro, etc.
    if "tubería" in texto and any("roza" in c["descripcion"].lower() for c in row["descomp_mano_obra"] + row["descomp_materiales"]):
        return True
    if "cuadro" in texto and any("cable" in c["descripcion"].lower() for c in row["descomp_materiales"]):
        return True
    return False


def detectar_debe_dividirse(row) -> bool:
    texto = str(row["texto"]).lower()
    # Mezcla de conceptos típicos
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
    return row["cantidad"] == 0


def detectar_sin_descomposicion(row) -> bool:
    return (
        len(row["descomp_materiales"]) == 0
        and len(row["descomp_mano_obra"]) == 0
        and len(row["descomp_maquinaria"]) == 0
    )


def detectar_sin_texto(row) -> bool:
    return es_texto_pobre(row["texto"])


def detectar_duplicadas(df: pd.DataFrame) -> pd.Series:
    # Duplicadas por texto + unidad + precio
    clave = df["texto"].str.lower().fillna("") + "|" + df["unidad"].str.lower().fillna("") + "|" + df["precio"].astype(str)
    return clave.duplicated(keep=False)


def detectar_contradictorias(df: pd.DataFrame) -> pd.Series:
    # MVP: detecta materiales contradictorios por texto (multicapa vs cobre, etc.)
    contrad = pd.Series(False, index=df.index)
    textos = df["texto"].str.lower().fillna("")
    multicapa = textos.str.contains("multicapa")
    cobre = textos.str.contains("cobre")
    # Si en el mismo presupuesto hay ambas para la misma familia (ej. "tubería")
    tuberia = textos.str.contains("tubería")
    mask_multi = tuberia & multicapa
    mask_cobre = tuberia & cobre
    if mask_multi.any() and mask_cobre.any():
        contrad = mask_multi | mask_cobre
    return contrad


def extraer_sistemas_desde_actuaciones(doc_text: str) -> list:
    # MVP: busca palabras clave de sistemas
    sistemas = []
    claves = ["electricidad", "clima", "climatización", "pci", "fontanería", "saneamiento", "ventilación"]
    for c in claves:
        if c in doc_text.lower():
            sistemas.append(c)
    return list(set(sistemas))


def detectar_faltantes_segun_actuaciones(df: pd.DataFrame, doc_text: str) -> pd.Series:
    """
    MVP: si en actuaciones aparece un sistema y no hay ninguna partida que lo mencione, marcamos faltante.
    Devuelve una serie booleana por partida, pero el concepto de "faltante" es global;
    aquí marcamos todas las filas como False y dejamos la info en un texto aparte.
    """
    sistemas = extraer_sistemas_desde_actuaciones(doc_text)
    textos = df["texto"].str.lower().fillna("")
    faltantes = []
    for s in sistemas:
        if not textos.str.contains(s).any():
            faltantes.append(s)
    # Para el MVP, devolvemos False en todas las filas y mostramos los sistemas faltantes en el panel lateral.
    serie = pd.Series(False, index=df.index)
    return serie, faltantes


def detectar_precios_bajos(df: pd.DataFrame) -> pd.Series:
    """
    MVP sin conexión real a IVE/CYPE/comerciales:
    - Marca como sospechosamente bajo si precio < 5 €
    - O si precio < 50% de la mediana de partidas similares por texto.
    """
    precios = df["precio"].fillna(0)
    sospechoso = precios < 5

    # Comparación simple por familia de texto (ej. "tubería", "cable", "detector")
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
    """
    MVP: si hay mano de obra pero con cantidades ridículas (<0.01 h por unidad),
    o si el texto implica instalación y no hay mano de obra.
    """
    texto = str(row["texto"]).lower()
    mo = row["descomp_mano_obra"]

    if any(k in texto for k in ["instalación", "montaje", "suministro e instalación"]):
        if len(mo) == 0:
            return True

    for c in mo:
        if c["cantidad"] < 0.01:
            return True

    return False


def detectar_partidas_criticas_coste(df: pd.DataFrame) -> pd.Series:
    """
    Marca como críticas:
      - importe > 5000 €
      - o importe > 10% del capítulo (MVP: del total)
      - o combinación precio bajo + cantidad alta
    """
    importe = df["importe"].fillna(df["cantidad"] * df["precio"])
    total = importe.sum() if len(df) else 0
    crit = pd.Series(False, index=df.index)

    crit |= importe > 5000
    if total > 0:
        crit |= importe > 0.1 * total

    # combinación peligrosa: precio bajo + cantidad alta
    crit |= (df["precio"] < 5) & (df["cantidad"] > 100)

    return crit


# =========================
# INTERFAZ STREAMLIT
# =========================

st.title("Auditor de presupuestos de obra (instalaciones)")

st.markdown(
    """
Herramienta para analizar un **BC3** y detectar:

- Partidas alzadas  
- Partidas incompletas  
- Partidas con elementos que sobran  
- Partidas que deberían dividirse  
- Partidas sin medición  
- Partidas sin descomposición  
- Partidas sin texto informativo  
- Partidas duplicadas  
- Partidas contradictorias  
- Partidas faltantes según actuaciones  
- Precios bajos  
- Mano de obra inexistente o irreal  
- Partidas críticas por coste  
"""
)

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
        st.warning("No se ha podido leer el DOCX. Se omite el análisis de actuaciones.")

if bc3_file is None:
    st.info("Sube un BC3 para empezar el análisis.")
    st.stop()

df = parse_bc3(bc3_file.read())
if df.empty:
    st.stop()

# Aplicar análisis fila a fila
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

# Duplicadas y contradictorias
df["duplicada"] = detectar_duplicadas(df)
df["contradictoria"] = detectar_contradictorias(df)

# Faltantes según actuaciones
if doc_text:
    serie_faltantes, faltantes_global = detectar_faltantes_segun_actuaciones(df, doc_text)
    df["faltante_segun_actuaciones"] = serie_faltantes
else:
    df["faltante_segun_actuaciones"] = False

# Panel lateral de resumen
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

# Filtros
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

# Mostrar tabla
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
