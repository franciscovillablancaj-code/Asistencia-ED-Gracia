import streamlit as st
import pandas as pd
import gspread
import io
import base64
import altair as alt
from pathlib import Path
from PIL import Image
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ---------------------------------------------------------------------------
# CONFIGURACIÓN — ajusta estos valores a tu planilla real
# ---------------------------------------------------------------------------
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

# ID de la hoja "índice": una Google Sheet con dos columnas: Persona | ID_Planilla
INDICE_SHEET_ID = "13TRbRXJayNXzjYiWuMYrUo3dDSIzIGlovKrkct7Gk6w"

# Nombre de la pestaña dentro de cada planilla individual
NOMBRE_PESTANA = "ASISTENCIA VOL"

# Fila donde están los encabezados de columna (antes hay metadata de la encargada)
FILA_ENCABEZADOS = 5

# Columnas de datos de cada voluntario/a (deben calzar EXACTO con el texto de la fila 5)
COL_REGION = "REGIÓN"
COL_HNA_ENCARGADA = "HNA ENCARGADA"
COL_PROGRAMA = "PROGRAMA"
COL_NOMBRE = "NOMBRE Y APELLIDO"
COL_IGLESIA = "IGLESIA"
COL_RANGO_EDAD = "RANGO EDAD"
COL_ROL = "ROL"

# Columnas de meses, en el mismo orden en que aparecen en la planilla
MESES = ["MR", "AB", "MY", "JN", "JL", "AG", "SEP", "OCT", "NOV", "DIC"]

# Letra que marca asistencia en la celda del mes
MARCA_PRESENTE = "P"

# Marca de agua: archivo de logo (debe estar junto a app.py) y qué tan visible es
LOGO_MARCA_AGUA = "logo.png"
OPACIDAD_MARCA_AGUA = 0.05

COLUMNAS_ESPERADAS = [
    COL_REGION, COL_HNA_ENCARGADA, COL_PROGRAMA, COL_NOMBRE,
    COL_IGLESIA, COL_RANGO_EDAD, COL_ROL,
] + MESES

# Orden esperado de las categorías del desplegable "RANGO EDAD"
ORDEN_RANGO_EDAD = ["menor", "10 a 13 años", "14 a 17 años", "18 a 29 años", "más de 30"]

# Texto de la opción "ver todo el año" en el selector de mes
OPCION_ANIO_COMPLETO = "Año completo (hasta la fecha)"

# Nombre completo de cada mes, solo para mostrar en pantalla (la planilla sigue
# usando las abreviaturas MR, AB, etc. como nombre de columna)
MES_NOMBRE_COMPLETO = {
    "MR": "Marzo", "AB": "Abril", "MY": "Mayo", "JN": "Junio", "JL": "Julio",
    "AG": "Agosto", "SEP": "Septiembre", "OCT": "Octubre", "NOV": "Noviembre", "DIC": "Diciembre",
}


def etiqueta_mes(m: str) -> str:
    """Texto a mostrar para un mes: su nombre completo, o tal cual si es la
    opción de año completo."""
    return m if m == OPCION_ANIO_COMPLETO else MES_NOMBRE_COMPLETO.get(m, m)

# ---------------------------------------------------------------------------
# CONEXIÓN A GOOGLE
# ---------------------------------------------------------------------------
@st.cache_resource
def get_client():
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)


@st.cache_resource
def get_drive_service():
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return build("drive", "v3", credentials=creds)


@st.cache_data(ttl=300)
def cargar_indice() -> pd.DataFrame:
    """Lee la hoja índice por POSICIÓN (columna A = Persona, columna B = ID_Planilla),
    sin depender de que el texto del encabezado calce exacto."""
    client = get_client()
    sh = client.open_by_key(INDICE_SHEET_ID)
    valores = sh.sheet1.get_all_values()

    filas = valores[1:]  # salta la fila de encabezado, sea cual sea su texto
    filas = [f for f in filas if len(f) >= 2 and f[0].strip() and f[1].strip()]

    df = pd.DataFrame(
        {"Persona": [f[0].strip() for f in filas], "ID_Planilla": [f[1].strip() for f in filas]}
    )
    return df


def _normalizar_encabezados(encabezados_crudos: list) -> list:
    encabezados, vistos = [], {}
    for i, h in enumerate(encabezados_crudos):
        h = str(h).strip() or f"_col{i}"
        if h in vistos:
            vistos[h] += 1
            h = f"{h}_{vistos[h]}"
        else:
            vistos[h] = 0
        encabezados.append(h)
    return encabezados


def _leer_desde_sheets(spreadsheet_id: str) -> pd.DataFrame:
    """Lee una Google Sheet nativa vía la API de Sheets (gspread)."""
    client = get_client()
    sh = client.open_by_key(spreadsheet_id)
    ws = sh.worksheet(NOMBRE_PESTANA)
    valores = ws.get_all_values()

    encabezados = _normalizar_encabezados(valores[FILA_ENCABEZADOS - 1])
    filas_datos = valores[FILA_ENCABEZADOS:]
    filas_datos = [f + [""] * (len(encabezados) - len(f)) for f in filas_datos]
    return pd.DataFrame(filas_datos, columns=encabezados)


def _leer_desde_office(spreadsheet_id: str) -> pd.DataFrame:
    """Descarga un .xlsx/.xls subido a Drive (no convertido a Sheets) y lo lee con pandas."""
    drive = get_drive_service()
    request = drive.files().get_media(fileId=spreadsheet_id)
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buffer.seek(0)

    df = pd.read_excel(
        buffer, sheet_name=NOMBRE_PESTANA, header=FILA_ENCABEZADOS - 1, dtype=str
    )
    df.columns = _normalizar_encabezados(list(df.columns))
    return df.fillna("")


def _normalizar_columnas_conocidas(df: pd.DataFrame) -> pd.DataFrame:
    """Renombra las columnas que calzan con las esperadas en CONFIGURACIÓN.
    Primero intenta calce exacto (ignorando mayúsculas/espacios); si no lo
    encuentra, busca la frase esperada como parte del encabezado real
    (para tolerar variantes como 'NOMBRE Y APELLIDO VOLUNTARIA/O')."""
    mapa = {}
    usados = set()

    disponibles = {str(c).strip().upper(): c for c in df.columns}
    for esperada in COLUMNAS_ESPERADAS:
        clave = esperada.strip().upper()
        if clave in disponibles and disponibles[clave] not in usados:
            mapa[disponibles[clave]] = esperada
            usados.add(disponibles[clave])

    pendientes = [e for e in COLUMNAS_ESPERADAS if e not in mapa.values()]
    for esperada in pendientes:
        clave = esperada.strip().upper()
        for col in df.columns:
            if col in usados:
                continue
            if clave in str(col).strip().upper():
                mapa[col] = esperada
                usados.add(col)
                break

    return df.rename(columns=mapa)


@st.cache_data(ttl=300)
def cargar_planilla(spreadsheet_id: str) -> pd.DataFrame:
    """Lee la planilla sin importar si es una Google Sheet nativa o un Excel
    subido a Drive sin convertir."""
    drive = get_drive_service()
    metadata = drive.files().get(fileId=spreadsheet_id, fields="mimeType").execute()

    if metadata["mimeType"] == "application/vnd.google-apps.spreadsheet":
        df = _leer_desde_sheets(spreadsheet_id)
    else:
        df = _leer_desde_office(spreadsheet_id)

    df = _normalizar_columnas_conocidas(df)

    requeridas = [COL_NOMBRE, COL_IGLESIA, COL_ROL]
    faltantes = [c for c in requeridas if c not in df.columns]
    if faltantes:
        raise KeyError(
            f"No encontré estas columnas en la planilla: {faltantes}. "
            f"Columnas disponibles: {list(df.columns)}"
        )

    if COL_NOMBRE in df.columns:
        df = df[df[COL_NOMBRE].astype(str).str.strip() != ""]
    return df.reset_index(drop=True)


@st.cache_data(ttl=300)
def cargar_todas_las_planillas() -> pd.DataFrame:
    """Descarga y junta las planillas de las 41 personas en un solo DataFrame,
    agregando de qué encargada viene cada fila. Las que fallan se omiten."""
    indice = cargar_indice()
    partes = []
    for _, fila in indice.iterrows():
        try:
            df_p = cargar_planilla(fila["ID_Planilla"])
            df_p = df_p.copy()
            df_p["_Encargada"] = fila["Persona"]
            partes.append(df_p)
        except Exception:
            continue
    if not partes:
        return pd.DataFrame()
    return pd.concat(partes, ignore_index=True, sort=False)


# ---------------------------------------------------------------------------
# ANÁLISIS
# ---------------------------------------------------------------------------
def meses_con_datos(df: pd.DataFrame) -> list:
    """Meses de la planilla que tienen al menos una marca de asistencia
    (así los meses futuros o vacíos no aparecen como opción)."""
    return [
        m for m in MESES
        if m in df.columns and df[m].astype(str).str.strip().str.upper().eq(MARCA_PRESENTE).any()
    ]


def analizar_voluntarios(df: pd.DataFrame, meses_analizar: list) -> pd.DataFrame:
    """Calcula asistencia por voluntario/a considerando solo los meses indicados
    (puede ser un solo mes, o varios para una vista de año completo)."""
    meses_validos = [m for m in meses_analizar if m in df.columns]
    resultado = df[[COL_NOMBRE, COL_IGLESIA, COL_ROL, COL_RANGO_EDAD]].copy()

    if meses_validos:
        resultado["Meses asistidos"] = df[meses_validos].apply(
            lambda fila: sum(str(v).strip().upper() == MARCA_PRESENTE for v in fila), axis=1
        )
    else:
        resultado["Meses asistidos"] = 0

    total_meses = max(len(meses_validos), 1)
    resultado["Meses considerados"] = len(meses_validos)
    resultado["% asistencia"] = (resultado["Meses asistidos"] / total_meses * 100).round(0)

    return resultado


def _normalizar_rango_edad(valor: str) -> str:
    """Hace calzar el valor real de la celda con una categoría conocida de
    ORDEN_RANGO_EDAD, ignorando mayúsculas/espacios. Si no calza con ninguna,
    deja el texto tal cual (para no perder datos silenciosamente)."""
    valor = str(valor).strip()
    for canonico in ORDEN_RANGO_EDAD:
        if valor.lower() == canonico.lower():
            return canonico
    return valor


def asistencia_por_rango_etario(analisis: pd.DataFrame) -> pd.DataFrame:
    """Para cada rango etario: cantidad de voluntarios, % de asistencia promedio
    y el total de asistencias (meses marcados) acumuladas por ese grupo."""
    datos = analisis.copy()
    datos["Rango etario"] = datos[COL_RANGO_EDAD].apply(_normalizar_rango_edad)

    resumen = datos.groupby("Rango etario").agg(**{
        "% asistencia promedio": ("% asistencia", "mean"),
        "Voluntarios": (COL_NOMBRE, "count"),
        "Total asistencias": ("Meses asistidos", "sum"),
    })
    resumen["% asistencia promedio"] = resumen["% asistencia promedio"].round(0)

    orden_presente = [r for r in ORDEN_RANGO_EDAD if r in resumen.index]
    otros = [r for r in resumen.index if r not in ORDEN_RANGO_EDAD]
    return resumen.reindex(orden_presente + otros)


@st.cache_resource
def _marca_agua_data_uri(ruta: str, opacidad: float):
    """Abre el logo, reduce su transparencia de verdad (a nivel de píxel) y lo
    devuelve como data URI listo para usar en CSS. Se cachea porque procesar
    la imagen en cada rerun sería innecesario."""
    archivo = Path(ruta)
    if not archivo.exists():
        return None

    img = Image.open(archivo).convert("RGBA")
    r, g, b, a = img.split()
    a = a.point(lambda px: int(px * opacidad))
    img.putalpha(a)

    buffer = io.BytesIO()
    img.save(buffer, format="PNG")
    b64 = base64.b64encode(buffer.getvalue()).decode()
    return f"data:image/png;base64,{b64}"


def agregar_marca_agua(ruta: str = LOGO_MARCA_AGUA, opacidad: float = OPACIDAD_MARCA_AGUA):
    """Pone el logo repetido de fondo, en varios tamaños y posiciones. Al ir
    como background-image del propio contenedor (no como capa aparte), queda
    detrás del contenido siempre, sin depender de z-index."""
    data_uri = _marca_agua_data_uri(ruta, opacidad)
    if not data_uri:
        return  # si no está el logo, simplemente no muestra nada (no rompe la app)

    url = f'url("{data_uri}")'
    imagenes = ", ".join([url] * 6)
    tamanos = "110px, 220px, 70px, 170px, 90px, 150px"
    posiciones = "6% 12%, 82% 8%, 25% 50%, 68% 62%, 12% 85%, 88% 88%"

    st.markdown(
        f"""
        <style>
        .stApp {{
            background-repeat: no-repeat;
            background-image: {imagenes};
            background-size: {tamanos};
            background-position: {posiciones};
            background-attachment: fixed;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# INTERFAZ
# ---------------------------------------------------------------------------
try:
    _icono_pagina = Image.open(LOGO_MARCA_AGUA)
except Exception:
    _icono_pagina = "🙌"

st.set_page_config(page_title="ED-Gracia 2026", page_icon=_icono_pagina, layout="wide")
agregar_marca_agua()

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Baloo+2:wght@700;800&display=swap');
    .titulo-ed-gracia {
        font-family: 'Baloo 2', sans-serif;
        font-size: 3rem;
        font-weight: 800;
        background: linear-gradient(90deg, #46F014, #0AD2FA, #FA14A0);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        margin: 0 0 0.5rem 0;
        line-height: 1.1;
    }
    </style>
    <div class="titulo-ed-gracia">ED-Gracia 2026</div>
    """,
    unsafe_allow_html=True,
)

if "modo" not in st.session_state:
    st.session_state.modo = None

col_a, col_b = st.columns(2)
if col_a.button(
    "Seleccionar encargada", use_container_width=True,
    type="primary" if st.session_state.modo == "encargada" else "secondary",
):
    st.session_state.modo = "encargada"
    st.rerun()
if col_b.button(
    "Análisis completo", use_container_width=True,
    type="primary" if st.session_state.modo == "completo" else "secondary",
):
    st.session_state.modo = "completo"
    st.rerun()

if st.session_state.modo is None:
    st.info("Elige una opción para comenzar: revisar a una encargada específica, o ver el análisis de todas.")
    st.stop()

indice = cargar_indice()
if indice.empty:
    st.error("No se pudo cargar la hoja índice. Revisa INDICE_SHEET_ID y los permisos de la service account.")
    st.stop()

# ---------------------------------------------------------------------------
# MODO: una encargada específica
# ---------------------------------------------------------------------------
if st.session_state.modo == "encargada":
    persona_sel = st.selectbox("Selecciona a la HNA encargada", indice["Persona"].tolist())
    spreadsheet_id = indice.loc[indice["Persona"] == persona_sel, "ID_Planilla"].iloc[0]

    with st.spinner(f"Cargando planilla de {persona_sel}..."):
        df = cargar_planilla(spreadsheet_id)

    if df.empty:
        st.warning("No se encontraron voluntarios/as en esta planilla.")
        st.stop()

    meses_disponibles = meses_con_datos(df)
    opciones_mes = [OPCION_ANIO_COMPLETO] + meses_disponibles
    mes_sel = st.selectbox("Selecciona el mes a analizar", opciones_mes, format_func=etiqueta_mes)

    meses_analizar = meses_disponibles if mes_sel == OPCION_ANIO_COMPLETO else [mes_sel]
    analisis = analizar_voluntarios(df, meses_analizar)

    st.metric("Asistencia promedio", f"{analisis['% asistencia'].mean():.0f}%")

    st.subheader(f"Detalle por voluntario/a — {etiqueta_mes(mes_sel)}")
    st.dataframe(
        analisis.sort_values("% asistencia"),
        use_container_width=True,
        column_config={
            "% asistencia": st.column_config.ProgressColumn(
                "% asistencia", format="%d%%", min_value=0, max_value=100
            )
        },
    )

    with st.expander("Ver planilla completa"):
        st.dataframe(df, use_container_width=True)

    st.subheader(f"Participación por rango etario — {etiqueta_mes(mes_sel)}")
    resumen_edad = asistencia_por_rango_etario(analisis)
    if not resumen_edad.empty:
        datos_grafico = resumen_edad.reset_index()
        grafico_edad = (
            alt.Chart(datos_grafico)
            .mark_bar()
            .encode(
                x=alt.X("Rango etario", sort=ORDEN_RANGO_EDAD, title=None),
                y=alt.Y("% asistencia promedio", title="% asistencia promedio"),
            )
        )
        st.altair_chart(grafico_edad, use_container_width=True)
        st.dataframe(resumen_edad, use_container_width=True)
    else:
        st.caption("No hay datos de rango etario para mostrar en esta planilla.")

# ---------------------------------------------------------------------------
# MODO: análisis completo (todas las encargadas)
# ---------------------------------------------------------------------------
else:
    with st.spinner("Cargando las 41 planillas, puede tardar un poco..."):
        df_global = cargar_todas_las_planillas()

    if df_global.empty:
        st.warning("No se pudo cargar ninguna planilla.")
        st.stop()

    meses_globales = meses_con_datos(df_global)
    analisis_global = analizar_voluntarios(df_global, meses_globales)

    iglesias = sorted({
        str(v).strip() for v in df_global.get(COL_IGLESIA, []) if str(v).strip()
    })
    asistencia_perfecta = int((analisis_global["% asistencia"] == 100).sum())

    col1, col2 = st.columns(2)
    col1.metric("Iglesias participantes", len(iglesias))
    col2.metric("Voluntarios con asistencia perfecta", asistencia_perfecta)

    with st.expander(f"Ver listado de iglesias ({len(iglesias)})"):
        for ig in iglesias:
            st.write(f"- {ig}")

    st.subheader("Participación por rango etario — total hasta la fecha")
    resumen_edad_global = asistencia_por_rango_etario(analisis_global)
    if not resumen_edad_global.empty:
        datos_grafico = resumen_edad_global.reset_index()
        grafico_global = (
            alt.Chart(datos_grafico)
            .mark_bar()
            .encode(
                x=alt.X("Rango etario", sort=ORDEN_RANGO_EDAD, title=None),
                y=alt.Y("Voluntarios", title="Total de voluntarios"),
            )
        )
        st.altair_chart(grafico_global, use_container_width=True)
        st.dataframe(resumen_edad_global, use_container_width=True)
    else:
        st.caption("No hay datos de rango etario para mostrar.")

    with st.expander("Resumen por encargada"):
        analisis_global["_Encargada"] = df_global["_Encargada"].values
        resumen_encargadas = analisis_global.groupby("_Encargada").agg(**{
            "Voluntarios": (COL_NOMBRE, "count"),
            "Asistencia promedio (%)": ("% asistencia", "mean"),
        }).round(0)
        st.dataframe(resumen_encargadas, use_container_width=True)

# ---------------------------------------------------------------------------
# PIE DE PÁGINA (común a ambos modos)
# ---------------------------------------------------------------------------
st.divider()
if Path(LOGO_MARCA_AGUA).exists():
    _, col_logo, _ = st.columns([2, 1, 2])
    with col_logo:
        st.image(LOGO_MARCA_AGUA, use_container_width=True)
