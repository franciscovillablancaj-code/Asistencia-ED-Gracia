import streamlit as st
import pandas as pd
import gspread
import io
import altair as alt
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from datetime import date

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
MES_A_NUMERO = {"MR": 3, "AB": 4, "MY": 5, "JN": 6, "JL": 7, "AG": 8,
                "SEP": 9, "OCT": 10, "NOV": 11, "DIC": 12}

# Letra que marca asistencia en la celda del mes
MARCA_PRESENTE = "P"

COLUMNAS_ESPERADAS = [
    COL_REGION, COL_HNA_ENCARGADA, COL_PROGRAMA, COL_NOMBRE,
    COL_IGLESIA, COL_RANGO_EDAD, COL_ROL,
] + MESES

# Orden esperado de las categorías del desplegable "RANGO EDAD"
ORDEN_RANGO_EDAD = ["menor", "10 a 13 años", "14 a 17 años", "18 a 29 años", "más de 30"]

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


# ---------------------------------------------------------------------------
# ANÁLISIS
# ---------------------------------------------------------------------------
def mes_actual_abrev():
    hoy_mes = date.today().month
    for abrev, numero in MES_A_NUMERO.items():
        if numero == hoy_mes:
            return abrev
    return None  # ene/feb no están en la planilla


def meses_transcurridos() -> list:
    hoy_mes = date.today().month
    return [m for m in MESES if MES_A_NUMERO[m] <= hoy_mes]


def analizar_voluntarios(df: pd.DataFrame) -> pd.DataFrame:
    meses_validos = [m for m in meses_transcurridos() if m in df.columns]
    resultado = df[[COL_NOMBRE, COL_IGLESIA, COL_ROL, COL_RANGO_EDAD]].copy()

    if meses_validos:
        resultado["Meses asistidos"] = df[meses_validos].apply(
            lambda fila: sum(str(v).strip().upper() == MARCA_PRESENTE for v in fila), axis=1
        )
    else:
        resultado["Meses asistidos"] = 0

    total_meses = max(len(meses_validos), 1)
    resultado["Meses transcurridos"] = len(meses_validos)
    resultado["% asistencia"] = (resultado["Meses asistidos"] / total_meses * 100).round(0)

    mes_actual = mes_actual_abrev()
    if mes_actual and mes_actual in df.columns:
        resultado["Al día (mes actual)"] = (
            df[mes_actual].astype(str).str.strip().str.upper().eq(MARCA_PRESENTE)
        )
    else:
        resultado["Al día (mes actual)"] = None

    return resultado


def tendencia_mensual(df: pd.DataFrame) -> pd.DataFrame:
    """% de voluntarios que asistieron cada mes transcurrido, para toda la planilla."""
    meses_validos = [m for m in meses_transcurridos() if m in df.columns]
    datos = []
    for m in meses_validos:
        marcados = df[m].astype(str).str.strip().str.upper().eq(MARCA_PRESENTE).sum()
        pct = (marcados / len(df) * 100) if len(df) else 0
        datos.append({"Mes": m, "% asistencia": round(pct)})
    return pd.DataFrame(datos).set_index("Mes")


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


# ---------------------------------------------------------------------------
# INTERFAZ
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Control de planillas de asistencia", layout="wide")
st.title("Control de planillas de asistencia")

indice = cargar_indice()
if indice.empty:
    st.error("No se pudo cargar la hoja índice. Revisa INDICE_SHEET_ID y los permisos de la service account.")
    st.stop()

persona_sel = st.selectbox("Selecciona a la HNA encargada", indice["Persona"].tolist())
spreadsheet_id = indice.loc[indice["Persona"] == persona_sel, "ID_Planilla"].iloc[0]

with st.spinner(f"Cargando planilla de {persona_sel}..."):
    df = cargar_planilla(spreadsheet_id)

if df.empty:
    st.warning("No se encontraron voluntarios/as en esta planilla.")
    st.stop()

analisis = analizar_voluntarios(df)
mes_actual = mes_actual_abrev()

col1, col2, col3, col4 = st.columns(4)
col1.metric("Voluntarios/as", len(df))
col2.metric("Meses transcurridos", len(meses_transcurridos()))
col3.metric("Asistencia promedio", f"{analisis['% asistencia'].mean():.0f}%")
if mes_actual:
    pendientes = int((analisis["Al día (mes actual)"] == False).sum())
    col4.metric(f"Sin marcar en {mes_actual}", pendientes)
else:
    col4.metric("Mes actual", "fuera de temporada")

st.subheader("Participación por rango etario")
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

st.subheader("Detalle por voluntario/a")
st.dataframe(
    analisis.sort_values("% asistencia"),
    use_container_width=True,
    column_config={
        "% asistencia": st.column_config.ProgressColumn(
            "% asistencia", format="%d%%", min_value=0, max_value=100
        )
    },
)

st.subheader("Tendencia mensual de asistencia")
tendencia = tendencia_mensual(df)
if not tendencia.empty:
    st.line_chart(tendencia)
else:
    st.caption("Aún no hay meses transcurridos este año para mostrar una tendencia.")

with st.expander("Ver planilla completa"):
    st.dataframe(df, use_container_width=True)

with st.expander("Resumen de las 41 personas (puede tardar unos segundos)"):
    if st.button("Cargar resumen general"):
        resumen = []
        for _, row in indice.iterrows():
            try:
                df_p = cargar_planilla(row["ID_Planilla"])
                a_p = analizar_voluntarios(df_p)
                resumen.append({
                    "Persona": row["Persona"],
                    "Voluntarios": len(df_p),
                    "Asistencia promedio (%)": round(a_p["% asistencia"].mean(), 0) if len(df_p) else 0,
                    f"Sin marcar en {mes_actual or '-'}": int((a_p["Al día (mes actual)"] == False).sum()) if mes_actual else "N/D",
                })
            except Exception as e:
                resumen.append({"Persona": row["Persona"], "error": str(e)})
        st.dataframe(pd.DataFrame(resumen), use_container_width=True)
