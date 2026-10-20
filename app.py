
import io
from datetime import date, datetime

import pandas as pd
import streamlit as st


# =========================================================
# CONFIGURACIÓN Y ESTILO
# =========================================================

st.set_page_config(
    page_title="DataCleaner",
    page_icon="🧹",
    layout="wide"
)

st.markdown(
    """
    <style>
    .block-container {
        max-width: 1400px;
        padding-top: 1rem;
        padding-bottom: 3rem;
    }

        h1 {
            letter-spacing: -1px;
        }

        .dc-subtitle {
            color: #77817c;
            font-size: 1.05rem;
            margin-top: -12px;
            margin-bottom: 28px;
        }

        .dc-eyebrow {
            color: #47765b;
            font-size: 0.78rem;
            font-weight: 700;
            letter-spacing: 1.5px;
            text-transform: uppercase;
            margin-bottom: 5px;
        }



        div[data-testid="stMetric"] {
            background: #edf4ee !important;
            border: 1px solid #d3e2d5;
            padding: 10px 14px;
            border-radius: 12px;
            width: fit-content;
            min-width: 150px;
        }

        div[data-testid="stMetric"] [data-testid="stMetricLabel"],
        div[data-testid="stMetric"] [data-testid="stMetricLabel"] *,
        div[data-testid="stMetric"] label,
        div[data-testid="stMetric"] label * {
            color: #405548 !important;
            opacity: 1 !important;
        }

        div[data-testid="stMetric"] [data-testid="stMetricValue"],
        div[data-testid="stMetric"] [data-testid="stMetricValue"] * {
            color: #183d2b !important;
            opacity: 1 !important;
        }

        div[data-testid="stMetric"] p {
            color: #405548 !important;
        }

    </style>
    """,
    unsafe_allow_html=True
)


# =========================================================
# FUNCIONES AUXILIARES
# =========================================================

def contar_vacios(df):
    """Cuenta celdas nulas y cadenas vacías o con espacios."""
    total = int(df.isna().sum().sum())

    for columna in df.select_dtypes(
        include=["object", "string"]
    ).columns:
        total += int(
            df[columna].map(
                lambda valor: (
                    isinstance(valor, str)
                    and not valor.strip()
                )
            ).sum()
        )

    return total


def normalizar_vacios(df):
    """Convierte cadenas vacías o con espacios en valores nulos."""
    resultado = df.copy()

    for columna in resultado.select_dtypes(
        include=["object", "string"]
    ).columns:
        resultado[columna] = resultado[columna].replace(
            r"^\s*$",
            pd.NA,
            regex=True
        )

    return resultado


def convertir_a_fecha(valor):
    """Convierte formatos conocidos de fecha sin intercambiar día y mes."""

    if pd.isna(valor):
        return pd.NaT

    if isinstance(valor, (datetime, date, pd.Timestamp)):
        return pd.Timestamp(valor).date()

    texto = str(valor).strip()

    if not texto:
        return pd.NaT

    formatos = [
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d",
        "%Y.%m.%d",
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%d.%m.%Y",
        "%b %d, %Y",
        "%B %d, %Y",
        "%d %b %Y",
        "%d %B %Y",
    ]

    for formato in formatos:
        try:
            return datetime.strptime(
                texto,
                formato
            ).date()
        except ValueError:
            continue

    return pd.NaT


def formatear_fechas_vista(df):
    """
    Prepara una copia visual de la tabla.
    Las fechas se muestran sin hora, sin modificar los datos reales.
    """
    resultado = df.copy()

    for columna in resultado.columns:
        serie = resultado[columna]

        es_fecha = (
            pd.api.types.is_datetime64_any_dtype(serie.dtype)
            or serie.map(
                lambda valor: isinstance(
                    valor,
                    (date, datetime, pd.Timestamp)
                )
            ).any()
        )

        if es_fecha:
            resultado[columna] = serie.map(
                lambda valor: (
                    valor.strftime("%Y-%m-%d")
                    if isinstance(
                        valor,
                        (date, datetime, pd.Timestamp)
                    )
                    else valor
                )
            )

    return resultado


def cargar_archivo(archivo):
    """Lee un archivo CSV o Excel."""
    if archivo.name.lower().endswith(".csv"):
        try:
            return pd.read_csv(
                archivo,
                encoding="utf-8-sig"
            )
        except UnicodeDecodeError:
            archivo.seek(0)
            return pd.read_csv(
                archivo,
                encoding="latin-1"
            )

    return pd.read_excel(archivo)


def exportar_excel(df):
    """Genera un Excel con las fechas visualizadas sin hora."""
    salida = io.BytesIO()

    with pd.ExcelWriter(
        salida,
        engine="openpyxl",
        date_format="YYYY-MM-DD",
        datetime_format="YYYY-MM-DD"
    ) as writer:
        df.to_excel(
            writer,
            index=False,
            sheet_name="Datos limpios"
        )

        hoja = writer.sheets["Datos limpios"]

        for fila in hoja.iter_rows():
            for celda in fila:
                if celda.is_date:
                    celda.number_format = "yyyy-mm-dd"

        # Ajustar el ancho de las columnas para facilitar la lectura.
        for columna in hoja.columns:
            letra = columna[0].column_letter

            ancho_maximo = max(
                len(str(celda.value or ""))
                for celda in columna
            )

            hoja.column_dimensions[letra].width = min(
                max(ancho_maximo + 2, 12),
                35
            )

        hoja.freeze_panes = "A2"
        hoja.auto_filter.ref = hoja.dimensions

    return salida.getvalue()


# =========================================================
# ENCABEZADO
# =========================================================
st.markdown(
    '<div class="dc-eyebrow">'
    'PROYECTO ACADÉMICO · CIENCIA DE DATOS'
    '</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div style="color: #77817c; font-size: 0.9rem; '
    'margin-bottom: 8px;">'
    'Desarrollado por Emmanuel Sotelo · 2026'
    '</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="dc-eyebrow">WEB APP PARA LA LIMPIEZA DE DATOS</div>',
    unsafe_allow_html=True
)

st.title("🧹 DataCleaner")

st.markdown(
    '<div class="dc-subtitle">'
    'Prepará tus archivos, corregí inconsistencias y exportá '
    'datos listos para trabajar.'
    '</div>',
    unsafe_allow_html=True
)


# =========================================================
# 1. CARGA DEL ARCHIVO
# =========================================================

st.subheader("01 · Archivo de origen")

st.write(
    "Comenzá cargando el archivo que querés revisar. "
    "Podés trabajar con archivos CSV o Excel."
)

archivo = st.file_uploader(
    "Seleccionar archivo",
    type=["xlsx", "csv"],
    help="Formatos admitidos: .csv y .xlsx"
)


if archivo is None:
    st.info(
        "Tu espacio de trabajo aparecerá acá cuando cargues un archivo."
    )

    st.markdown("---")
    st.caption(
        "DataCleaner · Cargá, limpiá y exportá tus datos."
    )

    st.stop()


# =========================================================
# LECTURA Y PREPARACIÓN
# =========================================================

try:
    dataframe_original = cargar_archivo(archivo)

    if dataframe_original.shape[1] == 0:
        st.error("El archivo no contiene columnas.")
        st.stop()

    dataframe = normalizar_vacios(dataframe_original)

except Exception as error:
    st.error(f"No se pudo leer el archivo: {error}")
    st.stop()


columnas_originales = list(dataframe.columns)

st.success(f"Archivo cargado: {archivo.name}")


# =========================================================
# 2. DIAGNÓSTICO DEL ARCHIVO
# =========================================================

st.subheader("02 · Diagnóstico")

st.write(
    "Un resumen rápido del contenido original, antes de aplicar cambios."
)

col1, col2, col3 = st.columns(3)

with col1:
    st.metric(
        "Filas",
        f"{len(dataframe_original):,}"
    )

with col2:
    st.metric(
        "Columnas",
        f"{len(dataframe_original.columns):,}"
    )

with col3:
    st.metric(
        "Valores vacíos",
        f"{contar_vacios(dataframe_original):,}"
    )


with st.expander("Ver tabla original", expanded=False):
    st.caption(
        "Esta tabla es una vista del archivo cargado. "
        "Las operaciones se aplican sobre una copia."
    )

    st.dataframe(
        formatear_fechas_vista(dataframe_original),
        use_container_width=True,
        hide_index=True
    )


st.divider()


# =========================================================
# 3. HERRAMIENTAS DE LIMPIEZA
# =========================================================

st.subheader("03 · Herramientas de limpieza")

st.write(
    "Abrí solamente las herramientas que necesitás. "
    "Las opciones que no actives no modificarán los datos."
)


# ---------------------------------------------------------
# 3.1 COLUMNAS
# ---------------------------------------------------------

with st.expander("📋  Columnas y nombres", expanded=True):

    st.caption(
        "Eliminá columnas innecesarias, elegí cuáles conservar "
        "o cambiá el nombre de una columna."
    )

    modo_columnas = st.selectbox(
        "¿Qué querés hacer con las columnas?",
        [
            "Conservar todas las columnas",
            "Eliminar columnas seleccionadas",
            "Conservar solo las columnas seleccionadas"
        ],
        key="modo_columnas"
    )

    columnas_seleccionadas = []

    if modo_columnas != "Conservar todas las columnas":
        columnas_seleccionadas = st.multiselect(
            "Columnas involucradas",
            options=columnas_originales,
            key="columnas_seleccionadas"
        )

    if modo_columnas == "Eliminar columnas seleccionadas":
        columnas_finales = [
            columna
            for columna in columnas_originales
            if columna not in columnas_seleccionadas
        ]

    elif (
        modo_columnas == "Conservar solo las columnas seleccionadas"
        and columnas_seleccionadas
    ):
        columnas_finales = [
            columna
            for columna in columnas_originales
            if columna in columnas_seleccionadas
        ]

    else:
        columnas_finales = columnas_originales.copy()

    if not columnas_finales:
        st.warning(
            "No se pueden eliminar todas las columnas. "
            "Se conservarán las originales."
        )
        columnas_finales = columnas_originales.copy()

    st.markdown("**Renombrar una columna**")

    columna_renombrar = st.selectbox(
        "Columna que querés cambiar",
        options=["No renombrar"] + columnas_finales,
        key="columna_renombrar"
    )

    nuevo_nombre = ""

    if columna_renombrar != "No renombrar":
        nuevo_nombre = st.text_input(
            "Nuevo nombre de la columna",
            placeholder="Ej.: nombre_cliente",
            key="nuevo_nombre"
        ).strip()

        if (
            nuevo_nombre
            and nuevo_nombre in columnas_finales
            and nuevo_nombre != columna_renombrar
        ):
            st.warning(
                "Ese nombre ya existe. Elegí uno diferente."
            )
            nuevo_nombre = ""


# Preparar el esquema de columnas que utilizarán las herramientas.
esquema = dataframe[columnas_finales].copy()

if columna_renombrar != "No renombrar" and nuevo_nombre:
    esquema = esquema.rename(
        columns={columna_renombrar: nuevo_nombre}
    )

columnas_trabajo = list(esquema.columns)

columnas_texto = [
    columna
    for columna in columnas_trabajo
    if (
        pd.api.types.is_object_dtype(esquema[columna].dtype)
        or pd.api.types.is_string_dtype(esquema[columna].dtype)
    )
]


# ---------------------------------------------------------
# 3.2 TEXTOS
# ---------------------------------------------------------

with st.expander("🔤  Limpieza de textos", expanded=False):

    st.caption(
        "Quitá espacios sobrantes o normalizá el uso de "
        "mayúsculas y minúsculas."
    )

    
    

    limpiar_espacios = st.checkbox(
        "Eliminar espacios al principio y al final",
        help="No elimina los espacios que separan palabras.",
        key="limpiar_espacios"
    )


    normalizar_texto = st.selectbox(
        "Formato de texto",
        [
            "No modificar mayúsculas/minúsculas",
            "Convertir a MAYÚSCULAS",
            "Convertir a minúsculas",
            "Convertir a Nombre Propio"
        ],
        key="normalizar_texto"
    )

    columnas_texto_seleccionadas = []

    if (
        normalizar_texto != "No modificar mayúsculas/minúsculas"
        and columnas_texto
    ):
        columnas_texto_seleccionadas = st.multiselect(
            "Aplicar el formato a estas columnas",
            options=columnas_texto,
            default=[],
            key="columnas_texto"
        )

        st.caption(
            "Solo se modificarán las columnas seleccionadas. "
            "Podés limpiar nombres sin alterar correos u otros datos."
        )



# ---------------------------------------------------------
# 3.3 TIPOS DE DATOS
# ---------------------------------------------------------

with st.expander("🔄  Tipos de datos", expanded=False):

    st.caption(
        "Convertí una o varias columnas a texto, número o fecha."
    )

    convertir_tipos = st.checkbox(
        "Activar conversión de tipos",
        key="convertir_tipos"
    )

    columnas_convertir = []
    tipo_destino = "Texto"

    if convertir_tipos:
        columnas_convertir = st.multiselect(
            "Columnas que querés convertir",
            options=columnas_trabajo,
            key="columnas_convertir"
        )

        tipo_destino = st.selectbox(
            "Tipo de destino",
            ["Texto", "Número", "Fecha"],
            key="tipo_destino"
        )

        if tipo_destino == "Fecha":
            st.info(
                "Se admiten formatos de fecha conocidos, como "
                "AAAA-MM-DD, DD/MM/AAAA y nombres de meses en inglés. "
                "Las fechas no válidas quedarán vacías. "
                "La hora no se conservará."
            )


# ---------------------------------------------------------
# 3.4 VALORES VACÍOS
# ---------------------------------------------------------

with st.expander("🕳️  Valores vacíos", expanded=False):

    st.caption(
        "Elegí qué hacer con las celdas vacías o que contienen "
        "solamente espacios."
    )

    accion_vacios = st.selectbox(
        "Acción",
        [
            "No modificar valores vacíos",
            "Completar valores vacíos",
            "Eliminar filas con valores vacíos"
        ],
        key="accion_vacios"
    )

    valor_relleno = "Sin dato"

    if accion_vacios == "Completar valores vacíos":
        valor_relleno = st.text_input(
            "Texto para completar las celdas",
            value="Sin dato",
            key="valor_relleno"
        )

        st.caption(
            "Se utilizará el mismo valor en todas las celdas vacías, "
            "incluso si pertenecen a columnas numéricas o de fecha."
        )


# ---------------------------------------------------------
# 3.5 DUPLICADOS
# ---------------------------------------------------------

with st.expander("🧬  Filas duplicadas", expanded=False):

    st.caption(
        "Detectá registros repetidos y conservá la primera aparición."
    )

    eliminar_duplicados = st.checkbox(
        "Activar eliminación de duplicados",
        key="eliminar_duplicados"
    )

    criterio_duplicados = "Todas las columnas"
    columnas_duplicados = []

    if eliminar_duplicados:
        criterio_duplicados = st.radio(
            "¿Cómo identificar registros repetidos?",
            [
                "Todas las columnas",
                "Columnas seleccionadas"
            ],
            key="criterio_duplicados"
        )

        if criterio_duplicados == "Columnas seleccionadas":
            columnas_duplicados = st.multiselect(
                "Columnas que se compararán",
                options=columnas_trabajo,
                key="columnas_duplicados"
            )

            st.caption(
                "Dos filas se considerarán duplicadas si coinciden "
                "en las columnas elegidas."
            )

# ---------------------------------------------------------
# 3.6 REEMPLAZOS
# ---------------------------------------------------------

with st.expander("🔁 Reemplazar valores", expanded=False):

    st.caption(
        "Reemplazá varios valores de una columna mediante reglas "
        "que podés agregar o quitar."
    )

    columna_reemplazo = st.selectbox(
        "Columna donde reemplazar",
        options=columnas_trabajo,
        key="columna_reemplazo"
    )

    ignorar_mayusculas = st.checkbox(
        "Ignorar mayúsculas y minúsculas",
        value=True,
        key="ignorar_mayusculas"
    )

    if "reglas_reemplazo" not in st.session_state:
        st.session_state.reglas_reemplazo = [
            {"id": 0, "original": "", "nuevo": ""}
        ]

    if "siguiente_id_reemplazo" not in st.session_state:
        st.session_state.siguiente_id_reemplazo = 1

    for regla in st.session_state.reglas_reemplazo:
        regla_id = regla["id"]

        col_original, col_nuevo, col_eliminar = st.columns(
            [2, 2, 0.5]
        )

        with col_original:
            st.text_input(
                "Valor original",
                key=f"reemplazo_original_{regla_id}"
            )

        with col_nuevo:
            st.text_input(
                "Reemplazar por",
                key=f"reemplazo_nuevo_{regla_id}"
            )

        with col_eliminar:
            st.write("")
            st.write("")
            if st.button("✕", key=f"eliminar_regla_{regla_id}"):
                st.session_state.reglas_reemplazo = [
                    r for r in st.session_state.reglas_reemplazo
                    if r["id"] != regla_id
                ]
                st.rerun()

    if st.button("➕ Agregar regla", key="agregar_regla"):
        nuevo_id = st.session_state.siguiente_id_reemplazo

        st.session_state.reglas_reemplazo.append({
            "id": nuevo_id,
            "original": "",
            "nuevo": ""
        })

        st.session_state.siguiente_id_reemplazo += 1
        st.rerun()


st.divider()


# =========================================================
# 4. APLICAR LAS OPERACIONES
# =========================================================

st.subheader("04 · Resultado de la limpieza")


dataframe_limpio = dataframe[columnas_finales].copy()

if columna_renombrar != "No renombrar" and nuevo_nombre:
    dataframe_limpio = dataframe_limpio.rename(
        columns={columna_renombrar: nuevo_nombre}
    )


# Contadores del resumen.
cambios_espacios = 0
valores_completados = 0
filas_eliminadas_vacios = 0
duplicados_eliminados = 0
conversiones_fallidas = 0
valores_reemplazados = 0


# ---------------------------------------------------------
# Limpiar espacios
# ---------------------------------------------------------

if limpiar_espacios:
    for columna in dataframe_limpio.columns:
        serie = dataframe_limpio[columna]

        if (
            pd.api.types.is_object_dtype(serie.dtype)
            or pd.api.types.is_string_dtype(serie.dtype)
        ):
            limpios = serie.str.strip()

            cambios_espacios += int(
                (
                    serie.notna()
                    & serie.ne(limpios).fillna(False)
                ).sum()
            )

            dataframe_limpio[columna] = limpios
# ---------------------------------------------------------
# Normalizar textos
# ---------------------------------------------------------

for columna in columnas_texto_seleccionadas:
    if columna not in dataframe_limpio.columns:
        continue

    serie = dataframe_limpio[columna].astype("string")

    if normalizar_texto == "Convertir a MAYÚSCULAS":
        dataframe_limpio[columna] = serie.str.upper()

    elif normalizar_texto == "Convertir a minúsculas":
        dataframe_limpio[columna] = serie.str.lower()

    elif normalizar_texto == "Convertir a Nombre Propio":
        dataframe_limpio[columna] = serie.str.title()


# ---------------------------------------------------------
# Convertir tipos de datos
# ---------------------------------------------------------

if convertir_tipos:
    for columna in columnas_convertir:
        serie_original = dataframe_limpio[columna]
        no_vacios = serie_original.notna()

        if tipo_destino == "Texto":
            convertida = serie_original.astype("string")

        elif tipo_destino == "Número":
            convertida = pd.to_numeric(
                serie_original,
                errors="coerce"
            )

        else:
            convertida = serie_original.map(convertir_a_fecha)

        conversiones_fallidas += int(
            (
                no_vacios
                & convertida.isna()
            ).sum()
        )

        dataframe_limpio[columna] = convertida


# ---------------------------------------------------------
# Reemplazar valores
# ---------------------------------------------------------

columna_objetivo = st.session_state.get("columna_reemplazo")

if columna_objetivo in dataframe_limpio.columns:
    ignorar_mayusculas_actual = st.session_state.get(
        "ignorar_mayusculas",
        True
    )

    for regla in st.session_state.get("reglas_reemplazo", []):
        regla_id = regla["id"]

        valor_original = st.session_state.get(
            f"reemplazo_original_{regla_id}", ""
        )

        valor_nuevo = st.session_state.get(
            f"reemplazo_nuevo_{regla_id}", ""
        )

        if not valor_original.strip():
            continue

        if not valor_nuevo.strip():
            continue

        serie = dataframe_limpio[columna_objetivo]

        if ignorar_mayusculas_actual:
            mascara = (
                serie.astype("string")
                .str.strip()
                .str.casefold()
                .eq(valor_original.strip().casefold())
                .fillna(False)
            )
        else:
            mascara = (
                serie.astype("string")
                .str.strip()
                .eq(valor_original.strip())
                .fillna(False)
            )

        valores_reemplazados += int(mascara.sum())

        if not pd.api.types.is_object_dtype(
            dataframe_limpio[columna_objetivo].dtype
        ):
            dataframe_limpio[columna_objetivo] = (
                dataframe_limpio[columna_objetivo].astype("object")
            )

        dataframe_limpio.loc[
            mascara, columna_objetivo
        ] = valor_nuevo


# ---------------------------------------------------------
# Completar o eliminar valores vacíos
# ---------------------------------------------------------

if accion_vacios == "Completar valores vacíos":
    if valor_relleno.strip():
        valores_completados = contar_vacios(dataframe_limpio)

        dataframe_limpio = dataframe_limpio.fillna(
            valor_relleno
        )
    else:
        st.warning(
            "Escribí un valor para completar las celdas vacías."
        )

elif accion_vacios == "Eliminar filas con valores vacíos":
    filas_antes = len(dataframe_limpio)

    dataframe_limpio = dataframe_limpio.dropna()

    filas_eliminadas_vacios = (
        filas_antes - len(dataframe_limpio)
    )


# ---------------------------------------------------------
# Eliminar duplicados
# ---------------------------------------------------------

if eliminar_duplicados:
    if criterio_duplicados == "Todas las columnas":
        subset = None
    else:
        subset = columnas_duplicados

    if subset is None or subset:
        filas_antes = len(dataframe_limpio)

        dataframe_limpio = dataframe_limpio.drop_duplicates(
            subset=subset,
            keep="first"
        )

        duplicados_eliminados = (
            filas_antes - len(dataframe_limpio)
        )
    else:
        st.warning(
            "Seleccioná al menos una columna para identificar "
            "los duplicados."
        )



# =========================================================
# 6. VISTA PREVIA
# =========================================================

st.markdown("#### Tabla resultante")

st.caption(
    "Esta es una vista previa de los datos que se exportarán. "
    "Las fechas se muestran sin la hora."
)

st.dataframe(
    formatear_fechas_vista(dataframe_limpio),
    use_container_width=True,
    hide_index=True
)


# =========================================================
# 7. RESUMEN DE OPERACIONES
# =========================================================

with st.expander("Ver resumen de las operaciones", expanded=False):

    col1, col2 = st.columns(2)

    with col1:
        st.write("**Correcciones realizadas**")
        st.write(f"Espacios corregidos: {cambios_espacios}")
        st.write(f"Valores reemplazados: {valores_reemplazados}")
        st.write(f"Valores vacíos completados: {valores_completados}")
        st.write(f"Duplicados eliminados: {duplicados_eliminados}")

    with col2:
        st.write("**Cambios en los registros**")
        st.write(
            f"Filas eliminadas por valores vacíos: "
            f"{filas_eliminadas_vacios}"
        )
        st.write(
            f"Conversiones que no se pudieron realizar: "
            f"{conversiones_fallidas}"
        )


# =========================================================
# 8. EXPORTAR
# =========================================================

st.divider()

st.subheader("05 · Exportar archivo")

st.write(
    "Descargá una copia limpia para seguir trabajando "
    "en Excel u otra herramienta."
)

formato = st.radio(
    "Formato de descarga",
    ["Excel (.xlsx)", "CSV (.csv)"],
    horizontal=True,
    key="formato_exportacion"
)

if formato == "Excel (.xlsx)":
    try:
        archivo_excel = exportar_excel(dataframe_limpio)

        st.download_button(
            "📥 Descargar Excel limpio",
            data=archivo_excel,
            file_name="datos_limpios.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True
        )

    except Exception as error:
        st.error(f"No se pudo generar el Excel: {error}")

else:
    csv_limpio = dataframe_limpio.to_csv(
        index=False,
        date_format="%Y-%m-%d"
    ).encode("utf-8-sig")

    st.download_button(
        "📥 Descargar CSV limpio",
        data=csv_limpio,
        file_name="datos_limpios.csv",
        mime="text/csv",
        use_container_width=True
    )


st.markdown("---")

st.caption(
    "DataCleaner · Limpieza de datos sencilla, clara y controlada."
)
