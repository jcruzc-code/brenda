# ============================================================
# STREAMLIT APP — CHECK LIST SUPERVISIÓN
# Mantiene las tablas A-J del análisis original
# Agrega: filtros globales + matriz Cliente-Unidad x Fecha
# ============================================================

import io
import re
from datetime import datetime, date

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


# ─────────────────────────────────────────────────────────────
# 1. CONFIGURACIÓN GENERAL
# ─────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Dashboard Check List Supervisión",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📋 Dashboard Check List Supervisión")
st.caption("Análisis operativo con filtros globales, tablas A-J y matriz Cliente-Unidad por Fecha.")

st.markdown(
    """
    <style>
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h2 {
        border-bottom: 1px solid rgba(255, 255, 255, 0.18);
        padding-bottom: 0.35rem;
        margin-bottom: 0.75rem;
    }
    [data-testid="stSidebar"] label {
        font-weight: 600;
    }
    [data-testid="stSidebar"] [data-baseweb="input"],
    [data-testid="stSidebar"] [data-baseweb="select"],
    [data-testid="stSidebar"] [data-baseweb="tag"] {
        border-radius: 10px;
    }
    [data-testid="stSidebar"] [data-baseweb="input"] {
        background: rgba(255, 255, 255, 0.02);
        border: 1px solid rgba(255, 255, 255, 0.12);
    }
    [data-testid="stSidebar"] .stDateInput {
        padding: 0.25rem;
        border-radius: 10px;
        background: rgba(0, 0, 0, 0.22);
    }
    </style>
    """,
    unsafe_allow_html=True
)

ANIO_MIN = 2020
ANIO_MAX = 2027

# Asumo que unidad = sede.
# Si tu archivo tiene otra columna de unidad, cambia esta variable.
UNIDAD_COL = "sede"
FECHA_BASE_COL = "fecha_base"


# ─────────────────────────────────────────────────────────────
# 2. UTILIDADES
# ─────────────────────────────────────────────────────────────

def normalizar_texto(x):
    if pd.isna(x):
        return np.nan
    x = str(x).strip()
    x = re.sub(r"\s+", " ", x)
    return x if x else np.nan


def normalizar_upper(x):
    if pd.isna(x):
        return np.nan
    x = str(x).strip().upper()
    x = re.sub(r"\s+", " ", x)
    return x if x else np.nan


def parse_fecha(s):
    return pd.to_datetime(s, errors="coerce", dayfirst=True)


def parse_marca_temporal(s):
    # Primero intenta formato tipo Excel/Forms m/d/Y H:M:S.
    dt = pd.to_datetime(s, errors="coerce", dayfirst=False)
    if isinstance(dt, pd.Series):
        faltantes = dt.isna()
        if faltantes.any():
            dt_alt = pd.to_datetime(s, errors="coerce", dayfirst=True)
            dt.loc[faltantes] = dt_alt.loc[faltantes]
    return dt


def limpiar_cliente(x):
    if pd.isna(x):
        return np.nan
    x = str(x).strip()
    x = x.split("/")[0].strip()
    x = re.sub(r"\s+", " ", x)
    x = x.upper()
    return x if x else np.nan


def parse_operarios(val):
    if pd.isna(val):
        return np.nan
    match = re.search(r"\d+", str(val))
    return int(match.group()) if match else np.nan


def tiene_problema(val):
    """
    1 = tiene problema
    0 = no tiene problema
    NaN = sin respuesta

    Lógica pensada para texto libre.
    Evita marcar como "sin problema" frases como:
    - NO CUENTA CON...
    - NO TIENE...
    - FALTA...
    """
    if pd.isna(val):
        return np.nan

    v = str(val).strip().upper()
    v = re.sub(r"\s+", " ", v)

    if v == "":
        return np.nan

    frases_con_problema = [
        "NO CUENTA",
        "NO TIENE",
        "NO PRESENTÓ",
        "NO PRESENTO",
        "FALTA",
        "FALTAN",
        "PENDIENTE",
        "PENDIENTES",
        "INCOMPLETO",
        "INCOMPLETA",
        "VENCIDO",
        "VENCIDA",
        "RECLAMO",
        "QUEJA",
        "DEUDA",
        "ADEUDA",
        "ATRASO",
        "DEMORA",
        "MAL ESTADO",
        "SIN STOCK",
        "SIN EPP",
        "SIN MATERIAL",
        "SIN UNIFORME",
        "OBSERVACIÓN",
        "OBSERVACION",
        "INCIDENTE",
        "ACCIDENTE",
    ]

    frases_sin_problema = [
        "NO",
        "NO.",
        "NO PRESENTA",
        "NO PRESENTA PROBLEMAS",
        "NO HAY",
        "NO HAY OBSERVACIONES",
        "NINGUNO",
        "NINGUNA",
        "SIN NOVEDAD",
        "SIN OBSERVACIONES",
        "SIN PROBLEMA",
        "SIN PROBLEMAS",
        "CONFORME",
        "OK",
        "TODO OK",
        "TODO CONFORME",
        "N/A",
        "NA",
        "NO APLICA",
    ]

    for frase in frases_con_problema:
        if frase in v:
            return 1

    for frase in frases_sin_problema:
        if v == frase or v.startswith(frase + " "):
            return 0

    return 1


def lista_documentos(x):
    if pd.isna(x):
        return []
    return [i.strip().upper() for i in str(x).split(",") if i.strip()]


def limpiar_nombre_hoja(nombre):
    nombre = re.sub(r"[\[\]\:\*\?\/\\]", "_", nombre)
    return nombre[:31]


def convertir_fechas_a_str(df_export):
    out = df_export.copy()
    for col in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[col]):
            out[col] = out[col].dt.strftime("%Y-%m-%d")
    return out


def descargar_excel(tabs_dict):
    """
    Recibe diccionario:
    {
        "NombreHoja": dataframe
    }
    Devuelve bytes de Excel.
    """
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet_name, data in tabs_dict.items():
            safe_name = limpiar_nombre_hoja(sheet_name)
            if data is None:
                continue

            export_df = data.copy()

            if isinstance(export_df.index, pd.MultiIndex) or export_df.index.name is not None:
                export_df = export_df.reset_index()

            export_df = convertir_fechas_a_str(export_df)
            export_df.to_excel(writer, sheet_name=safe_name, index=False)

            workbook = writer.book
            worksheet = writer.sheets[safe_name]

            header_format = workbook.add_format({
                "bold": True,
                "bg_color": "#D9EAF7",
                "border": 1
            })

            for col_num, value in enumerate(export_df.columns.values):
                worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, 18)

            worksheet.freeze_panes(1, 0)
            if len(export_df.columns) > 0:
                worksheet.autofilter(0, 0, max(len(export_df), 1), len(export_df.columns) - 1)

    return output.getvalue()


# ─────────────────────────────────────────────────────────────
# 3. CARGA DE DATOS
# ─────────────────────────────────────────────────────────────

def leer_archivo_subido(uploaded_file):
    name = uploaded_file.name.lower()

    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file)

    if name.endswith(".xlsx") or name.endswith(".xls"):
        return pd.read_excel(uploaded_file)

    raise ValueError("Formato no soportado. Sube un .xlsx, .xls o .csv")


def leer_google_sheet(sheet_name, worksheet_name=None):
    """
    Opcional para Streamlit Cloud.
    Requiere secrets con este formato:

    [gcp_service_account]
    type = "service_account"
    project_id = "..."
    private_key_id = "..."
    private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
    client_email = "..."
    client_id = "..."
    auth_uri = "https://accounts.google.com/o/oauth2/auth"
    token_uri = "https://oauth2.googleapis.com/token"
    auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
    client_x509_cert_url = "..."
    """
    import gspread

    creds_dict = dict(st.secrets["gcp_service_account"])
    gc = gspread.service_account_from_dict(creds_dict)
    spreadsheet = gc.open(sheet_name)

    if worksheet_name:
        ws = spreadsheet.worksheet(worksheet_name)
    else:
        ws = spreadsheet.sheet1

    data = ws.get_all_values()
    return pd.DataFrame(data[1:], columns=data[0])


@st.cache_data(show_spinner=False)
def preparar_datos(df_raw):
    df_raw = df_raw.copy()
    df_raw.columns = [str(c).strip() for c in df_raw.columns]

    # Mapeo por posición, igual que tu código original.
    # Esto evita depender de nombres largos de Google Forms.
    COLUMN_NAMES_BY_POSITION = {
        0:  "marca_temporal",
        1:  "email_responsable",
        2:  "fecha_visita",
        3:  "responsable",
        4:  "cliente_raw",
        5:  "sede",
        6:  "cantidad_operarios",
        7:  "hora_llegada",
        8:  "motivo_visita",
        9:  "colaborador_entrevistado",
        10: "status_documentacion",
        11: "problema_materiales",
        12: "problema_pagos",
        13: "problema_destaque",
        14: "problema_ssoma",
        15: "evidencias",
        16: "fecha_siguiente_visita",
        17: "nombre_contacto_cliente",
        18: "cargo_contacto_cliente",
        19: "telefono_contacto_cliente",
        20: "puntuacion_cliente",
        21: "seguimiento_por",
        22: "oportunidades_mejora",
    }

    rename_map = {}
    for idx, new_name in COLUMN_NAMES_BY_POSITION.items():
        if idx < len(df_raw.columns):
            rename_map[df_raw.columns[idx]] = new_name

    df = df_raw.rename(columns=rename_map).copy()

    # Asegurar columnas mínimas por si falta alguna
    columnas_necesarias = list(COLUMN_NAMES_BY_POSITION.values())
    for col in columnas_necesarias:
        if col not in df.columns:
            df[col] = np.nan

    # Limpieza de texto
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].apply(normalizar_texto)

    # Fechas
    df["fecha_visita"] = parse_fecha(df["fecha_visita"])
    df["marca_temporal"] = parse_marca_temporal(df["marca_temporal"])
    df["fecha_siguiente_visita"] = parse_fecha(df["fecha_siguiente_visita"])

    # Fecha base confiable para filtro: marca_temporal sin hora.
    df[FECHA_BASE_COL] = df["marca_temporal"].dt.normalize()

    # Filtro de fechas extremas usando fecha base.
    df = df[
        df[FECHA_BASE_COL].dt.year.between(ANIO_MIN, ANIO_MAX, inclusive="both")
        | df[FECHA_BASE_COL].isna()
    ].copy()

    # Campos limpios
    df["cliente"] = df["cliente_raw"].apply(limpiar_cliente)
    df["responsable_norm"] = df["responsable"].apply(normalizar_upper)
    df["sede"] = df["sede"].apply(normalizar_upper)

    if UNIDAD_COL not in df.columns:
        df["unidad"] = df["sede"]
    else:
        df["unidad"] = df[UNIDAD_COL].apply(normalizar_upper)

    df["motivo_visita_norm"] = df["motivo_visita"].apply(normalizar_upper)
    df["seguimiento_por_norm"] = df["seguimiento_por"].apply(normalizar_upper)

    # Numéricos
    df["cantidad_operarios_num"] = df["cantidad_operarios"].apply(parse_operarios)
    df["puntuacion_cliente"] = pd.to_numeric(df["puntuacion_cliente"], errors="coerce")

    # Tiempo
    df["anio"] = df[FECHA_BASE_COL].dt.year
    df["mes"] = df[FECHA_BASE_COL].dt.month
    df["mes_nombre"] = df[FECHA_BASE_COL].dt.strftime("%Y-%m")
    df["fecha_dia"] = df[FECHA_BASE_COL].dt.strftime("%Y-%m-%d")
    df["semana_iso"] = df[FECHA_BASE_COL].dt.strftime("%G-S%V")
    df["dia_semana"] = df[FECHA_BASE_COL].dt.day_name()

    # Problemas
    df["flag_problema_materiales"] = df["problema_materiales"].apply(tiene_problema)
    df["flag_problema_pagos"] = df["problema_pagos"].apply(tiene_problema)
    df["flag_problema_destaque"] = df["problema_destaque"].apply(tiene_problema)
    df["flag_problema_ssoma"] = df["problema_ssoma"].apply(tiene_problema)

    problem_cols = [
        "flag_problema_materiales",
        "flag_problema_pagos",
        "flag_problema_destaque",
        "flag_problema_ssoma",
    ]

    df["total_problemas"] = df[problem_cols].sum(axis=1, skipna=True)
    df["tiene_algun_problema"] = (df["total_problemas"] > 0).astype(int)

    # Documentación
    df["status_doc_lista"] = df["status_documentacion"].apply(lista_documentos)

    doc_map = {
        "CONTROL DE ASISTENCIA": "doc_control_de_asistencia",
        "FICHAS TECNICAS": "doc_fichas_tecnicas",
        "FICHAS TÉCNICAS": "doc_fichas_tecnicas",
        "PLANES DE TRABAJO": "doc_planes_de_trabajo",
        "HOJAS DE SEGURIDAD": "doc_hojas_de_seguridad",
    }

    for col_name in set(doc_map.values()):
        df[col_name] = 0

    for doc_name, col_name in doc_map.items():
        df[col_name] = df.apply(
            lambda r: 1 if doc_name in r["status_doc_lista"] else r[col_name],
            axis=1
        )

    # Días
    df["dias_hasta_siguiente"] = (
        df["fecha_siguiente_visita"] - df[FECHA_BASE_COL]
    ).dt.days

    if df[FECHA_BASE_COL].notna().any():
        fecha_ref = df[FECHA_BASE_COL].max()
        df["dias_desde_visita"] = (fecha_ref - df[FECHA_BASE_COL]).dt.days
    else:
        df["dias_desde_visita"] = np.nan

    return df


# ─────────────────────────────────────────────────────────────
# 4. TABLAS DE ANÁLISIS
# ─────────────────────────────────────────────────────────────

def construir_periodo(data, granularidad):
    data = data.copy()

    if granularidad == "Día":
        data["periodo"] = data[FECHA_BASE_COL].dt.strftime("%Y-%m-%d")
    elif granularidad == "Semana":
        data["periodo"] = data[FECHA_BASE_COL].dt.strftime("%G-S%V")
    elif granularidad == "Mes":
        data["periodo"] = data[FECHA_BASE_COL].dt.strftime("%Y-%m")
    elif granularidad == "Año":
        data["periodo"] = data[FECHA_BASE_COL].dt.strftime("%Y")
    else:
        data["periodo"] = data[FECHA_BASE_COL].dt.strftime("%Y-%m-%d")

    return data


def tabla_a_visitas_por_motivo(data):
    return (
        data["motivo_visita_norm"]
        .fillna("SIN DATO")
        .value_counts()
        .reset_index()
        .rename(columns={"motivo_visita_norm": "motivo_visita", "count": "visitas"})
    )


def tabla_b_top_clientes(data):
    resumen = (
        data.groupby("cliente", dropna=False)
        .agg(
            visitas=(FECHA_BASE_COL, "count"),
            unidades=("unidad", "nunique"),
            responsables=("responsable_norm", "nunique"),
            puntuacion_prom=("puntuacion_cliente", "mean"),
            operarios_prom=("cantidad_operarios_num", "mean"),
            visitas_con_problemas=("tiene_algun_problema", "sum"),
            total_problemas=("total_problemas", "sum"),
            prob_materiales=("flag_problema_materiales", "sum"),
            prob_pagos=("flag_problema_pagos", "sum"),
            prob_destaque=("flag_problema_destaque", "sum"),
            prob_ssoma=("flag_problema_ssoma", "sum"),
        )
        .reset_index()
        .sort_values("visitas", ascending=False)
    )

    resumen["puntuacion_prom"] = resumen["puntuacion_prom"].round(2)
    resumen["operarios_prom"] = resumen["operarios_prom"].round(1)
    resumen["% visitas con problemas"] = (
        resumen["visitas_con_problemas"] / resumen["visitas"] * 100
    ).round(1)

    return resumen


def tabla_c_responsables(data):
    resumen = (
        data.groupby("responsable_norm", dropna=False)
        .agg(
            visitas=(FECHA_BASE_COL, "count"),
            clientes_atendidos=("cliente", "nunique"),
            unidades_visitadas=("unidad", "nunique"),
            puntuacion_prom=("puntuacion_cliente", "mean"),
            visitas_con_problemas=("tiene_algun_problema", "sum"),
            total_problemas_detectados=("total_problemas", "sum"),
            emergencias=("motivo_visita_norm", lambda x: (x == "POR EMERGENCIA").sum()),
        )
        .reset_index()
        .sort_values("visitas", ascending=False)
    )

    resumen["puntuacion_prom"] = resumen["puntuacion_prom"].round(2)
    resumen["% visitas con problemas"] = (
        resumen["visitas_con_problemas"] / resumen["visitas"] * 100
    ).round(1)

    return resumen


def tabla_d_evolucion(data, granularidad="Mes"):
    data = construir_periodo(data, granularidad)

    evolucion = (
        data.groupby("periodo", dropna=False)
        .agg(
            visitas=(FECHA_BASE_COL, "count"),
            clientes=("cliente", "nunique"),
            unidades=("unidad", "nunique"),
            responsables=("responsable_norm", "nunique"),
            emergencias=("motivo_visita_norm", lambda x: (x == "POR EMERGENCIA").sum()),
            visitas_con_problemas=("tiene_algun_problema", "sum"),
            total_problemas=("total_problemas", "sum"),
            puntuacion_prom=("puntuacion_cliente", "mean"),
        )
        .reset_index()
        .sort_values("periodo")
    )

    evolucion["puntuacion_prom"] = evolucion["puntuacion_prom"].round(2)
    return evolucion


def tabla_e_problemas(data):
    total_visitas = len(data)

    registros = []
    for flag, nombre in [
        ("flag_problema_materiales", "Materiales / Uniformes"),
        ("flag_problema_pagos", "Pagos"),
        ("flag_problema_destaque", "Destaque de personal"),
        ("flag_problema_ssoma", "SSOMA / SIG"),
    ]:
        total_con_resp = data[flag].notna().sum()
        con_problema = data[flag].sum()
        pct_resp = con_problema / total_con_resp * 100 if total_con_resp else 0
        pct_visitas = con_problema / total_visitas * 100 if total_visitas else 0

        registros.append({
            "tipo_problema": nombre,
            "problemas": int(con_problema),
            "registros_con_respuesta": int(total_con_resp),
            "% sobre registros con respuesta": round(pct_resp, 1),
            "% sobre total visitas": round(pct_visitas, 1),
        })

    return pd.DataFrame(registros).sort_values("problemas", ascending=False)


def tabla_f_puntuacion(data):
    dist = (
        data["puntuacion_cliente"]
        .dropna()
        .value_counts()
        .sort_index()
        .reset_index()
        .rename(columns={"puntuacion_cliente": "puntuacion", "count": "cantidad"})
    )

    total = dist["cantidad"].sum() if not dist.empty else 0
    dist["%"] = (dist["cantidad"] / total * 100).round(1) if total else 0
    return dist


def tabla_g_documentacion(data):
    total = len(data)

    registros = []
    for col, nombre in [
        ("doc_control_de_asistencia", "Control de Asistencia"),
        ("doc_fichas_tecnicas", "Fichas Técnicas"),
        ("doc_planes_de_trabajo", "Planes de Trabajo"),
        ("doc_hojas_de_seguridad", "Hojas de Seguridad"),
    ]:
        cnt = data[col].sum()
        pct = cnt / total * 100 if total else 0
        registros.append({
            "documento": nombre,
            "frecuencia": int(cnt),
            "% sobre visitas": round(pct, 1),
        })

    return pd.DataFrame(registros).sort_values("frecuencia", ascending=False)


def tabla_h_sedes_problemas(data):
    resumen = (
        data.groupby("unidad", dropna=False)
        .agg(
            visitas=(FECHA_BASE_COL, "count"),
            clientes=("cliente", "nunique"),
            total_problemas=("total_problemas", "sum"),
            visitas_con_problemas=("tiene_algun_problema", "sum"),
            promedio_problemas=("total_problemas", "mean"),
            puntuacion_prom=("puntuacion_cliente", "mean"),
        )
        .reset_index()
        .sort_values("total_problemas", ascending=False)
    )

    resumen["promedio_problemas"] = resumen["promedio_problemas"].round(2)
    resumen["puntuacion_prom"] = resumen["puntuacion_prom"].round(2)
    resumen["% visitas con problemas"] = (
        resumen["visitas_con_problemas"] / resumen["visitas"] * 100
    ).round(1)

    return resumen


def tabla_i_clientes_peor_puntuacion(data, min_visitas=5):
    clientes_punt = (
        data.groupby("cliente", dropna=False)
        .agg(
            visitas=(FECHA_BASE_COL, "count"),
            unidades=("unidad", "nunique"),
            puntuacion_prom=("puntuacion_cliente", "mean"),
            total_problemas=("total_problemas", "sum"),
            visitas_con_problemas=("tiene_algun_problema", "sum"),
        )
        .reset_index()
    )

    clientes_punt = clientes_punt[
        (clientes_punt["visitas"] >= min_visitas)
        & clientes_punt["puntuacion_prom"].notna()
    ].copy()

    clientes_punt["puntuacion_prom"] = clientes_punt["puntuacion_prom"].round(2)
    clientes_punt["% visitas con problemas"] = (
        clientes_punt["visitas_con_problemas"] / clientes_punt["visitas"] * 100
    ).round(1)

    return clientes_punt.sort_values(["puntuacion_prom", "total_problemas"], ascending=[True, False])


def tabla_j_dias_siguiente_visita(data):
    dias_sig = (
        data.groupby("motivo_visita_norm", dropna=False)["dias_hasta_siguiente"]
        .agg(["count", "mean", "median", "min", "max"])
        .reset_index()
        .rename(columns={
            "motivo_visita_norm": "motivo_visita",
            "count": "registros",
            "mean": "dias_prom",
            "median": "dias_mediana",
            "min": "dias_min",
            "max": "dias_max",
        })
        .sort_values("dias_prom")
    )

    for col in ["dias_prom", "dias_mediana"]:
        dias_sig[col] = dias_sig[col].round(1)

    return dias_sig


def tabla_k_matriz_cliente_unidad_fecha(data, granularidad="Día", min_visitas=1):
    data = construir_periodo(data, granularidad)

    matriz = (
        data.groupby(["cliente", "unidad", "periodo"], dropna=False)
        .size()
        .reset_index(name="visitas")
        .pivot_table(
            index=["cliente", "unidad"],
            columns="periodo",
            values="visitas",
            aggfunc="sum",
            fill_value=0,
        )
    )

    if matriz.empty:
        return matriz

    matriz["TOTAL_VISITAS"] = matriz.sum(axis=1)
    matriz = matriz[matriz["TOTAL_VISITAS"] >= min_visitas]
    matriz = matriz.sort_values("TOTAL_VISITAS", ascending=False)

    cols = ["TOTAL_VISITAS"] + [c for c in matriz.columns if c != "TOTAL_VISITAS"]
    matriz = matriz[cols]

    return matriz


def resumen_cliente_unidad(data):
    resumen = (
        data.groupby(["cliente", "unidad"], dropna=False)
        .agg(
            visitas=(FECHA_BASE_COL, "count"),
            primera_visita=(FECHA_BASE_COL, "min"),
            ultima_visita=(FECHA_BASE_COL, "max"),
            responsables=("responsable_norm", "nunique"),
            responsables_lista=("responsable_norm", lambda x: ", ".join(sorted(set([i for i in x.dropna()])))),
            puntuacion_prom=("puntuacion_cliente", "mean"),
            operarios_prom=("cantidad_operarios_num", "mean"),
            visitas_con_problemas=("tiene_algun_problema", "sum"),
            total_problemas=("total_problemas", "sum"),
            prob_materiales=("flag_problema_materiales", "sum"),
            prob_pagos=("flag_problema_pagos", "sum"),
            prob_destaque=("flag_problema_destaque", "sum"),
            prob_ssoma=("flag_problema_ssoma", "sum"),
            dias_prom_siguiente_visita=("dias_hasta_siguiente", "mean"),
        )
        .reset_index()
        .sort_values(["visitas", "total_problemas"], ascending=False)
    )

    resumen["puntuacion_prom"] = resumen["puntuacion_prom"].round(2)
    resumen["operarios_prom"] = resumen["operarios_prom"].round(1)
    resumen["dias_prom_siguiente_visita"] = resumen["dias_prom_siguiente_visita"].round(1)
    resumen["primera_visita"] = resumen["primera_visita"].dt.strftime("%Y-%m-%d")
    resumen["ultima_visita"] = resumen["ultima_visita"].dt.strftime("%Y-%m-%d")
    resumen["% visitas con problemas"] = (
        resumen["visitas_con_problemas"] / resumen["visitas"] * 100
    ).round(1)

    return resumen


# ─────────────────────────────────────────────────────────────
# 5. SIDEBAR: CARGA Y FILTROS
# ─────────────────────────────────────────────────────────────

st.sidebar.header("1) Cargar datos")

modo_carga = st.sidebar.radio(
    "Origen de datos",
    ["Subir Excel/CSV", "Google Sheets con st.secrets"],
    index=0
)

df_raw = None

if modo_carga == "Subir Excel/CSV":
    uploaded_file = st.sidebar.file_uploader(
        "Sube el archivo de respuestas",
        type=["xlsx", "xls", "csv"]
    )

    if uploaded_file is not None:
        df_raw = leer_archivo_subido(uploaded_file)
    else:
        st.info("Sube tu archivo Excel/CSV para iniciar el análisis.")
        st.stop()

else:
    sheet_name = st.sidebar.text_input(
        "Nombre del Google Sheet",
        value="CHECK LIST SUPERVISIÓN (Respuestas)"
    )
    worksheet_name = st.sidebar.text_input(
        "Nombre de pestaña opcional",
        value=""
    )

    if st.sidebar.button("Cargar Google Sheet"):
        try:
            df_raw = leer_google_sheet(
                sheet_name=sheet_name,
                worksheet_name=worksheet_name if worksheet_name.strip() else None
            )
        except Exception as e:
            st.error(f"No se pudo cargar Google Sheets: {e}")
            st.stop()
    else:
        st.info("Configura `st.secrets` y presiona Cargar Google Sheet.")
        st.stop()


df = preparar_datos(df_raw)

if df.empty:
    st.error("El archivo no contiene registros válidos después de la limpieza.")
    st.stop()


st.sidebar.header("2) Filtros globales")

fecha_min = df[FECHA_BASE_COL].min()
fecha_max = df[FECHA_BASE_COL].max()

if pd.isna(fecha_min) or pd.isna(fecha_max):
    st.error("No se encontraron fechas validas en la columna marca temporal.")
    st.stop()

rango_fecha = st.sidebar.date_input(
    "Rango de fecha (marca temporal)",
    value=(fecha_min.date(), fecha_max.date()),
    min_value=fecha_min.date(),
    max_value=fecha_max.date()
)

if isinstance(rango_fecha, tuple) and len(rango_fecha) == 2:
    fecha_inicio, fecha_fin = rango_fecha
elif isinstance(rango_fecha, date):
    # Permite filtrar un único día cuando el control devuelve una sola fecha.
    fecha_inicio, fecha_fin = rango_fecha, rango_fecha
elif isinstance(rango_fecha, (list, tuple)) and len(rango_fecha) == 1:
    fecha_inicio, fecha_fin = rango_fecha[0], rango_fecha[0]
else:
    fecha_inicio, fecha_fin = fecha_min.date(), fecha_max.date()

responsables_sel = st.sidebar.multiselect(
    "Responsables",
    options=sorted(df["responsable_norm"].dropna().unique())
)

clientes_sel = st.sidebar.multiselect(
    "Clientes",
    options=sorted(df["cliente"].dropna().unique())
)

unidades_sel = st.sidebar.multiselect(
    "Unidades / sedes",
    options=sorted(df["unidad"].dropna().unique())
)

motivos_sel = st.sidebar.multiselect(
    "Motivos de visita",
    options=sorted(df["motivo_visita_norm"].dropna().unique())
)

seguimiento_sel = st.sidebar.multiselect(
    "Seguimiento por",
    options=sorted(df["seguimiento_por_norm"].dropna().unique())
)

tipo_problema = st.sidebar.selectbox(
    "Filtro por problemas",
    [
        "Todos",
        "Solo visitas con problemas",
        "Solo visitas sin problemas",
        "Problemas de materiales / uniformes",
        "Problemas de pagos",
        "Problemas de destaque",
        "Problemas SSOMA / SIG",
    ]
)

granularidad = st.sidebar.selectbox(
    "Granularidad para evolución y matriz",
    ["Día", "Semana", "Mes", "Año"],
    index=2
)

min_visitas_matriz = st.sidebar.number_input(
    "Mínimo de visitas para matriz Cliente-Unidad",
    min_value=1,
    max_value=100,
    value=1,
    step=1
)

min_visitas_peor_punt = st.sidebar.number_input(
    "Mínimo de visitas para tabla I",
    min_value=1,
    max_value=100,
    value=5,
    step=1
)

top_n = st.sidebar.slider(
    "Top N para tablas principales",
    min_value=5,
    max_value=100,
    value=15,
    step=5
)

modo_compacto = st.sidebar.toggle("Vista rápida compacta", value=True)


# Aplicar filtros
data_base = df.copy()

if responsables_sel:
    data_base = data_base[data_base["responsable_norm"].isin(responsables_sel)]

if clientes_sel:
    data_base = data_base[data_base["cliente"].isin(clientes_sel)]

if unidades_sel:
    data_base = data_base[data_base["unidad"].isin(unidades_sel)]

if motivos_sel:
    data_base = data_base[data_base["motivo_visita_norm"].isin(motivos_sel)]

if seguimiento_sel:
    data_base = data_base[data_base["seguimiento_por_norm"].isin(seguimiento_sel)]

if tipo_problema == "Solo visitas con problemas":
    data_base = data_base[data_base["tiene_algun_problema"] == 1]
elif tipo_problema == "Solo visitas sin problemas":
    data_base = data_base[data_base["tiene_algun_problema"] == 0]
elif tipo_problema == "Problemas de materiales / uniformes":
    data_base = data_base[data_base["flag_problema_materiales"] == 1]
elif tipo_problema == "Problemas de pagos":
    data_base = data_base[data_base["flag_problema_pagos"] == 1]
elif tipo_problema == "Problemas de destaque":
    data_base = data_base[data_base["flag_problema_destaque"] == 1]
elif tipo_problema == "Problemas SSOMA / SIG":
    data_base = data_base[data_base["flag_problema_ssoma"] == 1]

fecha_inicio_dt = pd.to_datetime(fecha_inicio)
# Incluir todo el día final (23:59:59.999999) para que un rango de un solo día funcione.
fecha_fin_dt = pd.to_datetime(fecha_fin) + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1)

data = data_base[
    (data_base[FECHA_BASE_COL] >= fecha_inicio_dt)
    & (data_base[FECHA_BASE_COL] <= fecha_fin_dt)
].copy()

if data.empty:
    st.warning("No hay datos con los filtros seleccionados.")
    st.stop()


# ─────────────────────────────────────────────────────────────
# 6. GENERACIÓN DE TABLAS
# ─────────────────────────────────────────────────────────────

tabla_a = tabla_a_visitas_por_motivo(data)
tabla_b = tabla_b_top_clientes(data)
tabla_c = tabla_c_responsables(data)
tabla_d = tabla_d_evolucion(data, granularidad)
tabla_e = tabla_e_problemas(data)
tabla_f = tabla_f_puntuacion(data)
tabla_g = tabla_g_documentacion(data)
tabla_h = tabla_h_sedes_problemas(data)
tabla_i = tabla_i_clientes_peor_puntuacion(data, min_visitas=min_visitas_peor_punt)
tabla_j = tabla_j_dias_siguiente_visita(data)
tabla_k = tabla_k_matriz_cliente_unidad_fecha(
    data,
    granularidad=granularidad,
    min_visitas=min_visitas_matriz
)
tabla_cu = resumen_cliente_unidad(data)

# Resumen ejecutivo compacto
tabla_supervisor_visitas = tabla_c[[
    "responsable_norm",
    "visitas",
    "unidades_visitadas",
    "clientes_atendidos",
    "% visitas con problemas",
]].copy()

ratio_sede_cliente = (
    data.groupby(["responsable_norm", "unidad", "cliente"], dropna=False)
    .size()
    .reset_index(name="visitas")
)
if not ratio_sede_cliente.empty:
    total_por_supervisor = ratio_sede_cliente.groupby("responsable_norm")["visitas"].transform("sum")
    ratio_sede_cliente["ratio_visitas_%"] = (ratio_sede_cliente["visitas"] / total_por_supervisor * 100).round(1)
    ratio_sede_cliente = ratio_sede_cliente.sort_values(
        ["responsable_norm", "visitas", "ratio_visitas_%"],
        ascending=[True, False, False]
    )

promedio_sedes_por_supervisor = tabla_c["unidades_visitadas"].mean() if not tabla_c.empty else 0

unidades_periodo = set(data["unidad"].dropna().unique())
unidades_universo = set(data_base["unidad"].dropna().unique())
sedes_no_visitadas = pd.DataFrame({
    "unidad_no_visitada": sorted(unidades_universo - unidades_periodo)
})


# ─────────────────────────────────────────────────────────────
# 7. KPIS
# ─────────────────────────────────────────────────────────────

total_visitas = len(data)
clientes_unicos = data["cliente"].nunique()
unidades_unicas = data["unidad"].nunique()
responsables_unicos = data["responsable_norm"].nunique()
punt_prom = data["puntuacion_cliente"].mean()
visitas_con_problemas = data["tiene_algun_problema"].sum()
pct_problemas = visitas_con_problemas / total_visitas * 100 if total_visitas else 0
total_problemas = data["total_problemas"].sum()
emergencias = (data["motivo_visita_norm"] == "POR EMERGENCIA").sum()

c1, c2, c3, c4 = st.columns(4)
c1.metric("Visitas", f"{total_visitas:,.0f}")
c2.metric("Clientes", f"{clientes_unicos:,.0f}")
c3.metric("Unidades / sedes", f"{unidades_unicas:,.0f}")
c4.metric("Responsables", f"{responsables_unicos:,.0f}")

c5, c6, c7, c8 = st.columns(4)
c5.metric("Puntuación promedio", f"{punt_prom:.2f}" if pd.notna(punt_prom) else "S/D")
c6.metric("Visitas con problemas", f"{visitas_con_problemas:,.0f}", f"{pct_problemas:.1f}%")
c7.metric("Problemas detectados", f"{total_problemas:,.0f}")
c8.metric("Emergencias", f"{emergencias:,.0f}")

c9, c10 = st.columns(2)
c9.metric("Promedio de sedes por supervisor", f"{promedio_sedes_por_supervisor:.1f}")
c10.metric("Sedes no visitadas (periodo)", f"{len(sedes_no_visitadas):,.0f}")


# ─────────────────────────────────────────────────────────────
# 8. DESCARGA GENERAL
# ─────────────────────────────────────────────────────────────

tabs_export = {
    "Datos_Filtrados": data,
    "Resumen_Supervisores": tabla_supervisor_visitas,
    "Ratio_Sede_Cliente": ratio_sede_cliente,
    "Sedes_No_Visitadas": sedes_no_visitadas,
    "K_Matriz_Cliente_Unidad": tabla_k,
    "Resumen_Cliente_Unidad": tabla_cu,
    "A_Visitas_por_Motivo": tabla_a,
    "B_Top_Clientes": tabla_b,
    "C_Responsables": tabla_c,
    "D_Evolucion": tabla_d,
    "E_Problemas": tabla_e,
    "F_Puntuacion": tabla_f,
    "G_Documentacion": tabla_g,
    "H_Sedes_Problemas": tabla_h,
    "I_Peor_Puntuacion": tabla_i,
    "J_Dias_Siguiente_Visita": tabla_j,
}

excel_bytes = descargar_excel(tabs_export)

st.download_button(
    "⬇️ Descargar Excel con análisis filtrado",
    data=excel_bytes,
    file_name=f"analisis_supervision_filtrado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.divider()


# ─────────────────────────────────────────────────────────────
# 9. TABS DE VISUALIZACIÓN
# ─────────────────────────────────────────────────────────────

if modo_compacto:
    tab_resumen, tab_detalle = st.tabs(["📌 Resumen rápido", "Detalle"])
else:
    tab_resumen, tab_detalle, tab_k, tab_a, tab_b, tab_c, tab_d, tab_e, tab_f, tab_g, tab_h, tab_i, tab_j = st.tabs([
        "📌 Resumen rápido",
        "Detalle",
        "K) Cliente-Unidad x Fecha",
        "A) Motivo",
        "B) Clientes",
        "C) Responsables",
        "D) Evolución",
        "E) Problemas",
        "F) Puntuación",
        "G) Documentación",
        "H) Sedes/Unidades",
        "I) Peor puntuación",
        "J) Siguiente visita",
    ])

with tab_resumen:
    st.subheader("Vista ejecutiva rápida")

    r1c1, r1c2 = st.columns(2)
    with r1c1:
        st.plotly_chart(
            px.bar(
                tabla_supervisor_visitas.head(top_n).sort_values("visitas"),
                x="visitas",
                y="responsable_norm",
                orientation="h",
                text="visitas",
                title=f"Visitas por supervisor (Top {top_n})",
                hover_data=["unidades_visitadas", "clientes_atendidos", "% visitas con problemas"],
            ),
            use_container_width=True,
        )
    with r1c2:
        st.plotly_chart(
            px.bar(
                tabla_a.head(top_n),
                x="motivo_visita",
                y="visitas",
                text="visitas",
                title="Motivos de visita",
            ),
            use_container_width=True,
        )

    r2c1, r2c2 = st.columns(2)
    with r2c1:
        top_ratio = ratio_sede_cliente.head(top_n).copy()
        if not top_ratio.empty:
            top_ratio["sede_cliente"] = top_ratio["unidad"].fillna("SIN DATO") + " | " + top_ratio["cliente"].fillna("SIN DATO")
            st.plotly_chart(
                px.bar(
                    top_ratio.sort_values("ratio_visitas_%"),
                    x="ratio_visitas_%",
                    y="sede_cliente",
                    color="responsable_norm",
                    orientation="h",
                    title=f"Ratio de visitas por sede + cliente (Top {top_n})",
                    hover_data=["visitas"],
                ),
                use_container_width=True,
            )
    with r2c2:
        st.plotly_chart(
            px.line(
                tabla_d,
                x="periodo",
                y="visitas",
                markers=True,
                title=f"Tendencia de visitas por {granularidad.lower()}",
            ),
            use_container_width=True,
        )

    st.subheader("Tablas clave")
    st.dataframe(tabla_supervisor_visitas.head(top_n), use_container_width=True, height=320)
    st.dataframe(ratio_sede_cliente.head(top_n), use_container_width=True, height=320)
    st.dataframe(sedes_no_visitadas, use_container_width=True, height=240)

if not modo_compacto:
    with tab_k:
        st.subheader("K) Cliente + Unidad por columnas de Fecha")
        st.dataframe(tabla_k, use_container_width=True, height=560)
        st.subheader("Resumen Cliente + Unidad")
        st.dataframe(tabla_cu, use_container_width=True, height=520)

    with tab_a:
        st.subheader("A) Visitas por motivo")
        st.plotly_chart(px.bar(tabla_a, x="motivo_visita", y="visitas", text="visitas"), use_container_width=True)
        st.dataframe(tabla_a, use_container_width=True, height=520)

    with tab_b:
        st.subheader("B) Top clientes")
        st.dataframe(tabla_b.head(top_n), use_container_width=True, height=520)

    with tab_c:
        st.subheader("C) Responsables")
        st.dataframe(tabla_c.head(top_n), use_container_width=True, height=520)

    with tab_d:
        st.subheader(f"D) Evolución por {granularidad.lower()}")
        st.dataframe(tabla_d, use_container_width=True, height=520)

    with tab_e:
        st.subheader("E) Problemas")
        st.dataframe(tabla_e, use_container_width=True, height=420)

    with tab_f:
        st.subheader("F) Puntuación")
        st.dataframe(tabla_f, use_container_width=True, height=420)

    with tab_g:
        st.subheader("G) Documentación")
        st.dataframe(tabla_g, use_container_width=True, height=420)

    with tab_h:
        st.subheader("H) Sedes / unidades")
        st.dataframe(tabla_h.head(top_n), use_container_width=True, height=520)

    with tab_i:
        st.subheader(f"I) Peor puntuación (mín. {min_visitas_peor_punt} visitas)")
        st.dataframe(tabla_i.head(top_n), use_container_width=True, height=520)

    with tab_j:
        st.subheader("J) Días promedio hasta siguiente visita por motivo")
        st.dataframe(tabla_j, use_container_width=True, height=520)

with tab_detalle:
    st.subheader("Detalle filtrado")

    detalle_cols = [
        "marca_temporal",
        "fecha_visita",
        "responsable",
        "responsable_norm",
        "email_responsable",
        "cliente",
        "cliente_raw",
        "unidad",
        "sede",
        "cantidad_operarios_num",
        "hora_llegada",
        "motivo_visita",
        "colaborador_entrevistado",
        "status_documentacion",
        "problema_materiales",
        "flag_problema_materiales",
        "problema_pagos",
        "flag_problema_pagos",
        "problema_destaque",
        "flag_problema_destaque",
        "problema_ssoma",
        "flag_problema_ssoma",
        "total_problemas",
        "tiene_algun_problema",
        "puntuacion_cliente",
        "fecha_siguiente_visita",
        "dias_hasta_siguiente",
        "nombre_contacto_cliente",
        "cargo_contacto_cliente",
        "telefono_contacto_cliente",
        "seguimiento_por",
        "oportunidades_mejora",
        "evidencias",
    ]

    detalle_cols = [c for c in detalle_cols if c in data.columns]
    detalle = data[detalle_cols].sort_values("marca_temporal", ascending=False)

    st.dataframe(detalle, use_container_width=True, height=650)
