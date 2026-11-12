"""
Parqueadero Puente Madera · Dashboard de experiencia de servicio
=================================================================

Aplicación Streamlit para analizar la encuesta de satisfacción del
parqueadero Puente Madera de la universidad.

Ejecutar:
    streamlit run app.py

Fuente de datos (por orden de prioridad):
    1. Archivo .xlsx cargado desde la barra lateral.
    2. Ruta indicada en la variable de entorno PARQUEADERO_DATA_PATH.
    3. ``datos.xlsx`` en la misma carpeta que este archivo.

Estructura del código:
    1. Configuración y sistema de diseño (paleta, tipografía, CSS).
    2. Carga y mapeo de columnas (los encabezados reales son largos e
       irregulares; se identifican por palabras clave normalizadas).
    3. Limpieza y cálculos (funciones puras, sin dependencias de UI).
    4. Componentes visuales reutilizables (tarjetas HTML y gráficos Plotly).
    5. Secciones del dashboard.
    6. Punto de entrada ``main()``.
"""

from __future__ import annotations

import html
import inspect
import io
import math
import os
import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# =============================================================================
# 1. CONFIGURACIÓN Y SISTEMA DE DISEÑO
# =============================================================================

APP_TITLE = "Parqueadero Puente Madera"
DEFAULT_DATA_FILE = "datos.xlsx"
DATA_PATH_ENV_VAR = "PARQUEADERO_DATA_PATH"


class C:
    """Paleta semántica. Cada color tiene una función fija en toda la app."""

    bg = "#F5F7FB"
    surface = "#FFFFFF"
    navy = "#1E3A5F"          # marca / dato principal
    blue = "#315B85"          # dato secundario
    accent = "#4F8CFF"        # interacción / sección activa
    text = "#172B4D"
    muted = "#718096"
    border = "#E5EAF1"
    grid = "#EEF2F7"
    positive = "#16A37B"      # favorable / promotores / nunca hay filas
    warning = "#E9A23B"       # pasivos / atención
    negative = "#D95C5C"      # desfavorable / detractores / problemas
    neutral = "#C5CFDC"       # neutral (3) / categorías no destacadas
    soft_blue = "#AFC3DB"     # categorías secundarias


# Escala Likert 1–5 (satisfacción): rojo → gris → verde.
LIKERT_COLORS = {1: "#D95C5C", 2: "#EDA5A0", 3: "#C5CFDC", 4: "#8CD3BC", 5: "#16A37B"}
# Escala de frecuencia de congestión (5 = siempre hay filas): verde → rojo.
CONGESTION_COLORS = {1: "#16A37B", 2: "#8CD3BC", 3: "#C5CFDC", 4: "#EDA5A0", 5: "#D95C5C"}
CONGESTION_LABELS = {1: "1 · Nunca", 2: "2", 3: "3", 4: "4", 5: "5 · Siempre"}

FONT_STACK = "Inter, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif"
PLOTLY_CONFIG = {"displayModeBar": False, "responsive": True}

# Umbrales de tamaño de muestra para advertencias.
N_VERY_SMALL = 10
N_SMALL = 30
N_MIN_SEGMENT = 5  # por debajo de esto un segmento se marca como no concluyente

# Iconos SVG (trazo, estilo lineal) para tarjetas y navegación.
ICONS = {
    "users": '<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/>',
    "car": '<path d="M19 17h2c.6 0 1-.4 1-1v-3c0-.9-.7-1.7-1.5-1.9L18 10l-2.7-3.6A2 2 0 0 0 13.7 6H8.3a2 2 0 0 0-1.6.8L4 10l-2.5 1.1C.7 11.3 0 12.1 0 13v3c0 .6.4 1 1 1h2"/><circle cx="7" cy="17" r="2"/><circle cx="17" cy="17" r="2"/>',
    "star": '<polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/>',
    "megaphone": '<path d="m3 11 18-5v12L3 14v-3z"/><path d="M11.6 16.8a3 3 0 1 1-5.8-1.6"/>',
    "alert": '<path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3Z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>',
    "message": '<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>',
    "clock": '<circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>',
    "trend_up": '<polyline points="22 7 13.5 15.5 8.5 10.5 2 17"/><polyline points="16 7 22 7 22 13"/>',
    "trend_down": '<polyline points="22 17 13.5 8.5 8.5 13.5 2 7"/><polyline points="16 17 22 17 22 11"/>',
    "search": '<circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>',
    "inbox": '<polyline points="22 12 16 12 14 15 10 15 8 12 2 12"/><path d="M5.45 5.11 2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.45-6.89A2 2 0 0 0 16.76 4H7.24a2 2 0 0 0-1.79 1.11z"/>',
    "id": '<rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="9" cy="10" r="2"/><path d="M15 8h2"/><path d="M15 12h2"/><path d="M7 16h10"/>',
    "parking": '<rect x="3" y="3" width="18" height="18" rx="4"/><path d="M9 17V7h4a3 3 0 0 1 0 6H9"/>',
}


def icon(name: str, size: int = 16, color: str = "currentColor") -> str:
    """Devuelve un icono SVG en línea."""
    return (
        f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" '
        f'stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
        f"{ICONS.get(name, '')}</svg>"
    )


# Secciones de navegación: clave → (etiqueta, título de página, descripción)
SECTIONS: dict[str, tuple[str, str, str]] = {
    "overview": ("Resumen ejecutivo", "Resumen ejecutivo",
                 "Panorama general del servicio a partir de las respuestas de la encuesta."),
    "profile": ("Perfil", "Perfil de los encuestados",
                "Quiénes son las personas que usan el parqueadero."),
    "usage": ("Uso del parqueadero", "Uso del parqueadero",
              "Patrones de utilización y diferencias entre perfiles de usuario."),
    "satisfaction": ("Satisfacción", "Satisfacción del servicio",
                     "Evaluación de nueve dimensiones en escala de 1 a 5."),
    "improvements": ("Inconvenientes y mejoras", "Inconvenientes y oportunidades de mejora",
                     "Problemas reportados y propuestas abiertas de los usuarios."),
    "nps": ("Recomendación · NPS", "Recomendación del servicio",
            "Probabilidad de recomendar el parqueadero (escala 0–10) y Net Promoter Score."),
    "explorer": ("Explorador de datos", "Explorador de datos",
                 "Consulta, filtra y descarga los registros de la encuesta."),
    "method": ("Metodología", "Metodología y calidad de datos",
               "Definiciones, denominadores y tratamiento de valores faltantes."),
}


# =============================================================================
# 2. MAPEO DE COLUMNAS
# =============================================================================


def normalize_text(value: object) -> str:
    """Minúsculas, sin tildes, sin espacios duros ni signos repetidos."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    text = str(value).replace("\xa0", " ")
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.lower()
    return re.sub(r"\s+", " ", text).strip()


@dataclass(frozen=True)
class ColumnSpec:
    key: str
    label: str
    keywords: tuple[str, ...] = ()
    exact: str | None = None
    required: bool = False
    kind: str = "category"  # category | numeric | likert | nps | text | datetime | flag | id


# Dimensiones de satisfacción (escala 1 = insatisfecho … 5 = satisfecho).
SATISFACTION_KEYS = [
    "sat_availability", "sat_entry_time", "sat_floor", "sat_signage", "sat_cleanliness",
    "sat_layout", "sat_staff", "sat_info", "sat_resolution",
]

COLUMN_SPECS: list[ColumnSpec] = [
    ColumnSpec("id", "ID", exact="id", kind="id"),
    ColumnSpec("start_time", "Hora de inicio", ("hora de inicio",), kind="datetime"),
    ColumnSpec("end_time", "Hora de finalización", ("hora de finalizacion",), kind="datetime"),
    ColumnSpec("last_modified", "Hora de la última modificación", ("ultima modificacion",), kind="datetime"),
    ColumnSpec("used", "Ha usado el parqueadero", ("has utilizado", "parqueadero"), required=True, kind="flag"),
    ColumnSpec("consent", "Acepta participar", ("aceptas participar",), kind="flag"),
    ColumnSpec("role", "Relación con la universidad", ("relacion con la universidad",)),
    ColumnSpec("age", "Edad", exact="edad", kind="numeric"),
    ColumnSpec("vehicle", "Tipo de vehículo", ("tipo de vehiculo",)),
    ColumnSpec("frequency", "Frecuencia de uso", ("con que frecuencia utilizas",)),
    ColumnSpec("schedule", "Horario de uso", ("en que horario",)),
    ColumnSpec("tenure", "Antigüedad como usuario", ("hace cuanto tiempo",)),
    ColumnSpec("sat_availability", "Disponibilidad de espacios", ("disponibilidad de espacios",), required=True, kind="likert"),
    ColumnSpec("sat_entry_time", "Tiempo de ingreso", ("tiempo que tardas en ingresar",), required=True, kind="likert"),
    ColumnSpec("congestion", "Frecuencia de congestión o filas", ("congestion", "filas"), kind="likert"),
    ColumnSpec("sat_floor", "Estado del piso", ("estado del piso",), required=True, kind="likert"),
    ColumnSpec("sat_signage", "Señalización", ("senalizacion",), required=True, kind="likert"),
    ColumnSpec("sat_cleanliness", "Limpieza", ("limpieza",), required=True, kind="likert"),
    ColumnSpec("sat_layout", "Tamaño y distribución", ("tamano y distribucion",), required=True, kind="likert"),
    ColumnSpec("sat_staff", "Atención del personal", ("atencion del personal",), required=True, kind="likert"),
    ColumnSpec("sat_info", "Claridad de la información", ("clara es la informacion",), required=True, kind="likert"),
    ColumnSpec("sat_resolution", "Rapidez de solución", ("rapidez para solucionar",), required=True, kind="likert"),
    ColumnSpec("had_issue", "Tuvo inconvenientes", ("has tenido algun inconveniente",), kind="flag"),
    ColumnSpec("issue_text", "Principal inconveniente", ("principal inconveniente",), kind="text"),
    ColumnSpec("recommend", "Probabilidad de recomendar (0–10)", ("probable", "recomiend"), required=True, kind="nps"),
    ColumnSpec("improve_text", "Aspecto a mejorar", ("deberia mejorar",), kind="text"),
]
SPEC_BY_KEY = {spec.key: spec for spec in COLUMN_SPECS}
LABELS = {spec.key: spec.label for spec in COLUMN_SPECS}

# Orden lógico de categorías (las no previstas se agregan al final).
CATEGORY_ORDER = {
    "frequency": ["Todos los días", "2-3 veces por semana", "1-2 veces por semana", "Ocasionalmente"],
    "schedule": ["Mañana", "Mediodia", "Mediodía", "Tarde", "Noche"],
    "tenure": ["Menos de un semestre", "1-4 semestres", "Más de 4 semestres"],
    "role": ["Estudiante", "Profesor", "Administrativo"],
    "vehicle": ["Carro", "Moto", "Bicicleta", "Vehiculo institucional"],
    "had_issue": ["Si", "No"],
    "had_issue_label": ["Si", "No"],
    "age_group": ["≤ 20 años", "21–22 años", "≥ 23 años"],
}


def map_columns(columns: list[str]) -> tuple[dict[str, str], list[str]]:
    """Asocia cada clave interna con el encabezado real del Excel.

    Devuelve (mapeo clave→columna real, columnas reales sin asociar).
    """
    normalized = {col: normalize_text(col) for col in columns}
    mapping: dict[str, str] = {}
    used_cols: set[str] = set()
    for spec in COLUMN_SPECS:
        for col, norm in normalized.items():
            if col in used_cols:
                continue
            if spec.exact is not None:
                hit = norm == spec.exact
            else:
                hit = all(kw in norm for kw in spec.keywords)
            if hit:
                mapping[spec.key] = col
                used_cols.add(col)
                break
    unmapped = [c for c in columns if c not in used_cols]
    return mapping, unmapped


# =============================================================================
# 3. CARGA, LIMPIEZA Y CÁLCULOS (sin dependencias de interfaz)
# =============================================================================


@dataclass
class Dataset:
    raw: pd.DataFrame                       # datos originales, sin modificar
    data: pd.DataFrame                      # columnas internas estandarizadas
    mapping: dict[str, str]                 # clave interna → encabezado original
    unmapped: list[str]
    missing_required: list[str]
    source_name: str
    notes: list[str] = field(default_factory=list)


@st.cache_data(show_spinner=False)
def load_data(file_bytes: bytes) -> pd.DataFrame:
    """Lee el Excel (primera hoja) desde bytes. Cacheado por contenido."""
    return pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl")


def _yes_no(series: pd.Series) -> pd.Series:
    """Convierte Sí/No a booleano conservando NaN (vacío ≠ 'No')."""
    norm = series.map(normalize_text)
    out = pd.Series(pd.NA, index=series.index, dtype="boolean")
    out[norm.isin(["si", "s", "yes"])] = True
    out[norm.isin(["no", "n"])] = False
    return out


def _clean_category(series: pd.Series) -> pd.Series:
    """Quita espacios sobrantes sin alterar el texto de la respuesta."""
    def fix(v):
        if pd.isna(v):
            return np.nan
        text = re.sub(r"\s+", " ", str(v).replace("\xa0", " ")).strip()
        return text if text else np.nan
    return series.map(fix).astype(object)


def is_substantive(text: object) -> bool:
    """Un comentario es válido si contiene al menos una letra o dígito."""
    return bool(re.search(r"[0-9a-záéíóúñü]", normalize_text(text)))


def age_group(age: float) -> object:
    if pd.isna(age):
        return np.nan
    if age <= 20:
        return "≤ 20 años"
    if age <= 22:
        return "21–22 años"
    return "≥ 23 años"


def clean_data(raw: pd.DataFrame, source_name: str) -> Dataset:
    """Construye el Dataset de trabajo (la parte costosa está cacheada)."""
    raw = raw.copy()
    raw.columns = [str(c) for c in raw.columns]
    data, mapping, unmapped, missing_required, notes = _build_working_frame(raw)
    return Dataset(raw=raw, data=data, mapping=mapping, unmapped=unmapped,
                   missing_required=missing_required, source_name=source_name, notes=notes)


@st.cache_data(show_spinner=False)
def _build_working_frame(raw: pd.DataFrame):
    """DataFrame con nombres internos + metadatos del mapeo.

    Las respuestas originales no se modifican: solo se normalizan tipos
    (texto → número/booleano) y espacios en blanco. Devuelve solo tipos
    básicos para que la caché de Streamlit pueda serializarlos.
    """
    mapping, unmapped = map_columns(list(raw.columns))
    missing_required = [LABELS[s.key] for s in COLUMN_SPECS if s.required and s.key not in mapping]
    data = pd.DataFrame(index=raw.index)
    notes: list[str] = []

    for spec in COLUMN_SPECS:
        if spec.key not in mapping:
            continue
        col = raw[mapping[spec.key]]
        if spec.kind in ("likert", "nps", "numeric", "id"):
            values = pd.to_numeric(col, errors="coerce")
            if spec.kind == "likert":
                invalid = values.notna() & ~values.isin([1, 2, 3, 4, 5])
                if invalid.any():
                    notes.append(f"{spec.label}: {int(invalid.sum())} valor(es) fuera de 1–5 excluido(s).")
                values = values.where(~invalid)
            if spec.kind == "nps":
                invalid = values.notna() & ~values.isin(range(0, 11))
                if invalid.any():
                    notes.append(f"{spec.label}: {int(invalid.sum())} valor(es) fuera de 0–10 excluido(s).")
                values = values.where(~invalid)
            data[spec.key] = values
        elif spec.kind == "flag":
            data[spec.key] = _yes_no(col)
            data[spec.key + "_label"] = _clean_category(col)
        elif spec.kind == "datetime":
            data[spec.key] = pd.to_datetime(col, errors="coerce")
        else:  # category / text
            data[spec.key] = _clean_category(col)

    if "id" not in data:
        data["id"] = np.arange(1, len(data) + 1)
    if "age" in data:
        data["age_group"] = data["age"].map(age_group).astype(object)

    # Índice individual de satisfacción: media de las dimensiones respondidas.
    sat_cols = [k for k in SATISFACTION_KEYS if k in data]
    if sat_cols:
        data["sat_index"] = data[sat_cols].mean(axis=1, skipna=True)

    # Clasificación NPS por respondente.
    if "recommend" in data:
        data["nps_group"] = data["recommend"].map(nps_group).astype(object)

    # Comentarios: validez y temas (múltiples temas posibles por comentario).
    for key, rules in (("improve_text", IMPROVEMENT_THEMES), ("issue_text", ISSUE_THEMES)):
        if key in data:
            data[key + "_valid"] = data[key].map(is_substantive)
            data[key + "_themes"] = data[key].map(lambda t, r=rules: classify_themes(t, r) if is_substantive(t) else [])

    return data, mapping, unmapped, missing_required, notes


def nps_group(score: float) -> object:
    if pd.isna(score):
        return np.nan
    if score <= 6:
        return "Detractor"
    if score <= 8:
        return "Pasivo"
    return "Promotor"


# ---- Agrupación temática de comentarios -------------------------------------
# Reglas transparentes basadas en palabras clave presentes en los propios
# comentarios. Se comparan sobre texto normalizado (sin tildes, minúsculas).
# Un comentario puede pertenecer a varios temas. Los comentarios originales
# siempre se conservan y pueden consultarse.

IMPROVEMENT_THEMES: list[tuple[str, tuple[str, ...]]] = [
    ("Distancia, acceso y movilidad peatonal", (
        "lejos", "distancia", "caminata", "caminar", "acceso", "segundo piso", "segunda planta",
        "parte superior", "parqueaderos de arriba", "parqueaderos de la parte de arriba",
        "conexion", "llegar mas rapido", "peatones")),
    ("Espacios y distribución", (
        "espacio", "distribucion", "tamano", "pequenos", "organizacion", "abiertos", "muy pocos")),
    ("Estado del piso y de la vía", ("el piso", "del piso", "pavimentar", "barro", "inunda")),
    ("Tiempos, filas y flujo vehicular", ("tiempo", "demora", "congestion", "via aparte", "torniquete")),
    ("Atención del personal", ("personal",)),
    ("Señalización", ("senalizacion",)),
]
ISSUE_THEMES: list[tuple[str, tuple[str, ...]]] = [
    ("Falla en la lectura del GoPass", ("gopass", "go pass")),
    ("Ocupación / disponibilidad de espacios", ("mucho carro", "parquear", "lleno", "espacio")),
]
NON_PROPOSAL = {"nose", "no se", "ninguno", "ninguna", "nada", "no", "n/a", "na"}
THEME_OTHER = "Otros / sin clasificar"
THEME_NONE = "Sin propuesta concreta"


def classify_themes(text: object, rules: list[tuple[str, tuple[str, ...]]]) -> list[str]:
    norm = normalize_text(text)
    if norm.strip(" .") in NON_PROPOSAL:
        return [THEME_NONE]
    themes = [name for name, keywords in rules if any(kw in norm for kw in keywords)]
    return themes or [THEME_OTHER]


# ---- Métricas ----------------------------------------------------------------


def is_true(series: pd.Series) -> pd.Series:
    """Máscara segura para booleanos anulables (NA → False)."""
    return (series == True).fillna(False).astype(bool)  # noqa: E712


def is_false(series: pd.Series) -> pd.Series:
    return (series == False).fillna(False).astype(bool)  # noqa: E712


def safe_div(num: float, den: float) -> float:
    return float(num) / float(den) if den else float("nan")


def likert_stats(series: pd.Series) -> dict:
    """Estadísticos de una pregunta 1–5 usando solo respuestas válidas."""
    valid = series.dropna()
    n = int(len(valid))
    counts = valid.value_counts().reindex([1, 2, 3, 4, 5], fill_value=0).astype(int)
    return {
        "n": n,
        "mean": float(valid.mean()) if n else float("nan"),
        "sd": float(valid.std(ddof=1)) if n > 1 else float("nan"),
        "median": float(valid.median()) if n else float("nan"),
        "counts": counts,
        "low": int(counts[[1, 2]].sum()),   # 1–2
        "mid": int(counts[3]),              # 3
        "high": int(counts[[4, 5]].sum()),  # 4–5
        "pct_low": safe_div(counts[[1, 2]].sum(), n) * 100,
        "pct_mid": safe_div(counts[3], n) * 100,
        "pct_high": safe_div(counts[[4, 5]].sum(), n) * 100,
    }


def satisfaction_table(users: pd.DataFrame) -> pd.DataFrame:
    """Una fila por dimensión de satisfacción (excluye congestión)."""
    rows = []
    for key in SATISFACTION_KEYS:
        if key not in users:
            continue
        s = likert_stats(users[key])
        rows.append({
            "key": key, "Dimensión": LABELS[key], "n": s["n"], "Media": s["mean"], "DE": s["sd"],
            "Mediana": s["median"], "% favorable": s["pct_high"], "% neutral": s["pct_mid"],
            "% desfavorable": s["pct_low"], **{f"c{i}": int(s["counts"][i]) for i in range(1, 6)},
        })
    return pd.DataFrame(rows)


def nps_stats(series: pd.Series) -> dict:
    """NPS = % promotores (9–10) − % detractores (0–6), con IC 95 % aproximado."""
    valid = series.dropna()
    n = int(len(valid))
    prom = int((valid >= 9).sum())
    det = int((valid <= 6).sum())
    pas = n - prom - det
    p, d = safe_div(prom, n), safe_div(det, n)
    nps = (p - d) * 100 if n else float("nan")
    # Varianza por respondente de una variable que vale +1/0/−1: p + d − (p − d)².
    if n > 1:
        se = math.sqrt(max(p + d - (p - d) ** 2, 0) / n) * 100
        ci = 1.96 * se
    else:
        ci = float("nan")
    return {"n": n, "promoters": prom, "passives": pas, "detractors": det,
            "pct_prom": p * 100 if n else float("nan"), "pct_pas": safe_div(pas, n) * 100,
            "pct_det": d * 100 if n else float("nan"), "nps": nps, "ci95": ci,
            "mean": float(valid.mean()) if n else float("nan"),
            "dist": valid.value_counts().reindex(range(0, 11), fill_value=0).astype(int)}


def theme_counts(frame: pd.DataFrame, key: str) -> pd.DataFrame:
    """Número de comentarios válidos por tema (multi-etiqueta)."""
    col_valid, col_themes = key + "_valid", key + "_themes"
    if col_valid not in frame:
        return pd.DataFrame(columns=["Tema", "Comentarios", "%"])
    valid = frame[frame[col_valid].fillna(False).astype(bool)]
    n = len(valid)
    exploded = valid[col_themes].explode().dropna()
    counts = exploded.value_counts()
    out = counts.rename_axis("Tema").reset_index(name="Comentarios")
    out["%"] = out["Comentarios"] / n * 100 if n else np.nan
    return out


@dataclass
class Metrics:
    n_total: int
    n_users: int
    n_nonusers: int
    sat: pd.DataFrame
    sat_index: float
    pct_fav_global: float
    n_sat_respondents: int
    congestion: dict | None
    nps: dict | None
    n_issue_valid: int
    n_issue_yes: int


def prepare_metrics(filtered: pd.DataFrame) -> Metrics:
    users = filtered[is_true(filtered["used"])]
    nonusers = filtered[is_false(filtered["used"])]
    sat = satisfaction_table(users)
    sat_index = float(sat.loc[sat["n"] > 0, "Media"].mean()) if not sat.empty and (sat["n"] > 0).any() else float("nan")
    # % favorable global: respuestas 4–5 sobre todas las respuestas válidas de las 9 dimensiones.
    total_answers = int(sat[[f"c{i}" for i in range(1, 6)]].to_numpy().sum()) if not sat.empty else 0
    fav_answers = int(sat[["c4", "c5"]].to_numpy().sum()) if not sat.empty else 0
    sat_cols = [k for k in SATISFACTION_KEYS if k in users]
    n_sat_resp = int(users[sat_cols].notna().any(axis=1).sum()) if sat_cols else 0
    congestion = likert_stats(users["congestion"]) if "congestion" in users else None
    nps = nps_stats(users["recommend"]) if "recommend" in users else None
    if "had_issue" in users:
        n_issue_valid = int(users["had_issue"].notna().sum())
        n_issue_yes = int(is_true(users["had_issue"]).sum())
    else:
        n_issue_valid = n_issue_yes = 0
    return Metrics(
        n_total=len(filtered), n_users=len(users), n_nonusers=len(nonusers), sat=sat,
        sat_index=sat_index, pct_fav_global=safe_div(fav_answers, total_answers) * 100,
        n_sat_respondents=n_sat_resp, congestion=congestion, nps=nps,
        n_issue_valid=n_issue_valid, n_issue_yes=n_issue_yes,
    )


def segment_satisfaction(users: pd.DataFrame, by: str) -> pd.DataFrame:
    """Índice de satisfacción, recomendación y tamaño por segmento."""
    if by not in users or "sat_index" not in users:
        return pd.DataFrame()
    frame = users.dropna(subset=[by])
    rows = []
    for value, group in frame.groupby(by, sort=False, observed=True):
        idx = group["sat_index"].dropna()
        rec = group["recommend"].dropna() if "recommend" in group else pd.Series(dtype=float)
        rows.append({"Segmento": value, "n": int(len(group)),
                     "Índice satisfacción": float(idx.mean()) if len(idx) else np.nan,
                     "DE": float(idx.std(ddof=1)) if len(idx) > 1 else np.nan,
                     "Recomendación media (0–10)": float(rec.mean()) if len(rec) else np.nan})
    out = pd.DataFrame(rows)
    if out.empty:
        return out
    return order_frame(out, "Segmento", by)


def ordered_categories(values, key: str) -> list:
    """Ordena categorías según el orden lógico conocido; el resto por frecuencia."""
    present = list(dict.fromkeys([v for v in values if pd.notna(v)]))
    order = CATEGORY_ORDER.get(key, [])
    known = [o for o in order if o in present]
    rest = [v for v in present if v not in known]
    return known + rest


def order_frame(frame: pd.DataFrame, col: str, key: str) -> pd.DataFrame:
    if key not in CATEGORY_ORDER:
        return frame.sort_values("n", ascending=False).reset_index(drop=True)
    order = ordered_categories(frame[col].tolist(), key)
    return frame.set_index(col).loc[order].reset_index()


# =============================================================================
# 4. FILTROS
# =============================================================================

FILTER_CATEGORIES = [
    ("vehicle", "Tipo de vehículo"),
    ("role", "Relación con la universidad"),
    ("frequency", "Frecuencia de uso"),
    ("schedule", "Horario"),
    ("tenure", "Antigüedad"),
]
ISSUE_FILTER_OPTIONS = ["Todos", "Reportaron inconvenientes", "Sin inconvenientes"]


def filter_defaults(data: pd.DataFrame) -> dict:
    defaults = {f"f_{key}": [] for key, _ in FILTER_CATEGORIES}
    if "age" in data and data["age"].notna().any():
        defaults["f_age"] = (int(data["age"].min()), int(data["age"].max()))
    defaults["f_issue"] = ISSUE_FILTER_OPTIONS[0]
    return defaults


def init_filter_state(data: pd.DataFrame, signature: str) -> None:
    """Inicializa (o reinicia si cambia el archivo) el estado de los filtros."""
    if st.session_state.get("_data_signature") != signature:
        for k, v in filter_defaults(data).items():
            st.session_state[k] = v
        st.session_state["_data_signature"] = signature
    else:
        for k, v in filter_defaults(data).items():
            st.session_state.setdefault(k, v)


def reset_filters(data: pd.DataFrame) -> None:
    for k, v in filter_defaults(data).items():
        st.session_state[k] = v


def active_filters(data: pd.DataFrame) -> dict[str, str]:
    """Filtros activos → descripción legible."""
    active = {}
    for key, label in FILTER_CATEGORIES:
        sel = st.session_state.get(f"f_{key}", [])
        if sel:
            active[key] = f"{label}: {', '.join(map(str, sel))}"
    default_age = filter_defaults(data).get("f_age")
    age = st.session_state.get("f_age")
    if default_age and age and tuple(age) != tuple(default_age):
        active["age"] = f"Edad: {age[0]}–{age[1]}"
    issue = st.session_state.get("f_issue", ISSUE_FILTER_OPTIONS[0])
    if issue != ISSUE_FILTER_OPTIONS[0]:
        active["issue"] = issue
    return active


def apply_filters(data: pd.DataFrame, active: dict) -> pd.DataFrame:
    """Aplica los filtros de la barra lateral.

    Las variables de perfil solo existen para usuarios del parqueadero; por ello,
    cuando hay al menos un filtro de perfil activo, el subconjunto contiene
    únicamente usuarios que cumplen las condiciones.
    """
    mask = pd.Series(True, index=data.index)
    for key, _ in FILTER_CATEGORIES:
        sel = st.session_state.get(f"f_{key}", [])
        if sel and key in data:
            mask &= data[key].isin(sel)
    if "age" in active and "age" in data:
        lo, hi = st.session_state["f_age"]
        mask &= data["age"].between(lo, hi)
    issue = st.session_state.get("f_issue", ISSUE_FILTER_OPTIONS[0])
    if "had_issue" in data:
        if issue == ISSUE_FILTER_OPTIONS[1]:
            mask &= is_true(data["had_issue"])
        elif issue == ISSUE_FILTER_OPTIONS[2]:
            mask &= is_false(data["had_issue"])
    return data[mask.fillna(False).astype(bool)]


# =============================================================================
# 5. COMPONENTES VISUALES
# =============================================================================


def _width_kwargs(fn) -> dict:
    """Ancho completo compatible con versiones antiguas y nuevas de Streamlit."""
    try:
        params = inspect.signature(fn).parameters
    except (TypeError, ValueError):
        return {}
    if "width" in params and "use_container_width" not in params:
        return {"width": "stretch"}
    if "use_container_width" in params:
        return {"use_container_width": True}
    return {}


def fmt_num(value: float, decimals: int = 1) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "—"
    return f"{value:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_pct(value: float, decimals: int = 0) -> str:
    return "—" if value is None or (isinstance(value, float) and math.isnan(value)) else f"{fmt_num(value, decimals)} %"


def fmt_signed(value: float, decimals: int = 0) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "—"
    sign = "+" if value > 0 else ("−" if value < 0 else "")
    return f"{sign}{fmt_num(abs(value), decimals)}"


def esc(text: object) -> str:
    return html.escape(str(text))


def inject_css() -> None:
    st.markdown(
        f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, .stApp, .stMarkdown, .stText, button, input, textarea, select,
[data-testid="stWidgetLabel"], [data-testid="stSidebar"] {{
    font-family: {FONT_STACK};
}}
.stApp {{ background: {C.bg}; color: {C.text}; }}
header[data-testid="stHeader"] {{ background: transparent; }}
#MainMenu, footer, [data-testid="stStatusWidget"] {{ visibility: hidden; }}
.block-container, [data-testid="stMainBlockContainer"] {{
    padding-top: 2.2rem; padding-bottom: 3rem; max-width: 1360px;
}}

/* ---------- Sidebar ---------- */
section[data-testid="stSidebar"] {{
    background: {C.surface}; border-right: 1px solid {C.border};
}}
section[data-testid="stSidebar"] .block-container,
section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] {{ padding-top: 1.2rem; }}
.brand {{ display:flex; align-items:center; gap:.7rem; padding:.2rem .2rem 1.1rem .2rem;
          border-bottom:1px solid {C.border}; margin-bottom:1rem; }}
.brand-logo {{ width:36px; height:36px; border-radius:10px; background:{C.navy};
               display:flex; align-items:center; justify-content:center; flex-shrink:0; }}
.brand-name {{ font-weight:700; font-size:.95rem; color:{C.text}; line-height:1.15; }}
.brand-sub {{ font-size:.74rem; color:{C.muted}; }}
.side-label {{ font-size:.68rem; font-weight:600; letter-spacing:.08em; text-transform:uppercase;
               color:{C.muted}; margin:1.1rem 0 .45rem .15rem; }}

/* Navegación: el radio nativo se presenta como menú lateral. */
section[data-testid="stSidebar"] div[role="radiogroup"] {{ gap: 2px; }}
section[data-testid="stSidebar"] div[role="radiogroup"] > label {{
    width: 100%; padding: .48rem .7rem; border-radius: 8px; margin: 0;
    transition: background .15s ease; cursor: pointer; position: relative;
}}
section[data-testid="stSidebar"] div[role="radiogroup"] > label > div:first-child {{ display: none; }}
section[data-testid="stSidebar"] div[role="radiogroup"] > label p {{
    font-size: .88rem; font-weight: 500; color: {C.muted};
}}
section[data-testid="stSidebar"] div[role="radiogroup"] > label:hover {{ background: {C.bg}; }}
section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) {{
    background: #EEF4FF; box-shadow: inset 3px 0 0 {C.accent};
}}
section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) p {{
    color: {C.navy}; font-weight: 650;
}}
section[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p {{
    font-size: .8rem; font-weight: 600; color: {C.text};
}}
section[data-testid="stSidebar"] .stButton button {{ width: 100%; }}
span[data-baseweb="tag"] {{ background: #E8EFFA !important; color: {C.navy} !important; }}
span[data-baseweb="tag"] span {{ color: {C.navy} !important; }}

/* ---------- Encabezado de página ---------- */
.page-head {{ display:flex; justify-content:space-between; align-items:flex-end; gap:1.5rem;
              flex-wrap:wrap; margin-bottom:1.4rem; }}
.eyebrow {{ font-size:.72rem; font-weight:600; letter-spacing:.09em; text-transform:uppercase;
            color:{C.accent}; margin-bottom:.35rem; }}
.page-title {{ font-size:1.9rem; font-weight:750; letter-spacing:-.02em; color:{C.text};
               line-height:1.15; margin:0; }}
.page-sub {{ font-size:.95rem; color:{C.muted}; margin-top:.35rem; max-width:720px; }}
.chips {{ display:flex; gap:.45rem; flex-wrap:wrap; }}
.chip {{ display:inline-flex; align-items:center; gap:.35rem; font-size:.76rem; font-weight:500;
         color:{C.text}; background:{C.surface}; border:1px solid {C.border}; border-radius:999px;
         padding:.3rem .7rem; white-space:nowrap; }}
.chip svg {{ color:{C.muted}; }}
.chip-accent {{ background:#EEF4FF; border-color:#D6E4FF; color:{C.navy}; }}

/* ---------- Tarjetas KPI ---------- */
.kpi-grid {{ display:grid; grid-template-columns:repeat(auto-fit, minmax(190px, 1fr)); gap:14px;
             margin-bottom:18px; }}
.kpi {{ background:{C.surface}; border:1px solid {C.border}; border-radius:14px; padding:16px 18px 15px;
        box-shadow:0 1px 2px rgba(23,43,77,.04); display:flex; flex-direction:column; min-height:138px; }}
.kpi-top {{ display:flex; justify-content:space-between; align-items:center; }}
.kpi-label {{ font-size:.74rem; font-weight:600; color:{C.muted}; letter-spacing:.02em; }}
.kpi-icon {{ width:28px; height:28px; border-radius:8px; display:flex; align-items:center;
             justify-content:center; background:{C.bg}; color:{C.blue}; }}
.kpi-value {{ font-size:2.05rem; font-weight:750; letter-spacing:-.03em; color:{C.text};
              margin-top:.45rem; line-height:1; font-variant-numeric: tabular-nums; }}
.kpi-unit {{ font-size:.95rem; font-weight:500; color:{C.muted}; margin-left:.2rem; letter-spacing:0; }}
.kpi-desc {{ font-size:.78rem; color:{C.muted}; margin-top:.45rem; line-height:1.35; }}
.kpi-foot {{ margin-top:auto; padding-top:.6rem; }}
.badge {{ display:inline-flex; align-items:center; gap:.3rem; font-size:.7rem; font-weight:600;
          border-radius:6px; padding:.16rem .45rem; }}
.b-pos {{ background:#E5F6F0; color:#0F7A5C; }}
.b-neg {{ background:#FBEAEA; color:#B03E3E; }}
.b-warn {{ background:#FDF3E3; color:#9A6412; }}
.b-neu {{ background:{C.bg}; color:{C.muted}; }}
.b-acc {{ background:#EEF4FF; color:{C.navy}; }}

/* ---------- Tarjetas contenedoras ---------- */
div[data-testid="stVerticalBlockBorderWrapper"]:has(> div > div > div > div > .card-head),
[class*="st-key-card_"] {{
    background:{C.surface}; border:1px solid {C.border} !important; border-radius:14px !important;
    box-shadow:0 1px 2px rgba(23,43,77,.04);
}}
[class*="st-key-card_"] {{ padding: 18px 20px 10px 20px; }}
.card-head {{ display:flex; justify-content:space-between; align-items:flex-start; gap:1rem; }}
.card-title {{ font-size:.98rem; font-weight:650; color:{C.text}; margin:0; }}
.card-sub {{ font-size:.79rem; color:{C.muted}; margin-top:.15rem; line-height:1.4; }}
.section-gap {{ height: 6px; }}

/* ---------- Hallazgos ---------- */
.finding {{ display:flex; gap:.75rem; padding:.72rem 0; border-bottom:1px solid {C.grid}; }}
.finding:last-child {{ border-bottom:none; }}
.f-dot {{ width:8px; height:8px; border-radius:50%; margin-top:.42rem; flex-shrink:0; }}
.f-title {{ font-size:.86rem; font-weight:600; color:{C.text}; }}
.f-text {{ font-size:.8rem; color:{C.muted}; margin-top:.1rem; line-height:1.45; }}

/* ---------- Barras HTML / listas ---------- */
.split {{ display:flex; height:10px; border-radius:6px; overflow:hidden; margin:.55rem 0 .7rem; }}
.legend-row {{ display:flex; gap:1.1rem; flex-wrap:wrap; font-size:.78rem; color:{C.muted}; }}
.legend-row b {{ color:{C.text}; font-weight:650; }}
.sw {{ display:inline-block; width:9px; height:9px; border-radius:3px; margin-right:.35rem; }}
.stat-list {{ display:grid; grid-template-columns:1fr 1fr; gap:.2rem 1.2rem; margin-top:.4rem; }}
.stat-item {{ padding:.55rem 0; border-bottom:1px solid {C.grid}; }}
.stat-k {{ font-size:.72rem; color:{C.muted}; font-weight:500; }}
.stat-v {{ font-size:.9rem; color:{C.text}; font-weight:650; margin-top:.1rem; }}
.rank-row {{ display:flex; align-items:center; gap:.7rem; padding:.5rem 0; border-bottom:1px solid {C.grid}; }}
.rank-row:last-child {{ border-bottom:none; }}
.rank-n {{ width:22px; height:22px; border-radius:6px; font-size:.72rem; font-weight:700;
           display:flex; align-items:center; justify-content:center; flex-shrink:0; }}
.rank-name {{ font-size:.84rem; color:{C.text}; font-weight:550; flex:1; }}
.rank-val {{ font-size:.84rem; font-weight:700; color:{C.text}; font-variant-numeric:tabular-nums; }}

/* ---------- Comentarios ---------- */
.comment {{ background:{C.surface}; border:1px solid {C.border}; border-radius:12px; padding:.85rem 1rem;
            margin-bottom:.6rem; }}
.comment-text {{ font-size:.9rem; color:{C.text}; line-height:1.5; }}
.comment-meta {{ display:flex; gap:.4rem; flex-wrap:wrap; margin-top:.55rem; align-items:center; }}
.tag {{ font-size:.68rem; font-weight:600; border-radius:5px; padding:.12rem .42rem;
        background:#EEF4FF; color:{C.navy}; }}
.meta {{ font-size:.7rem; color:{C.muted}; }}

/* ---------- Estados vacíos y notas ---------- */
.empty {{ background:{C.surface}; border:1px dashed #CFD8E3; border-radius:14px; padding:2.2rem 1.5rem;
          text-align:center; }}
.empty-icon {{ width:44px; height:44px; border-radius:12px; background:{C.bg}; margin:0 auto .8rem;
               display:flex; align-items:center; justify-content:center; color:{C.muted}; }}
.empty-title {{ font-weight:650; color:{C.text}; font-size:.98rem; }}
.empty-text {{ color:{C.muted}; font-size:.84rem; margin-top:.3rem; }}
.note {{ display:flex; gap:.6rem; align-items:flex-start; background:#F8FAFD; border:1px solid {C.border};
         border-left:3px solid {C.warning}; border-radius:10px; padding:.65rem .9rem; font-size:.8rem;
         color:{C.text}; margin-bottom:1rem; line-height:1.45; }}
.note.info {{ border-left-color:{C.accent}; }}
.foot {{ font-size:.74rem; color:{C.muted}; margin-top:2rem; padding-top:1rem; border-top:1px solid {C.border}; }}
.method h4 {{ font-size:.95rem; margin:1.1rem 0 .3rem; color:{C.text}; }}
.method p, .method li {{ font-size:.86rem; color:{C.text}; line-height:1.55; }}

[data-testid="stDataFrame"] {{ border:1px solid {C.border}; border-radius:10px; }}
.stTabs [data-baseweb="tab-list"] {{ gap: 4px; }}
.stTabs [data-baseweb="tab"] {{ font-size:.85rem; }}

@media (max-width: 900px) {{
    .page-title {{ font-size:1.5rem; }}
    .kpi-value {{ font-size:1.7rem; }}
    .stat-list {{ grid-template-columns:1fr; }}
}}
</style>
""",
        unsafe_allow_html=True,
    )


def html_block(markup: str) -> None:
    st.markdown(markup, unsafe_allow_html=True)


def page_header(section: str, period: str | None, n_filtered: int, n_total: int, n_active: int) -> None:
    _, title, subtitle = SECTIONS[section]
    chips = []
    if period:
        chips.append(f'<span class="chip">{icon("clock", 13)}{esc(period)}</span>')
    chips.append(f'<span class="chip">{icon("users", 13)}{n_filtered} de {n_total} registros</span>')
    if n_active:
        chips.append(f'<span class="chip chip-accent">{n_active} filtro{"s" if n_active > 1 else ""} activo{"s" if n_active > 1 else ""}</span>')
    html_block(
        f'<div class="page-head"><div><div class="eyebrow">Encuesta de servicio · {esc(APP_TITLE)}</div>'
        f'<h1 class="page-title">{esc(title)}</h1><div class="page-sub">{esc(subtitle)}</div></div>'
        f'<div class="chips">{"".join(chips)}</div></div>'
    )


@dataclass
class Kpi:
    label: str
    value: str
    unit: str = ""
    desc: str = ""
    badge: str = ""
    badge_kind: str = "neu"
    icon_name: str = "users"


def kpi_row(items: list[Kpi]) -> None:
    cards = []
    for k in items:
        badge = f'<div class="kpi-foot"><span class="badge b-{k.badge_kind}">{esc(k.badge)}</span></div>' if k.badge else ""
        unit = f'<span class="kpi-unit">{esc(k.unit)}</span>' if k.unit else ""
        cards.append(
            f'<div class="kpi"><div class="kpi-top"><span class="kpi-label">{esc(k.label)}</span>'
            f'<span class="kpi-icon">{icon(k.icon_name, 15)}</span></div>'
            f'<div class="kpi-value">{esc(k.value)}{unit}</div>'
            f'<div class="kpi-desc">{esc(k.desc)}</div>{badge}</div>'
        )
    html_block(f'<div class="kpi-grid">{"".join(cards)}</div>')


def sample_badge(n: int) -> str:
    if n < N_VERY_SMALL:
        return f'<span class="badge b-neg">n = {n} · muestra muy pequeña</span>'
    if n < N_SMALL:
        return f'<span class="badge b-warn">n = {n} · muestra pequeña</span>'
    return f'<span class="badge b-neu">n = {n}</span>'


_card_counter = {"i": 0}


def card(title: str, subtitle: str = "", n: int | None = None, right_html: str = ""):
    """Contenedor con estilo de tarjeta y encabezado. Uso: ``with card(...):``"""
    _card_counter["i"] += 1
    key = f"card_{_card_counter['i']}"
    try:
        container = st.container(border=True, key=key)
    except TypeError:  # versiones sin parámetro ``key``
        container = st.container(border=True)
    right = right_html or (sample_badge(n) if n is not None else "")
    sub = f'<div class="card-sub">{esc(subtitle)}</div>' if subtitle else ""
    container.markdown(
        f'<div class="card-head"><div><div class="card-title">{esc(title)}</div>{sub}</div>'
        f'<div>{right}</div></div>', unsafe_allow_html=True)
    return container


def empty_state(title: str, text: str, icon_name: str = "inbox") -> None:
    html_block(
        f'<div class="empty"><div class="empty-icon">{icon(icon_name, 20)}</div>'
        f'<div class="empty-title">{esc(title)}</div><div class="empty-text">{esc(text)}</div></div>'
    )


def note(text: str, kind: str = "warn") -> None:
    cls = "note info" if kind == "info" else "note"
    html_block(f'<div class="{cls}"><div>{text}</div></div>')


def show_chart(fig: go.Figure, key: str) -> None:
    st.plotly_chart(fig, config=PLOTLY_CONFIG, key=key, **_width_kwargs(st.plotly_chart))


def show_table(frame: pd.DataFrame, **kwargs) -> None:
    st.dataframe(frame, hide_index=True, **_width_kwargs(st.dataframe), **kwargs)


# ---- Plotly ------------------------------------------------------------------


def style_fig(fig: go.Figure, height: int = 320, legend: bool = False, axes: bool = True) -> go.Figure:
    fig.update_layout(
        height=height,
        margin=dict(l=8, r=16, t=12, b=8),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family=FONT_STACK, size=12, color=C.text),
        separators=",.",
        showlegend=legend,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0,
                    font=dict(size=11, color=C.muted), bgcolor="rgba(0,0,0,0)"),
        hoverlabel=dict(bgcolor="white", bordercolor=C.border, font=dict(family=FONT_STACK, size=12, color=C.text)),
        bargap=0.32,
    )
    if not axes:
        return fig
    fig.update_xaxes(gridcolor=C.grid, zeroline=False, linecolor=C.border, tickfont=dict(color=C.muted, size=11),
                     title_font=dict(color=C.muted, size=11))
    fig.update_yaxes(gridcolor=C.grid, zeroline=False, linecolor=C.border, tickfont=dict(color=C.text, size=12),
                     title_font=dict(color=C.muted, size=11))
    return fig


def hbar_distribution(series: pd.Series, key: str | None = None, highlight_top: bool = True,
                      base_color: str = C.soft_blue, top_color: str = C.navy) -> go.Figure | None:
    """Barras horizontales con conteo y % sobre respuestas válidas."""
    valid = series.dropna()
    if valid.empty:
        return None
    counts = valid.value_counts()
    order = ordered_categories(counts.index.tolist(), key) if key in CATEGORY_ORDER else counts.index.tolist()
    counts = counts.reindex(order)
    n = int(counts.sum())
    pct = counts / n * 100
    max_val = counts.max()
    colors = [top_color if (highlight_top and v == max_val) else base_color for v in counts.values]
    fig = go.Figure(go.Bar(
        x=counts.values, y=[str(i) for i in counts.index], orientation="h", marker=dict(color=colors, cornerradius=4),
        text=[f"<b>{c}</b>  ·  {fmt_num(p, 0)} %" for c, p in zip(counts.values, pct.values)],
        textposition="outside", cliponaxis=False, textfont=dict(size=11, color=C.text),
        customdata=pct.values,
        hovertemplate=f"<b>%{{y}}</b><br>%{{x}} de {n} respuestas válidas<br>%{{customdata:.1f}} %<extra></extra>",
    ))
    fig.update_yaxes(autorange="reversed", showgrid=False)
    fig.update_xaxes(showticklabels=False, showgrid=False, range=[0, max_val * 1.35])
    return style_fig(fig, height=max(150, 46 * len(counts) + 40))


def age_histogram(ages: pd.Series) -> go.Figure | None:
    valid = ages.dropna()
    if valid.empty:
        return None
    counts = valid.astype(int).value_counts().sort_index()
    full = pd.Series(0, index=range(int(valid.min()), int(valid.max()) + 1))
    full.update(counts)
    median = float(valid.median())
    fig = go.Figure(go.Bar(
        x=full.index, y=full.values, marker=dict(color=C.blue, cornerradius=3),
        hovertemplate="<b>%{x} años</b><br>%{y} persona(s)<extra></extra>",
    ))
    fig.add_vline(x=median, line=dict(color=C.accent, width=1.5, dash="dot"))
    fig.add_annotation(x=median, y=1, yref="paper", text=f"Mediana {fmt_num(median, 0)}", showarrow=False,
                       xanchor="left", xshift=6, font=dict(size=11, color=C.accent))
    fig.update_xaxes(title="Edad (años)", dtick=2 if len(full) > 14 else 1, showgrid=False)
    fig.update_yaxes(title="Personas", dtick=2 if full.max() > 6 else 1)
    style_fig(fig, height=300)
    fig.update_layout(bargap=0.15)
    return fig


def satisfaction_means_chart(sat: pd.DataFrame, overall: float) -> go.Figure:
    frame = sat[sat["n"] > 0]
    best = frame["Media"].max()
    worst = frame["Media"].min()
    colors = []
    for m in frame["Media"]:
        if len(frame) > 1 and m == best:
            colors.append(C.positive)
        elif len(frame) > 1 and m == worst:
            colors.append(C.negative)
        else:
            colors.append(C.blue)
    custom = np.stack([frame["n"], frame["DE"].fillna(0), frame["% favorable"], frame["% desfavorable"]], axis=1)
    fig = go.Figure(go.Bar(
        x=frame["Media"], y=frame["Dimensión"], orientation="h", base=0,
        marker=dict(color=colors, cornerradius=4),
        text=[fmt_num(m, 2) for m in frame["Media"]], textposition="outside", cliponaxis=False,
        textfont=dict(size=12, color=C.text),
        customdata=custom,
        hovertemplate=("<b>%{y}</b><br>Media: %{x:.2f} / 5<br>Desv. estándar: %{customdata[1]:.2f}"
                       "<br>Favorable (4–5): %{customdata[2]:.0f} %<br>Desfavorable (1–2): %{customdata[3]:.0f} %"
                       "<br>n = %{customdata[0]}<extra></extra>"),
    ))
    if not math.isnan(overall):
        fig.add_vline(x=overall, line=dict(color=C.muted, width=1, dash="dot"))
        fig.add_annotation(x=overall, y=1.0, yref="paper", yanchor="bottom", text=f"Índice {fmt_num(overall, 2)}",
                           showarrow=False, font=dict(size=11, color=C.muted))
    fig.update_xaxes(range=[1, 5.45], tickvals=[1, 2, 3, 4, 5], title="Media (1 = insatisfecho · 5 = satisfecho)")
    fig.update_yaxes(autorange="reversed", showgrid=False)
    style_fig(fig, height=max(260, 40 * len(frame) + 80))
    fig.update_layout(margin=dict(l=8, r=16, t=30, b=8))
    return fig


def stacked_sentiment_chart(sat: pd.DataFrame) -> go.Figure:
    """Barras 100 % apiladas: desfavorable (1–2), neutral (3), favorable (4–5)."""
    frame = sat[sat["n"] > 0]
    parts = [("Desfavorable (1–2)", "% desfavorable", C.negative, ["c1", "c2"]),
             ("Neutral (3)", "% neutral", C.neutral, ["c3"]),
             ("Favorable (4–5)", "% favorable", C.positive, ["c4", "c5"])]
    fig = go.Figure()
    for name, col, color, count_cols in parts:
        counts = frame[count_cols].sum(axis=1)
        fig.add_trace(go.Bar(
            name=name, x=frame[col], y=frame["Dimensión"], orientation="h", marker=dict(color=color),
            text=[fmt_num(v, 0) + " %" if v >= 9 else "" for v in frame[col]], textposition="inside",
            insidetextanchor="middle", textfont=dict(size=11, color="white" if color != C.neutral else C.text),
            customdata=np.stack([counts, frame["n"]], axis=1),
            hovertemplate=f"<b>%{{y}}</b><br>{name}: %{{x:.0f}} %<br>%{{customdata[0]}} de %{{customdata[1]}} respuestas<extra></extra>",
        ))
    fig.update_layout(barmode="stack")
    fig.update_xaxes(range=[0, 100], ticksuffix=" %", showgrid=False)
    fig.update_yaxes(autorange="reversed", showgrid=False)
    return style_fig(fig, height=max(260, 40 * len(frame) + 80), legend=True)


def likert_heatmap(sat: pd.DataFrame) -> go.Figure:
    frame = sat[sat["n"] > 0]
    counts = frame[[f"c{i}" for i in range(1, 6)]].to_numpy()
    n = frame["n"].to_numpy()[:, None]
    pct = np.divide(counts, n, out=np.zeros_like(counts, dtype=float), where=n > 0) * 100
    x_labels = ["1", "2", "3", "4", "5"]
    vmax = max(60.0, float(pct.max()) if pct.size else 60.0)
    fig = go.Figure(go.Heatmap(
        z=pct, x=x_labels, y=frame["Dimensión"].tolist(), zmin=0, zmax=vmax,
        colorscale=[[0, "#F4F7FB"], [0.35, "#BFD3EA"], [0.7, "#5B84B1"], [1, C.navy]],
        showscale=False, xgap=3, ygap=3, customdata=counts,
        hovertemplate="<b>%{y}</b><br>Calificación %{x}: %{customdata} respuesta(s) · %{z:.0f} %<extra></extra>",
    ))
    for i, dim in enumerate(frame["Dimensión"]):
        for j in range(5):
            fig.add_annotation(x=x_labels[j], y=dim, text=f"{int(counts[i, j])}", showarrow=False,
                               font=dict(size=11, color="white" if pct[i, j] > vmax * 0.55 else C.text))
    fig.update_xaxes(side="top", title="Calificación", showgrid=False)
    fig.update_yaxes(autorange="reversed", showgrid=False)
    return style_fig(fig, height=max(260, 40 * len(frame) + 70))


def scale_bar_chart(stats: dict, colors: dict, labels: dict | None = None) -> go.Figure:
    counts = stats["counts"]
    n = stats["n"]
    x = [labels.get(i, str(i)) if labels else str(i) for i in range(1, 6)]
    pct = [safe_div(c, n) * 100 for c in counts.values]
    fig = go.Figure(go.Bar(
        x=x, y=counts.values, marker=dict(color=[colors[i] for i in range(1, 6)], cornerradius=4),
        text=[f"{c}" for c in counts.values], textposition="outside", cliponaxis=False,
        customdata=pct, hovertemplate="<b>%{x}</b><br>%{y} respuesta(s) · %{customdata:.0f} %<extra></extra>",
    ))
    fig.update_yaxes(showticklabels=False, showgrid=False, range=[0, max(counts.max(), 1) * 1.25])
    fig.update_xaxes(showgrid=False, type="category")
    return style_fig(fig, height=230)


def crosstab_heatmap(frame: pd.DataFrame, row_key: str, col_key: str) -> go.Figure | None:
    sub = frame.dropna(subset=[row_key, col_key])
    if sub.empty:
        return None
    table = pd.crosstab(sub[row_key], sub[col_key])
    table = table.reindex(index=ordered_categories(table.index.tolist(), row_key),
                          columns=ordered_categories(table.columns.tolist(), col_key))
    z = table.to_numpy()
    row_tot = z.sum(axis=1, keepdims=True)
    row_pct = np.divide(z, row_tot, out=np.zeros_like(z, dtype=float), where=row_tot > 0) * 100
    fig = go.Figure(go.Heatmap(
        z=z, x=[str(c) for c in table.columns], y=[str(i) for i in table.index],
        colorscale=[[0, "#F4F7FB"], [0.5, "#8DAACD"], [1, C.navy]], showscale=False, xgap=3, ygap=3,
        customdata=row_pct,
        hovertemplate="<b>%{y}</b> × <b>%{x}</b><br>%{z} persona(s)<br>%{customdata:.0f} % de la fila<extra></extra>",
    ))
    zmax = z.max() if z.size else 1
    for i, r in enumerate(table.index):
        for j, c in enumerate(table.columns):
            fig.add_annotation(x=str(c), y=str(r), text=str(int(z[i, j])), showarrow=False,
                               font=dict(size=12, color="white" if z[i, j] > zmax * 0.55 else C.text))
    fig.update_xaxes(side="top", showgrid=False)
    fig.update_yaxes(autorange="reversed", showgrid=False)
    return style_fig(fig, height=max(200, 52 * len(table.index) + 80))


def segment_chart(seg: pd.DataFrame, overall: float) -> go.Figure:
    colors = [C.blue if n >= N_MIN_SEGMENT else C.soft_blue for n in seg["n"]]
    fig = go.Figure(go.Bar(
        x=seg["Índice satisfacción"], y=[str(s) for s in seg["Segmento"]], orientation="h",
        marker=dict(color=colors, cornerradius=4),
        text=[f"{fmt_num(v, 2)}  ·  n = {n}" for v, n in zip(seg["Índice satisfacción"], seg["n"])],
        textposition="outside", cliponaxis=False, customdata=seg["n"],
        hovertemplate="<b>%{y}</b><br>Índice de satisfacción: %{x:.2f}<br>n = %{customdata}<extra></extra>",
    ))
    if not math.isnan(overall):
        fig.add_vline(x=overall, line=dict(color=C.muted, width=1, dash="dot"))
    fig.update_xaxes(range=[1, 5.6], tickvals=[1, 2, 3, 4, 5], title="Índice de satisfacción (1–5)")
    fig.update_yaxes(autorange="reversed", showgrid=False)
    return style_fig(fig, height=max(170, 50 * len(seg) + 70))


def nps_gauge(nps: float) -> go.Figure:
    fig = go.Figure(go.Indicator(
        mode="gauge+number", value=nps,
        number=dict(font=dict(size=46, family=FONT_STACK, color=C.text), valueformat=".0f"),
        gauge=dict(
            axis=dict(range=[-100, 100], tickvals=[-100, -50, 0, 50, 100], tickfont=dict(size=10, color=C.muted)),
            bar=dict(color=C.navy, thickness=0.28),
            bgcolor="white", borderwidth=0,
            steps=[dict(range=[-100, 0], color="#FBEAEA"), dict(range=[0, 30], color="#FDF3E3"),
                   dict(range=[30, 100], color="#E5F6F0")],
            threshold=dict(line=dict(color=C.text, width=2), thickness=0.8, value=0),
        ),
    ))
    style_fig(fig, height=230, axes=False)
    fig.update_layout(margin=dict(l=24, r=24, t=20, b=0))
    return fig


def nps_distribution_chart(dist: pd.Series) -> go.Figure:
    colors = [C.negative if s <= 6 else (C.warning if s <= 8 else C.positive) for s in dist.index]
    n = int(dist.sum())
    fig = go.Figure(go.Bar(
        x=[str(i) for i in dist.index], y=dist.values, marker=dict(color=colors, cornerradius=4),
        text=[str(v) if v else "" for v in dist.values], textposition="outside", cliponaxis=False,
        customdata=[safe_div(v, n) * 100 for v in dist.values],
        hovertemplate="<b>Puntuación %{x}</b><br>%{y} persona(s) · %{customdata:.0f} %<extra></extra>",
    ))
    fig.update_xaxes(type="category", title="Puntuación (0 = nada probable · 10 = muy probable)", showgrid=False)
    fig.update_yaxes(title="Personas", showgrid=True, range=[0, max(dist.max(), 1) * 1.2], dtick=2 if dist.max() > 6 else 1)
    return style_fig(fig, height=280)


def theme_chart(themes: pd.DataFrame, n_comments: int) -> go.Figure:
    frame = themes.copy()
    special = {THEME_OTHER, THEME_NONE}
    frame["_special"] = frame["Tema"].isin(special)
    frame = frame.sort_values(["_special", "Comentarios"], ascending=[True, False])
    top = frame.loc[~frame["_special"], "Comentarios"].max() if (~frame["_special"]).any() else None
    colors = [C.neutral if s else (C.navy if c == top else C.soft_blue) for s, c in zip(frame["_special"], frame["Comentarios"])]
    fig = go.Figure(go.Bar(
        x=frame["Comentarios"], y=frame["Tema"], orientation="h", marker=dict(color=colors, cornerradius=4),
        text=[f"<b>{c}</b>  ·  {fmt_num(p, 0)} %" for c, p in zip(frame["Comentarios"], frame["%"])],
        textposition="outside", cliponaxis=False,
        hovertemplate=f"<b>%{{y}}</b><br>%{{x}} de {n_comments} comentarios válidos<extra></extra>",
    ))
    fig.update_yaxes(autorange="reversed", showgrid=False)
    fig.update_xaxes(showticklabels=False, showgrid=False, range=[0, frame["Comentarios"].max() * 1.4])
    return style_fig(fig, height=max(180, 44 * len(frame) + 40))


# =============================================================================
# 6. SECCIONES
# =============================================================================


@dataclass
class Context:
    ds: Dataset
    filtered: pd.DataFrame
    users: pd.DataFrame
    metrics: Metrics
    active: dict
    period: str | None


def users_guard(ctx: Context, what: str = "este análisis") -> bool:
    """Muestra un estado vacío si no hay usuarios en el subconjunto filtrado."""
    if ctx.metrics.n_users == 0:
        empty_state("No hay usuarios del parqueadero con estos filtros",
                    f"Amplía o restablece los filtros de la barra lateral para ver {what}.", "search")
        return False
    return True


def filter_note(ctx: Context) -> None:
    if ctx.active:
        note(f"<b>Filtros activos:</b> {esc(' · '.join(ctx.active.values()))}. "
             "Las variables de perfil solo existen para usuarios del parqueadero, por lo que el subconjunto "
             "filtrado excluye a quienes no lo han utilizado.", "info")


def build_findings(ctx: Context) -> list[tuple[str, str, str]]:
    """Hallazgos calculados a partir de los datos filtrados (color, título, texto)."""
    m = ctx.metrics
    findings: list[tuple[str, str, str]] = []
    sat = m.sat[m.sat["n"] > 0]
    if len(sat) >= 2:
        best = sat.loc[sat["Media"].idxmax()]
        worst = sat.loc[sat["Media"].idxmin()]
        findings.append((C.positive, f"Mejor valorada: {best['Dimensión']}",
                         f"Media {fmt_num(best['Media'], 2)} / 5 · {fmt_pct(best['% favorable'])} de respuestas favorables (n = {int(best['n'])})."))
        findings.append((C.negative, f"Mayor oportunidad: {worst['Dimensión']}",
                         f"Media {fmt_num(worst['Media'], 2)} / 5 · {fmt_pct(worst['% desfavorable'])} de respuestas desfavorables (n = {int(worst['n'])})."))
    if m.congestion and m.congestion["n"]:
        c = m.congestion
        findings.append((C.warning, "Congestión al ingresar",
                         f"{fmt_pct(c['pct_high'])} de los usuarios encuentra filas con frecuencia (4–5) y "
                         f"{fmt_pct(c['pct_low'])} rara vez o nunca (1–2); n = {c['n']}."))
    if m.nps and m.nps["n"]:
        p = m.nps
        rel = "superan a" if p["detractors"] > p["promoters"] else ("son menos que" if p["detractors"] < p["promoters"] else "igualan a")
        findings.append((C.negative if p["nps"] < 0 else C.positive, f"NPS de {fmt_signed(p['nps'])}",
                         f"Los detractores ({p['detractors']}) {rel} los promotores ({p['promoters']}); "
                         f"{fmt_pct(p['pct_pas'])} son pasivos (n = {p['n']})."))
    themes = theme_counts(ctx.users, "improve_text")
    themes = themes[~themes["Tema"].isin([THEME_OTHER, THEME_NONE])] if not themes.empty else themes
    n_comments = int(ctx.users.get("improve_text_valid", pd.Series(dtype=bool)).fillna(False).astype(bool).sum())
    if not themes.empty and n_comments:
        top = themes.iloc[0]
        findings.append((C.accent, f"Tema de mejora más mencionado: {top['Tema']}",
                         f"Aparece en {int(top['Comentarios'])} de {n_comments} comentarios abiertos válidos "
                         f"({fmt_pct(top['%'])}). Un comentario puede tener varios temas."))
    return findings


def render_overview(ctx: Context) -> None:
    m = ctx.metrics
    filter_note(ctx)
    nps = m.nps or {}
    kpis = [
        Kpi("Encuestados", str(m.n_total), desc="Respuestas en el subconjunto actual", icon_name="users",
            badge="Total de la encuesta" if not ctx.active else "Subconjunto filtrado", badge_kind="neu"),
        Kpi("Usuarios del parqueadero", str(m.n_users), unit=f"/ {m.n_total}",
            desc="Han usado el parqueadero este semestre", icon_name="car",
            badge=f"{fmt_pct(safe_div(m.n_users, m.n_total) * 100)} de encuestados", badge_kind="acc"),
        Kpi("Índice de satisfacción", fmt_num(m.sat_index, 2), unit="/ 5",
            desc="Media de las 9 dimensiones (excluye congestión)", icon_name="star",
            badge=f"{fmt_pct(m.pct_fav_global)} respuestas favorables",
            badge_kind="pos" if m.pct_fav_global >= 60 else ("warn" if m.pct_fav_global >= 40 else "neg")),
        Kpi("Net Promoter Score", fmt_signed(nps.get("nps", float("nan"))),
            desc=f"Promotores {fmt_pct(nps.get('pct_prom'))} · Detractores {fmt_pct(nps.get('pct_det'))}",
            icon_name="megaphone",
            badge=(f"IC 95 % ± {fmt_num(nps.get('ci95'), 0)} pts" if nps.get("n", 0) > 1 else "Sin datos suficientes"),
            badge_kind="neg" if nps.get("nps", 0) < 0 else "pos"),
        Kpi("Reportaron inconvenientes", fmt_pct(safe_div(m.n_issue_yes, m.n_issue_valid) * 100),
            desc=f"{m.n_issue_yes} de {m.n_issue_valid} usuarios con respuesta", icon_name="alert",
            badge="Durante este semestre", badge_kind="neu"),
    ]
    kpi_row(kpis)

    if m.n_users == 0:
        users_guard(ctx, "los resultados")
        return

    left, right = st.columns([7, 5], gap="medium")
    with left:
        with card("¿Qué tan satisfechos están los usuarios?",
                  "Proporción de respuestas desfavorables, neutrales y favorables por dimensión, ordenadas por % favorable.",
                  n=m.n_sat_respondents):
            sat = m.sat[m.sat["n"] > 0].sort_values("% favorable", ascending=False)
            show_chart(stacked_sentiment_chart(sat), "ov_sentiment")
    with right:
        with card("Hallazgos principales", "Calculados sobre el subconjunto filtrado."):
            findings = build_findings(ctx)
            if findings:
                html_block("".join(
                    f'<div class="finding"><span class="f-dot" style="background:{color}"></span>'
                    f'<div><div class="f-title">{esc(t)}</div><div class="f-text">{esc(txt)}</div></div></div>'
                    for color, t, txt in findings))
            else:
                empty_state("Sin hallazgos", "No hay datos suficientes en este subconjunto.")

    c1, c2 = st.columns([5, 7], gap="medium")
    with c1:
        with card("Participación", "Encuestados que han usado o no el parqueadero este semestre."):
            pu = safe_div(m.n_users, m.n_total) * 100
            pn = safe_div(m.n_nonusers, m.n_total) * 100
            html_block(
                f'<div class="split"><div style="width:{0 if math.isnan(pu) else pu}%;background:{C.navy}"></div>'
                f'<div style="width:{0 if math.isnan(pn) else pn}%;background:{C.neutral}"></div></div>'
                f'<div class="legend-row"><span><span class="sw" style="background:{C.navy}"></span>Usuarios <b>{m.n_users}</b> ({fmt_pct(pu)})</span>'
                f'<span><span class="sw" style="background:{C.neutral}"></span>No usuarios <b>{m.n_nonusers}</b> ({fmt_pct(pn)})</span></div>'
                f'<div class="card-sub" style="margin-top:.8rem">Quienes no han usado el parqueadero finalizan la encuesta tras la '
                f'primera pregunta, por eso no aportan datos de perfil ni de satisfacción.</div>')
    with c2:
        with card("Perfil típico del usuario", "Categoría más frecuente de cada variable entre los usuarios."):
            items = []
            for key in ["role", "vehicle", "frequency", "schedule", "tenure"]:
                if key in ctx.users and ctx.users[key].notna().any():
                    vc = ctx.users[key].value_counts()
                    n_valid = int(vc.sum())
                    items.append((LABELS[key], f"{vc.index[0]} · {fmt_pct(vc.iloc[0] / n_valid * 100)}"))
            if "age" in ctx.users and ctx.users["age"].notna().any():
                a = ctx.users["age"].dropna()
                items.append(("Edad", f"Mediana {fmt_num(a.median(), 0)} años · rango {int(a.min())}–{int(a.max())}"))
            html_block('<div class="stat-list">' + "".join(
                f'<div class="stat-item"><div class="stat-k">{esc(k)}</div><div class="stat-v">{esc(v)}</div></div>'
                for k, v in items) + "</div>")


def render_demographics(ctx: Context) -> None:
    filter_note(ctx)
    note("Las preguntas de perfil solo se aplicaron a quienes han usado el parqueadero. "
         f"Los porcentajes se calculan sobre las respuestas válidas de cada pregunta (usuarios en el subconjunto: {ctx.metrics.n_users}).", "info")
    if not users_guard(ctx, "el perfil"):
        return
    u = ctx.users
    ages = u["age"].dropna() if "age" in u else pd.Series(dtype=float)
    share = lambda key, val: safe_div((u[key] == val).sum(), u[key].notna().sum()) * 100 if key in u else float("nan")  # noqa: E731
    kpi_row([
        Kpi("Edad mediana", fmt_num(ages.median(), 0) if len(ages) else "—", unit="años",
            desc=f"Rango {int(ages.min())}–{int(ages.max())} años" if len(ages) else "Sin datos", icon_name="id"),
        Kpi("Estudiantes", fmt_pct(share("role", "Estudiante")), desc="Proporción de usuarios que son estudiantes", icon_name="users"),
        Kpi("Llegan en carro", fmt_pct(share("vehicle", "Carro")), desc="Vehículo principal de acceso", icon_name="car"),
        Kpi("Usuarios con edad registrada", str(len(ages)), unit=f"/ {ctx.metrics.n_users}", desc="Respuestas válidas de edad", icon_name="id"),
    ])
    c1, c2 = st.columns([6, 6], gap="medium")
    with c1:
        with card("Distribución por edad", "Número de usuarios por edad en años cumplidos.", n=len(ages)):
            fig = age_histogram(ages)
            show_chart(fig, "dem_age") if fig else empty_state("Sin datos de edad", "")
        if len(ages) >= 4:
            q1, q3 = ages.quantile(0.25), ages.quantile(0.75)
            outliers = ages[(ages > q3 + 1.5 * (q3 - q1)) | (ages < q1 - 1.5 * (q3 - q1))]
            if len(outliers):
                note(f"El 50 % central de las edades está entre {fmt_num(q1, 0)} y {fmt_num(q3, 0)} años y hay "
                     f"{len(outliers)} valor(es) atípico(s) ({', '.join(str(int(a)) for a in sorted(outliers))}); "
                     "por eso se reporta la mediana y no la media.", "info")
    with c2:
        for key, title, sub in [("role", "Relación con la universidad", "Vínculo del usuario con la institución."),
                                ("vehicle", "Tipo de vehículo", "Medio principal para acceder al parqueadero.")]:
            if key not in u:
                continue
            with card(title, sub, n=int(u[key].notna().sum())):
                fig = hbar_distribution(u[key], key)
                show_chart(fig, f"dem_{key}") if fig else empty_state("Sin respuestas", "")


def render_usage_analysis(ctx: Context) -> None:
    filter_note(ctx)
    m = ctx.metrics
    u = ctx.users
    freq_share = safe_div((u["frequency"] == "Todos los días").sum(), u["frequency"].notna().sum()) * 100 if "frequency" in u else float("nan")
    morning = safe_div((u["schedule"] == "Mañana").sum(), u["schedule"].notna().sum()) * 100 if "schedule" in u else float("nan")
    veteran = safe_div((u["tenure"] == "Más de 4 semestres").sum(), u["tenure"].notna().sum()) * 100 if "tenure" in u else float("nan")
    kpi_row([
        Kpi("Tasa de uso", fmt_pct(safe_div(m.n_users, m.n_total) * 100), desc=f"{m.n_users} de {m.n_total} encuestados lo han usado", icon_name="car"),
        Kpi("Uso diario", fmt_pct(freq_share), desc="Usuarios que lo usan todos los días", icon_name="clock"),
        Kpi("Franja de la mañana", fmt_pct(morning), desc="Horario principal reportado", icon_name="clock"),
        Kpi("Más de 4 semestres", fmt_pct(veteran), desc="Usuarios con mayor antigüedad", icon_name="id"),
    ])
    if not users_guard(ctx, "los patrones de uso"):
        return
    c1, c2, c3 = st.columns(3, gap="medium")
    for col, (key, title, sub) in zip([c1, c2, c3], [
            ("frequency", "Frecuencia de uso", "¿Con qué frecuencia lo utilizan?"),
            ("schedule", "Horario principal", "Franja en que más lo utilizan."),
            ("tenure", "Antigüedad como usuario", "Tiempo que llevan usando el servicio.")]):
        if key not in u:
            continue
        with col:
            with card(title, sub, n=int(u[key].notna().sum())):
                fig = hbar_distribution(u[key], key)
                show_chart(fig, f"use_{key}") if fig else empty_state("Sin respuestas", "")

    st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
    extra_labels = {"age_group": "Grupo de edad", "had_issue_label": "Tuvo inconvenientes"}
    cross_options = {k: extra_labels.get(k, LABELS.get(k, k))
                     for k in ["frequency", "schedule", "tenure", "vehicle", "role", "age_group", "had_issue_label"] if k in u}

    left, right = st.columns([6, 6], gap="medium")
    with left:
        with card("Cruce de variables", "Número de usuarios en cada combinación. El tooltip muestra el % de la fila."):
            a, b = st.columns(2)
            keys = list(cross_options)
            row_key = a.selectbox("Filas", keys, index=keys.index("frequency") if "frequency" in keys else 0,
                                  format_func=cross_options.get, key="cross_rows")
            col_default = keys.index("schedule") if "schedule" in keys else min(1, len(keys) - 1)
            col_key = b.selectbox("Columnas", keys, index=col_default, format_func=cross_options.get, key="cross_cols")
            if row_key == col_key:
                empty_state("Selecciona dos variables distintas", "El cruce requiere dos variables diferentes.")
            else:
                fig = crosstab_heatmap(u, row_key, col_key)
                show_chart(fig, "use_cross") if fig else empty_state("Sin datos para este cruce", "")
    with right:
        with card("Satisfacción por segmento", "Índice individual = media de las 9 dimensiones respondidas por cada usuario.",
                  n=m.n_users):
            seg_key = st.selectbox("Segmentar por", list(cross_options),
                                   index=list(cross_options).index("tenure") if "tenure" in cross_options else 0,
                                   format_func=cross_options.get, key="seg_key")
            seg = segment_satisfaction(u, seg_key)
            if seg.empty:
                empty_state("Sin datos para segmentar", "")
            else:
                overall = float(u["sat_index"].mean())
                show_chart(segment_chart(seg, overall), "use_segment")
                if (seg["n"] < N_MIN_SEGMENT).any():
                    html_block(f'<div class="card-sub">Barras claras: segmentos con menos de {N_MIN_SEGMENT} personas; '
                               'sus valores no son concluyentes. Las diferencias observadas son descriptivas y no implican causalidad.</div>')
                else:
                    html_block('<div class="card-sub">Las diferencias observadas son descriptivas y no implican causalidad.</div>')


def render_satisfaction(ctx: Context) -> None:
    filter_note(ctx)
    m = ctx.metrics
    if not users_guard(ctx, "la satisfacción"):
        return
    sat = m.sat[m.sat["n"] > 0].copy()
    if sat.empty:
        empty_state("Sin respuestas de satisfacción", "No hay calificaciones válidas en este subconjunto.")
        return
    best = sat.loc[sat["Media"].idxmax()]
    worst = sat.loc[sat["Media"].idxmin()]
    kpi_row([
        Kpi("Índice de satisfacción", fmt_num(m.sat_index, 2), unit="/ 5", desc="Media de las medias de 9 dimensiones", icon_name="star"),
        Kpi("Respuestas favorables", fmt_pct(m.pct_fav_global), desc="Calificaciones 4–5 sobre el total de calificaciones válidas", icon_name="trend_up",
            badge=f"Desfavorables {fmt_pct(100 * safe_div(sat[['c1', 'c2']].to_numpy().sum(), sat[[f'c{i}' for i in range(1, 6)]].to_numpy().sum()))}", badge_kind="neg"),
        Kpi("Mejor valorada", best["Dimensión"], desc=f"Media {fmt_num(best['Media'], 2)} · {fmt_pct(best['% favorable'])} favorable", icon_name="trend_up",
            badge="Fortaleza", badge_kind="pos"),
        Kpi("Peor valorada", worst["Dimensión"], desc=f"Media {fmt_num(worst['Media'], 2)} · {fmt_pct(worst['% desfavorable'])} desfavorable", icon_name="trend_down",
            badge="Prioridad de mejora", badge_kind="neg"),
    ])
    html_block("<style>.kpi-grid .kpi:nth-child(3) .kpi-value, .kpi-grid .kpi:nth-child(4) .kpi-value"
               "{font-size:1.25rem; letter-spacing:-.01em; line-height:1.2;}</style>")

    sort_options = {"Media": "Media", "% favorable": "% favorable", "% desfavorable": "% desfavorable"}
    left, right = st.columns([7, 5], gap="medium")
    with left:
        with card("Media por dimensión", "Escala 1–5. Verde: mejor valorada · rojo: peor valorada · línea punteada: índice general.",
                  n=m.n_sat_respondents):
            order_by = st.radio("Ordenar por", list(sort_options), horizontal=True, key="sat_sort", label_visibility="collapsed")
            ordered = sat.sort_values(order_by, ascending=False)
            show_chart(satisfaction_means_chart(ordered, m.sat_index), "sat_means")
    with right:
        with card("Ranking de dimensiones", "Ordenadas por media; entre paréntesis, % de respuestas favorables."):
            ranked = sat.sort_values("Media", ascending=False).reset_index(drop=True)
            rows = []
            for i, r in ranked.iterrows():
                if i < 3:
                    bg, fg = "#E5F6F0", "#0F7A5C"
                elif i >= len(ranked) - 3:
                    bg, fg = "#FBEAEA", "#B03E3E"
                else:
                    bg, fg = C.bg, C.muted
                rows.append(f'<div class="rank-row"><span class="rank-n" style="background:{bg};color:{fg}">{i + 1}</span>'
                            f'<span class="rank-name">{esc(r["Dimensión"])}</span>'
                            f'<span class="meta">({fmt_pct(r["% favorable"])})</span>'
                            f'<span class="rank-val">{fmt_num(r["Media"], 2)}</span></div>')
            html_block("".join(rows))

    c1, c2 = st.columns([6, 6], gap="medium")
    with c1:
        with card("Favorable, neutral y desfavorable", "Desfavorable = 1–2 · Neutral = 3 · Favorable = 4–5. % sobre respuestas válidas de cada dimensión."):
            show_chart(stacked_sentiment_chart(sat.sort_values("% favorable", ascending=False)), "sat_stack")
    with c2:
        with card("Distribución de calificaciones", "Número de respuestas por calificación; la intensidad del color indica el % de la fila."):
            show_chart(likert_heatmap(sat.sort_values("Media", ascending=False)), "sat_heat")

    if m.congestion and m.congestion["n"]:
        c = m.congestion
        a, b = st.columns([7, 5], gap="medium")
        with a:
            with card("Frecuencia de congestión o filas al ingresar",
                      "Escala de frecuencia (1 = nunca · 5 = siempre). Se analiza aparte: no mide satisfacción y no entra en el índice.",
                      n=c["n"]):
                show_chart(scale_bar_chart(c, CONGESTION_COLORS, CONGESTION_LABELS), "sat_congestion")
        with b:
            with card("Lectura de la congestión", "Aquí un valor alto es negativo."):
                html_block(
                    '<div class="stat-list">'
                    f'<div class="stat-item"><div class="stat-k">Media de frecuencia</div><div class="stat-v">{fmt_num(c["mean"], 2)} / 5</div></div>'
                    f'<div class="stat-item"><div class="stat-k">Mediana</div><div class="stat-v">{fmt_num(c["median"], 0)}</div></div>'
                    f'<div class="stat-item"><div class="stat-k">Frecuente (4–5)</div><div class="stat-v" style="color:{C.negative}">{fmt_pct(c["pct_high"])} · {c["high"]} usuarios</div></div>'
                    f'<div class="stat-item"><div class="stat-k">Rara vez o nunca (1–2)</div><div class="stat-v" style="color:{C.positive}">{fmt_pct(c["pct_low"])} · {c["low"]} usuarios</div></div>'
                    '</div>')
                if "sat_entry_time" in ctx.users:
                    pair = ctx.users[["congestion", "sat_entry_time"]].dropna()
                    if len(pair) >= N_MIN_SEGMENT and pair.nunique().min() > 1:
                        r = pair.corr(method="spearman").iloc[0, 1]
                        html_block(f'<div class="card-sub" style="margin-top:.8rem">Correlación de Spearman entre congestión y satisfacción con '
                                   f'el tiempo de ingreso: <b>{fmt_num(r, 2)}</b> (n = {len(pair)}). Es una asociación, no una relación causal.</div>')

    with st.expander("Ver tabla de estadísticos por dimensión"):
        table = sat[["Dimensión", "n", "Media", "DE", "Mediana", "% favorable", "% neutral", "% desfavorable",
                     "c1", "c2", "c3", "c4", "c5"]].rename(columns={f"c{i}": f"Nota {i}" for i in range(1, 6)})
        show_table(table.round(2))


def comment_cards(frame: pd.DataFrame, text_key: str, show_meta: bool = True) -> None:
    for _, row in frame.iterrows():
        tags = "".join(f'<span class="tag">{esc(t)}</span>' for t in row.get(text_key + "_themes", []) or [])
        meta = []
        if show_meta:
            # Solo variables de baja identificabilidad (se omite vehículo: hay categorías con n = 1).
            for k in ["frequency", "tenure"]:
                if k in row and pd.notna(row[k]):
                    meta.append(str(row[k]))
        meta_html = f'<span class="meta">{esc(" · ".join(meta))}</span>' if meta else ""
        html_block(f'<div class="comment"><div class="comment-text">“{esc(row[text_key])}”</div>'
                   f'<div class="comment-meta">{tags}{meta_html}</div></div>')


def render_improvements(ctx: Context) -> None:
    filter_note(ctx)
    m = ctx.metrics
    if not users_guard(ctx, "los inconvenientes y mejoras"):
        return
    u = ctx.users
    has_improve = "improve_text" in u
    n_improve_answers = int(u["improve_text"].notna().sum()) if has_improve else 0
    n_improve_valid = int(u["improve_text_valid"].fillna(False).astype(bool).sum()) if has_improve else 0
    themes = theme_counts(u, "improve_text") if has_improve else pd.DataFrame()
    n_themes = int((~themes["Tema"].isin([THEME_OTHER, THEME_NONE])).sum()) if not themes.empty else 0
    kpi_row([
        Kpi("Reportaron inconvenientes", fmt_pct(safe_div(m.n_issue_yes, m.n_issue_valid) * 100),
            desc=f"{m.n_issue_yes} de {m.n_issue_valid} usuarios con respuesta", icon_name="alert",
            badge="Pocos casos: lectura cualitativa" if 0 < m.n_issue_yes < N_MIN_SEGMENT else "", badge_kind="warn"),
        Kpi("Comentarios de mejora", str(n_improve_valid), unit=f"/ {n_improve_answers}",
            desc="Comentarios con contenido sobre respuestas recibidas", icon_name="message"),
        Kpi("Temas identificados", str(n_themes), desc="Agrupación por palabras clave presentes en los textos", icon_name="search"),
    ])

    left, right = st.columns([5, 7], gap="medium")
    with left:
        with card("Inconvenientes reportados", "Respuestas a la pregunta condicional (solo quienes dijeron «Sí»).",
                  n=m.n_issue_yes):
            if "issue_text" in u:
                issues = u[is_true(u["had_issue"]) & u["issue_text_valid"].fillna(False).astype(bool)]
                if issues.empty:
                    empty_state("Sin inconvenientes descritos", "Nadie en este subconjunto describió un inconveniente.")
                else:
                    it = theme_counts(issues, "issue_text")
                    html_block('<div class="stat-list">' + "".join(
                        f'<div class="stat-item"><div class="stat-k">{esc(r["Tema"])}</div>'
                        f'<div class="stat-v">{int(r["Comentarios"])} de {len(issues)}</div></div>' for _, r in it.iterrows()) + "</div>")
                    st.markdown('<div style="height:.6rem"></div>', unsafe_allow_html=True)
                    comment_cards(issues, "issue_text")
            else:
                empty_state("Pregunta no disponible", "El archivo no contiene la pregunta de inconvenientes.")
    with right:
        with card("¿Qué deberían mejorar?", "Temas mencionados en la pregunta abierta. Un comentario puede tener varios temas; los % no suman 100.",
                  n=n_improve_valid):
            if themes.empty:
                empty_state("Sin comentarios", "No hay comentarios válidos en este subconjunto.")
            else:
                show_chart(theme_chart(themes, n_improve_valid), "imp_themes")
                with st.popover("¿Cómo se agrupan los comentarios?") if hasattr(st, "popover") else st.expander("¿Cómo se agrupan los comentarios?"):
                    st.markdown("Cada comentario se asigna a los temas cuyas palabras clave aparecen en el texto "
                                "(comparación sin tildes ni mayúsculas). No se usan modelos externos.")
                    st.markdown("\n".join(f"- **{name}**: {', '.join(kw)}" for name, kw in IMPROVEMENT_THEMES))
                    st.markdown(f"- **{THEME_NONE}**: respuestas como «Nose» o «Ninguno».\n- **{THEME_OTHER}**: sin coincidencias.")

    if has_improve:
        with card("Comentarios originales", "Texto tal como fue escrito por cada usuario. Busca o filtra por tema."):
            a, b = st.columns([6, 6])
            query = a.text_input("Buscar en comentarios", key="comment_query", placeholder="Ej.: piso, acceso, lejos…")
            # Opciones fijas (de todo el archivo) para que un filtro no invalide la selección guardada.
            all_users = ctx.ds.data[is_true(ctx.ds.data["used"])]
            all_themes = theme_counts(all_users, "improve_text")["Tema"].tolist()
            sel_themes = b.multiselect("Tema", all_themes, key="comment_themes", placeholder="Todos los temas")
            pool = u[u["improve_text_valid"].fillna(False).astype(bool)].copy()
            if query:
                q = normalize_text(query)
                pool = pool[pool["improve_text"].map(lambda t: q in normalize_text(t))]
            if sel_themes:
                pool = pool[pool["improve_text_themes"].map(lambda ts: any(t in ts for t in sel_themes))]
            html_block(f'<div class="card-sub" style="margin:.2rem 0 .7rem">{len(pool)} comentario(s) mostrados de {n_improve_valid}.</div>')
            if pool.empty:
                empty_state("Ningún comentario coincide", "Prueba con otro término o quita el filtro de tema.", "search")
            else:
                try:
                    scroller = st.container(height=520 if len(pool) > 6 else None, border=False)
                except TypeError:
                    scroller = st.container()
                with scroller:
                    comment_cards(pool, "improve_text")
            excluded = n_improve_answers - n_improve_valid
            if excluded:
                html_block(f'<div class="card-sub">{excluded} respuesta(s) sin contenido (p. ej. «.») no se muestran ni se cuentan como comentario válido.</div>')


def render_nps(ctx: Context) -> None:
    filter_note(ctx)
    if not users_guard(ctx, "la recomendación"):
        return
    p = ctx.metrics.nps
    if not p or not p["n"]:
        empty_state("Sin respuestas de recomendación", "No hay puntuaciones válidas (0–10) en este subconjunto.")
        return
    kpi_row([
        Kpi("Net Promoter Score", fmt_signed(p["nps"]), desc="% promotores − % detractores (rango −100 a +100)", icon_name="megaphone",
            badge=f"IC 95 % aprox. ± {fmt_num(p['ci95'], 0)} pts" if p["n"] > 1 else "n insuficiente", badge_kind="neg" if p["nps"] < 0 else "pos"),
        Kpi("Promotores", fmt_pct(p["pct_prom"]), desc=f"{p['promoters']} usuarios con 9–10", icon_name="trend_up", badge="9–10", badge_kind="pos"),
        Kpi("Pasivos", fmt_pct(p["pct_pas"]), desc=f"{p['passives']} usuarios con 7–8", icon_name="users", badge="7–8", badge_kind="warn"),
        Kpi("Detractores", fmt_pct(p["pct_det"]), desc=f"{p['detractors']} usuarios con 0–6", icon_name="trend_down", badge="0–6", badge_kind="neg"),
        Kpi("Puntuación media", fmt_num(p["mean"], 1), unit="/ 10", desc="Promedio simple; no equivale al NPS", icon_name="star"),
    ])
    left, right = st.columns([5, 7], gap="medium")
    with left:
        with card("Net Promoter Score", "Bandas de referencia: < 0 · 0–30 · > 30.", n=p["n"]):
            show_chart(nps_gauge(p["nps"]), "nps_gauge")
            segs = [(p["pct_det"], C.negative, "Detractores"), (p["pct_pas"], C.warning, "Pasivos"), (p["pct_prom"], C.positive, "Promotores")]
            html_block('<div class="split">' + "".join(f'<div style="width:{v}%;background:{c}"></div>' for v, c, _ in segs) + "</div>"
                       '<div class="legend-row">' + "".join(f'<span><span class="sw" style="background:{c}"></span>{n} <b>{fmt_pct(v)}</b></span>' for v, c, n in segs) + "</div>")
    with right:
        with card("Distribución de puntuaciones", "Rojo: detractores (0–6) · ámbar: pasivos (7–8) · verde: promotores (9–10).", n=p["n"]):
            show_chart(nps_distribution_chart(p["dist"]), "nps_dist")

    lo, hi = p["nps"] - p["ci95"], p["nps"] + p["ci95"]
    direction = ("negativo: hay más detractores que promotores" if p["nps"] < 0 else
                 "positivo: hay más promotores que detractores" if p["nps"] > 0 else "neutro: promotores y detractores se igualan")
    ci_text = (f"Con n = {p['n']}, el intervalo de confianza aproximado del 95 % va de {fmt_signed(lo)} a {fmt_signed(hi)} puntos"
               + (", y queda completamente por debajo de cero." if hi < 0 else
                  ", y queda completamente por encima de cero." if lo > 0 else ", e incluye el cero.")) if p["n"] > 1 else ""
    note(f"<b>Interpretación prudente.</b> El NPS es {direction}. {esc(ci_text)} "
         "La muestra es pequeña y no aleatoria, así que el resultado describe a quienes respondieron y no debe generalizarse a toda la comunidad universitaria. "
         "Las bandas de color del indicador son referencias convencionales, no metas institucionales.", "info")


def render_data_explorer(ctx: Context) -> None:
    filter_note(ctx)
    d = ctx.filtered
    if d.empty:
        empty_state("No hay registros con estos filtros", "Restablece los filtros para ver datos.", "search")
        return
    columns = {
        "id": "ID", "used_label": "Ha usado el parqueadero", "role": LABELS["role"], "age": "Edad",
        "vehicle": LABELS["vehicle"], "frequency": LABELS["frequency"], "schedule": LABELS["schedule"],
        "tenure": LABELS["tenure"], **{k: LABELS[k] for k in SATISFACTION_KEYS}, "congestion": "Congestión (1 nunca–5 siempre)",
        "sat_index": "Índice satisfacción (calculado)", "had_issue_label": "Tuvo inconvenientes",
        "issue_text": LABELS["issue_text"], "recommend": "Recomendación (0–10)", "nps_group": "Grupo NPS (calculado)",
        "improve_text": LABELS["improve_text"],
    }
    columns = {k: v for k, v in columns.items() if k in d}
    default_cols = [k for k in ["id", "used_label", "role", "vehicle", "frequency", "sat_index", "recommend", "improve_text"] if k in columns]

    with card("Registros de la encuesta", "Las marcas de tiempo se omiten para reducir la identificabilidad de los encuestados.", n=len(d)):
        a, b, c = st.columns([4, 5, 3])
        query = a.text_input("Buscar", key="exp_query", placeholder="Texto en cualquier columna…")
        selected = b.multiselect("Columnas", list(columns), default=default_cols, format_func=columns.get, key="exp_cols")
        scope = c.selectbox("Registros", ["Todos", "Solo usuarios", "Solo no usuarios"], key="exp_scope")
        view = d.copy()
        if scope == "Solo usuarios":
            view = view[is_true(view["used"])]
        elif scope == "Solo no usuarios":
            view = view[is_false(view["used"])]
        cols = selected or default_cols
        table = view[cols].rename(columns=columns)
        if query:
            q = normalize_text(query)
            match = table.astype(str).apply(lambda col: col.map(lambda v: q in normalize_text(v))).any(axis=1)
            table = table[match]
        if "Índice satisfacción (calculado)" in table:
            table["Índice satisfacción (calculado)"] = table["Índice satisfacción (calculado)"].round(2)
        html_block(f'<div class="card-sub" style="margin:.1rem 0 .6rem">{len(table)} registro(s) · {len(table.columns)} columna(s)</div>')
        if table.empty:
            empty_state("Sin coincidencias", "Ningún registro contiene el texto buscado.", "search")
        else:
            config = {}
            if hasattr(st, "column_config"):
                for k in SATISFACTION_KEYS + ["congestion"]:
                    if columns.get(k) in table:
                        config[columns[k]] = st.column_config.NumberColumn(format="%d")
                if "Recomendación (0–10)" in table:
                    config["Recomendación (0–10)"] = st.column_config.ProgressColumn(min_value=0, max_value=10, format="%d")
            show_table(table, column_config=config, height=min(560, 38 * len(table) + 40))
            d1, d2, _ = st.columns([2, 2, 6])
            d1.download_button("Descargar CSV", table.to_csv(index=False).encode("utf-8-sig"),
                               file_name="parqueadero_filtrado.csv", mime="text/csv", key="dl_csv")
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                table.to_excel(writer, index=False, sheet_name="Datos filtrados")
            d2.download_button("Descargar Excel", buffer.getvalue(), file_name="parqueadero_filtrado.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_xlsx")


def missing_report(ds: Dataset) -> pd.DataFrame:
    """Clasifica los faltantes de cada columna mapeada."""
    d = ds.data
    nonuser = is_false(d["used"])
    rows = []
    for spec in COLUMN_SPECS:
        if spec.key not in ds.mapping or spec.key == "used":
            continue
        col = d[spec.key]
        miss = col.isna()
        n_miss = int(miss.sum())
        if n_miss == 0:
            kind = "Sin faltantes"
        elif n_miss == len(d):
            kind = "Columna vacía en el archivo de origen"
        elif spec.key == "issue_text":
            problem = int((miss & is_true(d["had_issue"])).sum())
            kind = "Condicional: solo aplica a quienes tuvieron inconvenientes" + (f" · {problem} faltante(s) problemático(s)" if problem else "")
        else:
            expected = int((miss & nonuser).sum())
            problem = n_miss - expected
            kind = "Esperado: no usuarios (la encuesta termina en la primera pregunta)" if problem == 0 else \
                f"{expected} esperado(s) · {problem} faltante(s) problemático(s) entre usuarios"
        rows.append({"Variable": spec.label, "Respuestas válidas": int(col.notna().sum()), "Faltantes": n_miss, "Tipo de faltante": kind,
                     "Encabezado original": ds.mapping[spec.key]})
    return pd.DataFrame(rows)


def render_methodology(ctx: Context) -> None:
    ds = ctx.ds
    d = ds.data
    n_users = int(is_true(d["used"]).sum())
    left, right = st.columns([6, 6], gap="medium")
    with left:
        with card("Definiciones y denominadores"):
            html_block(f"""<div class="method">
<h4>Poblaciones</h4>
<ul><li><b>Encuestados</b>: todas las filas del archivo ({len(d)}).</li>
<li><b>Usuarios</b>: respondieron «Sí» a haber usado el parqueadero este semestre ({n_users}). Solo ellos ven el resto del cuestionario.</li>
<li><b>Respuestas válidas</b>: valores no vacíos dentro de la escala de cada pregunta. Cada porcentaje usa como denominador las respuestas válidas de esa pregunta; una respuesta vacía nunca se interpreta como negativa.</li></ul>
<h4>Satisfacción (1–5)</h4>
<ul><li>Favorable = 4–5 · Neutral = 3 · Desfavorable = 1–2.</li>
<li><b>Índice de satisfacción</b> = media de las medias de las 9 dimensiones. El índice individual (usado en segmentos) es la media de las dimensiones respondidas por cada persona.</li>
<li>La pregunta de <b>congestión</b> mide frecuencia (1 = nunca, 5 = siempre) y se analiza aparte; no forma parte del índice.</li></ul>
<h4>Recomendación (0–10)</h4>
<ul><li>Detractores 0–6 · Pasivos 7–8 · Promotores 9–10. NPS = % promotores − % detractores.</li>
<li>IC 95 % aproximado: ±1,96·√((p + d − (p − d)²)/n)·100, con p y d las proporciones de promotores y detractores.</li></ul>
<h4>Comentarios abiertos</h4>
<ul><li>Se agrupan por palabras clave presentes en los textos; un comentario puede tener varios temas. Las respuestas sin letras ni números (p. ej. «.») no cuentan como comentario válido. Los textos originales se conservan íntegros.</li></ul>
<h4>Alcance</h4>
<ul><li>La muestra es pequeña y no probabilística: los resultados son descriptivos. Las asociaciones entre variables no implican causalidad.</li>
<li>Los filtros de perfil solo aplican a usuarios, porque los no usuarios no respondieron esas preguntas.</li></ul></div>""")
    with right:
        with card("Calidad de los datos", f"Fuente: {ds.source_name} · {len(ds.raw)} filas × {ds.raw.shape[1]} columnas"):
            if ctx.period:
                html_block(f'<div class="card-sub">Periodo de recolección (según «Hora de inicio»): <b>{esc(ctx.period)}</b>.</div>')
            if "consent" in d:
                c_yes = int(is_true(d["consent"]).sum())
                html_block(f'<div class="card-sub" style="margin-top:.4rem">Consentimiento: {c_yes} de {n_users} usuarios aceptaron participar; '
                           'la pregunta no se mostró a los no usuarios.</div>')
            for n_ in ds.notes:
                note(esc(n_))
            if ds.unmapped:
                html_block(f'<div class="card-sub" style="margin-top:.4rem">Columnas no utilizadas: {esc(", ".join(ds.unmapped))}</div>')
            if "issue_text" in d:
                stray = d[~is_true(d["had_issue"]) & d["issue_text"].notna()]
                if len(stray):
                    html_block(f'<div class="card-sub" style="margin-top:.4rem">{len(stray)} persona(s) respondieron «No» a inconvenientes pero escribieron '
                               f'en la pregunta condicional ({esc(", ".join(stray["issue_text"].astype(str)))}); ese texto no se cuenta como inconveniente.</div>')
    with card("Mapa de columnas y valores faltantes", "Cómo se identificó cada pregunta del Excel y qué tipo de faltante presenta."):
        show_table(missing_report(ds))


# =============================================================================
# 7. SIDEBAR Y PUNTO DE ENTRADA
# =============================================================================


def resolve_source() -> tuple[bytes | None, str]:
    """Obtiene los bytes del Excel desde el uploader, variable de entorno o archivo local."""
    uploaded = st.session_state.get("uploaded_file")
    if uploaded is not None:
        return uploaded.getvalue(), uploaded.name
    env_path = os.environ.get(DATA_PATH_ENV_VAR)
    candidates = [Path(env_path)] if env_path else []
    candidates.append(Path(__file__).resolve().parent / DEFAULT_DATA_FILE)
    for path in candidates:
        if path.is_file():
            return path.read_bytes(), path.name
    return None, ""


def render_sidebar(ds: Dataset | None) -> str:
    with st.sidebar:
        html_block(f'<div class="brand"><div class="brand-logo">{icon("parking", 20, "white")}</div>'
                   f'<div><div class="brand-name">{esc(APP_TITLE)}</div><div class="brand-sub">Analítica de experiencia de servicio</div></div></div>')
        html_block('<div class="side-label">Secciones</div>')
        section = st.radio("Navegación", list(SECTIONS), format_func=lambda k: SECTIONS[k][0],
                           key="nav", label_visibility="collapsed")
        if ds is not None and not ds.missing_required:
            data = ds.data
            html_block('<div class="side-label">Filtros</div>')
            for key, label in FILTER_CATEGORIES:
                if key in data:
                    st.multiselect(label, ordered_categories(data[key].dropna().unique().tolist(), key),
                                   key=f"f_{key}", placeholder="Todos")
            if "f_age" in st.session_state and "age" in data:
                lo, hi = int(data["age"].min()), int(data["age"].max())
                if lo < hi:
                    st.slider("Edad", lo, hi, key="f_age")
            if "had_issue" in data:
                st.selectbox("Inconvenientes", ISSUE_FILTER_OPTIONS, key="f_issue")
            st.button("Restablecer filtros", on_click=reset_filters, args=(data,), key="reset_btn",
                      help="Vuelve a mostrar todos los registros")
        html_block('<div class="side-label">Fuente de datos</div>')
        with st.expander("Cambiar archivo Excel"):
            st.file_uploader("Archivo .xlsx con la misma estructura", type=["xlsx"], key="uploaded_file",
                             help=f"También puedes indicar una ruta con la variable de entorno {DATA_PATH_ENV_VAR}.")
            if ds is not None:
                st.caption(f"En uso: {ds.source_name}")
    return section


def collection_period(data: pd.DataFrame) -> str | None:
    if "start_time" not in data or data["start_time"].notna().sum() == 0:
        return None
    months = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
    lo, hi = data["start_time"].min(), data["start_time"].max()
    f = lambda t: f"{t.day} {months[t.month - 1]} {t.year}"  # noqa: E731
    return f"{f(lo)} – {f(hi)}" if lo.date() != hi.date() else f(lo)


RENDERERS = {
    "overview": render_overview, "profile": render_demographics, "usage": render_usage_analysis,
    "satisfaction": render_satisfaction, "improvements": render_improvements, "nps": render_nps,
    "explorer": render_data_explorer, "method": render_methodology,
}


def main() -> None:
    st.set_page_config(page_title=f"{APP_TITLE} · Dashboard", page_icon="🅿️", layout="wide",
                       initial_sidebar_state="expanded")
    inject_css()
    _card_counter["i"] = 0

    file_bytes, source_name = resolve_source()
    ds: Dataset | None = None
    error: str | None = None
    if file_bytes is None:
        error = (f"No se encontró «{DEFAULT_DATA_FILE}» junto a app.py. Cárgalo desde la barra lateral "
                 f"(Fuente de datos) o define la variable de entorno {DATA_PATH_ENV_VAR}.")
    else:
        try:
            with st.spinner("Cargando y validando datos…"):
                raw = load_data(file_bytes)
                ds = clean_data(raw, source_name)
        except Exception as exc:  # noqa: BLE001 — se informa al usuario sin detener la app
            error = f"No fue posible leer el archivo «{source_name}»: {exc}"

    if ds is not None and not ds.missing_required:
        init_filter_state(ds.data, f"{source_name}:{len(file_bytes)}")

    section = render_sidebar(ds)

    if error:
        html_block(f'<div class="page-head"><div><div class="eyebrow">{esc(APP_TITLE)}</div>'
                   f'<h1 class="page-title">No hay datos para mostrar</h1></div></div>')
        empty_state("Archivo no disponible", error, "alert")
        return
    assert ds is not None
    if ds.missing_required:
        html_block(f'<div class="page-head"><div><div class="eyebrow">{esc(APP_TITLE)}</div>'
                   f'<h1 class="page-title">El archivo no tiene la estructura esperada</h1></div></div>')
        empty_state("Faltan columnas obligatorias", "No se encontraron: " + "; ".join(ds.missing_required), "alert")
        with st.expander("Columnas presentes en el archivo"):
            st.write(list(ds.raw.columns))
        return

    active = active_filters(ds.data)
    filtered = apply_filters(ds.data, active)
    metrics = prepare_metrics(filtered)
    ctx = Context(ds=ds, filtered=filtered, users=filtered[is_true(filtered["used"])], metrics=metrics,
                  active=active, period=collection_period(ds.data))

    page_header(section, ctx.period, len(filtered), len(ds.data), len(active))
    RENDERERS[section](ctx)
    html_block(f'<div class="foot">Fuente: encuesta de servicio del parqueadero Puente Madera · {len(ds.data)} respuestas · '
               'Resultados descriptivos de una muestra pequeña y no probabilística.</div>')


if __name__ == "__main__":
    main()
