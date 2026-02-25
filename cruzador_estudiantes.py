# cruzador_estudiantes.py
import io
import re
import hashlib
from pathlib import Path
from typing import Dict, Optional, Tuple, List

import numpy as np
import pandas as pd
import streamlit as st
import difflib


# =========================
# Config
# =========================
ROOT = Path(__file__).resolve().parent
PLANTILLAS_DIR = ROOT / "plantillas"

DEFAULT_AKADEMIC = ROOT / "alumnos-en akademic.xlsx"
DEFAULT_CURLLE = ROOT / "Data Historica Notas Curlle (1).csv"
DEFAULT_PLANES = ROOT / "todos los planes-juntos.xlsx"
DEFAULT_TODOS_PLANES_AKADEMIC = ROOT / "TODOS_PLANES_AKADEMIC.xlsx"

P01_TEMPLATE = PLANTILLAS_DIR / "P01_existentes_en_akademic.xlsx"
P02_TEMPLATE = PLANTILLAS_DIR / "P02_no_akademic_no_curlle.xlsx"
P03_TEMPLATE = PLANTILLAS_DIR / "P03_no_akademic_si_curlle.xlsx"
P04_TEMPLATE = PLANTILLAS_DIR / "P04_matricula_desde_akademic.xlsx"
P05_TEMPLATE = PLANTILLAS_DIR / "P05_carga_notas_akademic.xlsx"


# =========================
# Helpers base (BLINDADOS)
# =========================
def norm_col(c: str) -> str:
    """
    Normaliza encabezados para que:
    - Respete '_' (underscore)
    - Convierta '.', '-', '/', etc. en espacios
    - Quite tildes y caracteres raros
    - Devuelva snake_case consistente
    """
    s = "" if c is None else str(c)
    s = s.strip().lower()

    # separadores comunes a espacio
    s = re.sub(r"[.\-\/\\]+", " ", s)

    # espacios múltiples
    s = re.sub(r"\s+", " ", s)

    # quitar tildes
    s = (
        s.replace("á", "a")
        .replace("é", "e")
        .replace("í", "i")
        .replace("ó", "o")
        .replace("ú", "u")
        .replace("ñ", "n")
    )

    # dejar solo letras, números, espacio y underscore
    s = re.sub(r"[^a-z0-9 _]+", "", s)

    # snake_case
    s = s.replace(" ", "_")
    s = re.sub(r"_+", "_", s).strip("_")
    return s


def norm_text_keep_spaces(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.strip().lower()
    s = re.sub(r"\s+", " ", s)
    s = (
        s.replace("á", "a")
        .replace("é", "e")
        .replace("í", "i")
        .replace("ó", "o")
        .replace("ú", "u")
        .replace("ñ", "n")
    )
    s = re.sub(r"[^a-z0-9 ]+", "", s)
    return s.strip()


def only_digits(s) -> str:
    if s is None:
        return ""
    return re.sub(r"\D+", "", str(s))


def zfill8(s) -> str:
    d = only_digits(s)
    if not d:
        return ""
    return d.zfill(8)


def periodo_formato(val) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    s = re.sub(r"\.0$", "", s)
    if re.fullmatch(r"\d{5}", s):  # 20201
        return f"{s[:4]}-{s[4:]}"
    return s


def pick_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    """
    Busca una columna del df usando candidatos.
    - normaliza candidates
    - además intenta match "suave" quitando underscores
      (ej: codigo_estudiante == codigoestudiante)
    """
    cols = list(df.columns)
    cols_set = set(cols)
    soft_map = {c.replace("_", ""): c for c in cols}

    for cand in candidates:
        cn = norm_col(cand)
        if cn in cols_set:
            return cn
        cn_soft = cn.replace("_", "")
        if cn_soft in soft_map:
            return soft_map[cn_soft]

    return None


def build_fullname_from_parts(padron: pd.DataFrame, ap_col: str, am_col: str, nom_col: str) -> pd.Series:
    return (
        padron[ap_col].fillna("").astype(str).str.strip()
        + " "
        + padron[am_col].fillna("").astype(str).str.strip()
        + " "
        + padron[nom_col].fillna("").astype(str).str.strip()
    ).str.replace(r"\s+", " ", regex=True).str.strip()


# =========================
# ✅ LECTURA EXCEL ESTABLE (SIN openpyxl)
# =========================
HEADER_HINTS = [
    "dni", "documento", "codigo_estudiante", "codigo_alumno", "codigo",
    "apellidos", "nombres", "apellido", "programa", "carrera",
    "plan", "plan_sigu", "periodo", "periodo_ingreso"
]


def _to_bytes(uploaded_or_path) -> bytes:
    if uploaded_or_path is None:
        return b""
    if hasattr(uploaded_or_path, "getvalue"):
        return uploaded_or_path.getvalue()
    if hasattr(uploaded_or_path, "read"):
        pos = uploaded_or_path.tell()
        b = uploaded_or_path.read()
        try:
            uploaded_or_path.seek(pos)
        except Exception:
            pass
        return b
    return Path(uploaded_or_path).read_bytes()


def _best_header_row_from_preview(preview_df: pd.DataFrame, max_rows: int = 60) -> int:
    best_i = 0
    best_score = -1
    hints = [norm_col(h) for h in HEADER_HINTS]

    for i in range(min(max_rows, len(preview_df))):
        row = preview_df.iloc[i].tolist()
        tokens = []
        for v in row:
            if v is None:
                continue
            s = str(v).strip()
            if s == "":
                continue
            tokens.append(norm_col(s))

        if not tokens:
            continue

        joined = " ".join(tokens)
        score = 0
        for hn in hints:
            if hn in tokens or hn in joined:
                score += 1

        non_empty = sum(1 for t in tokens if t != "")
        if non_empty >= 8:
            score += 2
        if non_empty >= 12:
            score += 2

        if score > best_score:
            best_score = score
            best_i = i

    if best_score < 2:
        return 0
    return best_i


def read_excel_any(file_or_path, sheet_name=0) -> pd.DataFrame:
    """
    Lee Excel usando SOLO python-calamine (ideal para Streamlit Cloud).
    Si falta calamine, muestra error claro.
    """
    bio = io.BytesIO(_to_bytes(file_or_path))

    try:
        # preview
        bio.seek(0)
        preview = pd.read_excel(
            bio, dtype=str, engine="calamine",
            sheet_name=sheet_name, header=None, nrows=60
        )
        header_i = _best_header_row_from_preview(preview, max_rows=60)

        # lectura final
        bio.seek(0)
        df = pd.read_excel(
            bio, dtype=str, engine="calamine",
            sheet_name=sheet_name, header=0, skiprows=header_i
        )
    except Exception as e:
        raise RuntimeError(
            "No se pudo leer el Excel.\n"
            "Asegúrate de tener instalado 'python-calamine'.\n"
            f"Error: {type(e).__name__}: {e}"
        )

    df.columns = [norm_col(c) for c in df.columns]
    return df


# =========================
# Curlle CSV (robusto)
# =========================
def find_header_row_in_curlle(csv_bytes_or_path) -> int:
    if hasattr(csv_bytes_or_path, "read"):
        raw = csv_bytes_or_path.read()
        csv_bytes_or_path.seek(0)
        lines = raw.splitlines()
    else:
        raw = Path(csv_bytes_or_path).read_bytes()
        lines = raw.splitlines()

    pats = [b"Codigo Alumno", "Código Alumno".encode("utf-8")]
    for i, line in enumerate(lines[:400]):
        if any(p in line for p in pats):
            return i
    return 3


def read_curlle(csv_file_or_path) -> pd.DataFrame:
    header_idx = find_header_row_in_curlle(csv_file_or_path)
    df = pd.read_csv(
        csv_file_or_path,
        dtype=str,
        sep=";",
        encoding="utf-8-sig",
        engine="python",
        skiprows=header_idx,
    )
    df.columns = [norm_col(c) for c in df.columns]
    return df


def template_columns(template_path: Path) -> Tuple[pd.DataFrame, str]:
    """
    ⚠️ Para leer templates, también usamos calamine (sin openpyxl).
    """
    wb = pd.read_excel(template_path, sheet_name=None, dtype=str, engine="calamine")
    sheet = list(wb.keys())[0]
    df = wb[sheet].copy()
    df.columns = ["" if c is None else str(c) for c in df.columns]
    return df, sheet


def align_df_to_template_df(template_path: Path, data: pd.DataFrame) -> pd.DataFrame:
    tpl_df, _ = template_columns(template_path)
    tpl_cols = list(tpl_df.columns)

    tpl_map = {norm_col(c): c for c in tpl_cols}
    data_cols_map = {norm_col(c): c for c in data.columns}

    out = pd.DataFrame(columns=tpl_cols)
    for nkey, original_tpl_col in tpl_map.items():
        if nkey in data_cols_map:
            out[original_tpl_col] = data[data_cols_map[nkey]]
        else:
            out[original_tpl_col] = ""
    return out


def _sanitize_df_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    if df is None:
        return pd.DataFrame()

    out = df.copy()
    out.columns = ["" if c is None else str(c) for c in out.columns]
    out = out.replace({np.nan: "", pd.NaT: ""})

    for c in out.columns:
        out[c] = out[c].map(lambda v: "" if v is None else str(v))

    return out


def build_multi_sheet_excel(sheets: Dict[str, pd.DataFrame]) -> bytes:
    """
    Genera Excel usando SOLO xlsxwriter (sin openpyxl).
    """
    safe_sheets = {}
    for sheet_name, df in sheets.items():
        sh = "" if sheet_name is None else str(sheet_name)
        sh = sh.replace("/", "_").replace("\\", "_").replace("[", "(").replace("]", ")").replace("*", "_").replace("?", "_").replace(":", "_")
        sh = sh[:31] if sh else "Sheet"
        safe_sheets[sh] = _sanitize_df_for_excel(df)

    bio = io.BytesIO()

    try:
        with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
            for sh, df in safe_sheets.items():
                df.to_excel(writer, index=False, sheet_name=sh)
        return bio.getvalue()
    except Exception as e:
        raise RuntimeError(
            "No se pudo generar el Excel.\n"
            "Asegúrate de tener instalado 'xlsxwriter'.\n"
            f"Error: {type(e).__name__}: {e}"
        )


# =========================
# Resto de tu lógica (SIN CAMBIOS)
# =========================
def resolve_course_name_and_origin(pl2: pd.DataFrame, carrera_key: str, plan_key: str, cod_curso_key: str) -> Tuple[str, str, str]:
    carrera_key = "" if carrera_key is None else str(carrera_key).strip()
    plan_key = "" if plan_key is None else str(plan_key).strip()
    cod_curso_key = "" if cod_curso_key is None else str(cod_curso_key).strip().upper()

    if cod_curso_key == "":
        return "", carrera_key, plan_key

    exact = pl2[
        (pl2["carrera_key"] == carrera_key)
        & (pl2["plan_key"] == plan_key)
        & (pl2["cod_curso_key"] == cod_curso_key)
    ]
    if len(exact) > 0:
        row = exact.iloc[0]
        curso = str(row.get("curso", "")).strip()
        return curso, carrera_key, plan_key

    by_course = pl2[pl2["cod_curso_key"] == cod_curso_key].copy()
    if len(by_course) == 0:
        return "", carrera_key, plan_key

    by_course["curso"] = by_course["curso"].fillna("").astype(str).str.strip()
    cursos_no_vacios = [c for c in by_course["curso"].unique().tolist() if c != ""]
    cursos_no_vacios = list(dict.fromkeys(cursos_no_vacios))

    if len(cursos_no_vacios) == 1:
        chosen_curso = cursos_no_vacios[0]
        chosen_row = by_course[by_course["curso"] == chosen_curso].iloc[0]
        carrera_res = str(chosen_row.get("carrera_key", "")).strip()
        plan_res = str(chosen_row.get("plan_key", "")).strip()
        return chosen_curso, carrera_res, plan_res

    return "", carrera_key, plan_key


def coalesce_series(a: pd.Series, b: pd.Series) -> pd.Series:
    a = a.fillna("").astype(str)
    b = b.fillna("").astype(str)
    out = a.copy()
    mask = out.str.strip().eq("")
    out[mask] = b[mask]
    return out


# =========================
# ✅ Cache robusto (por HASH)
# =========================
def _hash_bytes(b: bytes) -> str:
    return hashlib.md5(b).hexdigest()


@st.cache_data(show_spinner=False)
def _cached_read_excel_bytes(file_bytes: bytes, _md5: str) -> pd.DataFrame:
    return read_excel_any(io.BytesIO(file_bytes))


@st.cache_data(show_spinner=False)
def _cached_read_excel_path(path_str: str) -> pd.DataFrame:
    return read_excel_any(path_str)


@st.cache_data(show_spinner=False)
def _cached_read_curlle_bytes(file_bytes: bytes, _md5: str) -> pd.DataFrame:
    return read_curlle(io.BytesIO(file_bytes))


@st.cache_data(show_spinner=False)
def _cached_read_curlle_path(path_str: str) -> pd.DataFrame:
    return read_curlle(path_str)


# =========================
# UI
# =========================
st.set_page_config(page_title="Regularización AKA", layout="wide")
st.title("📌 Regularización AKA - Generador de Plantillas")

st.markdown(
    """
Sube el **padrón principal** (archivo con alumnos).
La app cruza contra:
- **Akademic** (existentes) por **DNI o Código**
- **Curlle** (histórica de notas) por **Código Alumno**
- **Planes** (todos los planes-juntos) para obtener **Nom. Curso**
- **TODOS_PLANES_AKADEMIC** para armar **CodigoCurso** correcto en P04
"""
)

with st.sidebar:
    st.header("📥 Cargas")

    st.subheader("📌 Base detectada en la raíz")
    st.caption(f"Carpeta actual: {ROOT}")

    ok_ak = DEFAULT_AKADEMIC.exists()
    ok_cu = DEFAULT_CURLLE.exists()
    ok_pl = DEFAULT_PLANES.exists()
    ok_tpa = DEFAULT_TODOS_PLANES_AKADEMIC.exists()

    st.write(("✅" if ok_ak else "❌"), DEFAULT_AKADEMIC.name)
    st.write(("✅" if ok_cu else "❌"), DEFAULT_CURLLE.name)
    st.write(("✅" if ok_pl else "❌"), DEFAULT_PLANES.name)
    st.write(("✅" if ok_tpa else "❌"), DEFAULT_TODOS_PLANES_AKADEMIC.name)

    st.divider()
    padron_file = st.file_uploader("1) Padrón (OBLIGATORIO) - Excel/CSV", type=["xlsx", "xls", "csv"])

    st.divider()
    usar_override = st.checkbox("Quiero reemplazar archivos base (override)", value=False)

    akademic_file = None
    curlle_file = None
    planes_file = None
    todos_planes_ak_file = None

    if usar_override:
        akademic_file = st.file_uploader("Akademic (xlsx)", type=["xlsx", "xls"])
        curlle_file = st.file_uploader("Curlle (csv)", type=["csv"])
        planes_file = st.file_uploader("Planes (xlsx)", type=["xlsx", "xls"])
        todos_planes_ak_file = st.file_uploader("TODOS_PLANES_AKADEMIC (xlsx)", type=["xlsx", "xls"])

    st.divider()
    run_btn = st.button("🚀 Generar plantillas", type="primary")

if not padron_file:
    st.info("Sube el **padrón** para empezar.")
    st.stop()

if not run_btn:
    st.stop()

if not akademic_file and not DEFAULT_AKADEMIC.exists():
    st.error(f"No encuentro {DEFAULT_AKADEMIC.name} en la raíz. Marca override y súbelo.")
    st.stop()
if not curlle_file and not DEFAULT_CURLLE.exists():
    st.error(f"No encuentro {DEFAULT_CURLLE.name} en la raíz. Marca override y súbelo.")
    st.stop()
if not planes_file and not DEFAULT_PLANES.exists():
    st.error(f"No encuentro {DEFAULT_PLANES.name} en la raíz. Marca override y súbelo.")
    st.stop()
if not todos_planes_ak_file and not DEFAULT_TODOS_PLANES_AKADEMIC.exists():
    st.error(f"No encuentro {DEFAULT_TODOS_PLANES_AKADEMIC.name} en la raíz. Marca override y súbelo.")
    st.stop()

# =========================
# ✅ Guard por tamaño
# =========================
MAX_MB = 35
if padron_file and padron_file.name.lower().endswith((".xlsx", ".xls")):
    size_mb = len(padron_file.getvalue()) / (1024 * 1024)
    if size_mb > MAX_MB:
        st.error(f"Tu padrón pesa {size_mb:.1f} MB. Recomendado < {MAX_MB} MB o súbelo como CSV.")
        st.stop()


# =========================
# Load Padrón (estable)
# =========================
with st.spinner("Leyendo PADRÓN..."):
    if padron_file.name.lower().endswith(".csv"):
        padron = pd.read_csv(padron_file, dtype=str, encoding="utf-8-sig")
        padron.columns = [norm_col(c) for c in padron.columns]
    else:
        padron_bytes = padron_file.getvalue()
        padron = _cached_read_excel_bytes(padron_bytes, _hash_bytes(padron_bytes))

with st.expander("🧪 Debug padrón: columnas detectadas"):
    st.write(padron.columns.tolist()[:120])


# =========================
# TODO: desde acá sigue tu código tal cual (P03, P04, etc.)
# =========================
st.warning(
    "✅ Tu lectura y exportación ya están blindadas para Streamlit Cloud (sin openpyxl).\n"
    "Ahora pega debajo TODO tu bloque restante (desde 'Detección columnas PADRÓN' hasta el final) "
    "SIN CAMBIAR NADA."
)