# ELCATEX - Centro de Optimización de Loteo
# Aplicación monolítica: parser + motor + interfaz Streamlit
# Ejecución: streamlit run app_loteo_elcatex.py

# ============================================
# reglas_operativas_parser.py
# Parser robusto de la hoja REGLAS_OPERATIVAS (formato semi-libre) hacia
# una estructura `params` editable desde la UI de Streamlit.
# ============================================

import re
import pandas as pd
import numpy as np

DEFAULT_MAX_WIDTHS_BY_CAT = {
    "A-4000": 4, "B-3300": 4,
    "C-2600": 3, "D-2200": 3, "F-2200": 3,
    "E-1100": 2, "G-1100": 2,
}

# Pares de bloques permitidos para mezclar en un mismo lote, según
# COMBINACION_PRIORIDAD: PAST DUE+DUE(VENCIDOS), +AHEAD, AHEAD+AHEAD2, OTROS solo con AHEAD2.
DEFAULT_ALLOWED_PAIRS = [
    ("VENCIDOS", "VENCIDOS"),
    ("VENCIDOS", "DUE"),
    ("VENCIDOS", "AHEAD"),
    ("DUE", "DUE"),
    ("DUE", "AHEAD"),
    ("AHEAD", "AHEAD"),
    ("AHEAD", "AHEAD2"),
    ("OTROS", "AHEAD2"),
]

RULE_TOKENS = ["ANCHO18", "COMBO_ANCHOS", "COLOR_R", "FAMILIA"]


def _num(x, default=None):
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return default
        return float(x)
    except Exception:
        return default


def _first_number_in_text(text, default=None):
    if not text:
        return default
    m = re.search(r"(\d+(\.\d+)?)", str(text))
    if m:
        try:
            return float(m.group(1))
        except Exception:
            return default
    return default


def find_reglas_operativas_table(xlsm_path):
    """Localiza la fila de encabezado REGLAS/PRODUCTO/CLAVE1.../OBSERVACION
    dentro de la hoja REGLAS_OPERATIVAS (puede no estar en la fila 1)."""
    raw = pd.read_excel(xlsm_path, sheet_name="REGLAS_OPERATIVAS", engine="openpyxl", header=None)
    header_row = None
    for r in range(min(20, len(raw))):
        vals = [str(v).strip().upper() if pd.notna(v) else "" for v in raw.iloc[r].tolist()]
        if "REGLAS" in vals and "PRODUCTO" in vals:
            header_row = r
            break
    if header_row is None:
        raise ValueError("No se encontró la fila de encabezado (REGLAS/PRODUCTO/CLAVE1...) en REGLAS_OPERATIVAS.")
    return raw, header_row


def parse_reglas_operativas(xlsm_path):
    """Lee la hoja REGLAS_OPERATIVAS y construye:
      - reglas_raw: lista de dicts (cada fila tal cual, para mostrar como referencia en la UI)
      - params_default: dict con la estructura `params` por defecto, lista para edición en Streamlit
      - df_cap_default: DataFrame con la sub-tabla CAPACIDAD TINTORERIA
    """
    raw, header_row = find_reglas_operativas_table(xlsm_path)
    cols = [str(v).strip().upper() if pd.notna(v) else "" for v in raw.iloc[header_row].tolist()]
    body = raw.iloc[header_row + 1:].copy()
    body.columns = cols[:len(body.columns)]

    # localizar dónde empieza la sub-tabla CAPACIDAD TINTORERIA (columna REGLAS == 'CAPACIDAD TINTORERIA')
    if "REGLAS" not in body.columns:
        raise ValueError("La hoja REGLAS_OPERATIVAS no tiene columna REGLAS reconocible.")

    cap_mask = body["REGLAS"].astype(str).str.strip().str.upper() == "CAPACIDAD TINTORERIA"
    cap_rows = body[cap_mask]
    rules_rows = body[~cap_mask & body["REGLAS"].notna() & (body["REGLAS"].astype(str).str.strip() != "")]

    # ---------- Reglas con tabla (RESTRICCION_FAMILIA / COLOR / ANCHO) ----------
    restr_fam = {}
    restr_color = {}
    restr_ancho = {}
    reglas_raw = []

    obs_minimo_ancho = None
    obs_maximo_ancho = None
    obs_max_skus = None
    obs_split_minimo = None

    for _, r in rules_rows.iterrows():
        tag = str(r.get("REGLAS", "")).strip().upper()
        producto = str(r.get("PRODUCTO", "")).strip().upper() if pd.notna(r.get("PRODUCTO", None)) else ""
        c1 = _num(r.get("CLAVE1", None))
        c2 = _num(r.get("CLAVE2", None))
        c3 = _num(r.get("CLAVE3", None))
        c4 = _num(r.get("CLAVE4", None))
        obs = str(r.get("OBSERVACION", "")).strip() if pd.notna(r.get("OBSERVACION", None)) else ""

        reglas_raw.append({
            "REGLAS": tag, "PRODUCTO": producto,
            "CLAVE1": c1, "CLAVE2": c2, "CLAVE3": c3, "CLAVE4": c4,
            "OBSERVACION": obs
        })

        if tag == "RESTRICCION_FAMILIA" and producto:
            caps = [v for v in [c1, c2, c3, c4] if v is not None]
            if caps:
                restr_fam[producto] = caps

        elif tag == "RESTRICCION_COLOR" and producto:
            # aplica a TODOS -> se guarda bajo clave especial; la UI puede expandir a colores reales
            if c1 is not None:
                restr_color[producto] = c1

        elif tag == "RESTRICCION_ANCHO" and producto:
            if c1 is not None:
                restr_ancho[producto] = {"limite": c1, "prioridades": [c2] if c2 is not None else []}

        elif tag == "MINIMO ANCHO":
            obs_minimo_ancho = _first_number_in_text(obs, 1.0)

        elif tag == "MAXIMO ANCHO":
            obs_maximo_ancho = _first_number_in_text(obs, 6.0)

        elif tag == "MAXIMO SKUS":
            obs_max_skus = _first_number_in_text(obs, 6.0)

        elif tag == "SPLIT_MINIMO":
            obs_split_minimo = _first_number_in_text(obs, 500.0)

    # ---------- Sub-tabla CAPACIDAD TINTORERIA ----------
    # En la plantilla real, la fila de encabezado de esta sub-tabla está embebida
    # (columna PRODUCTO=CATEGORIA, CLAVE1=MINIMO, CLAVE2=MAXIMO, CLAVE3=CAPACIDAD, CLAVE4=MIX)
    cap_data_rows = cap_rows[cap_rows["PRODUCTO"].notna() & (cap_rows["PRODUCTO"].astype(str).str.strip().str.upper() != "CATEGORIA")]
    cap_records = []
    for _, r in cap_data_rows.iterrows():
        categoria = str(r.get("PRODUCTO", "")).strip()
        if not categoria:
            continue
        minimo = _num(r.get("CLAVE1", None))
        maximo = _num(r.get("CLAVE2", None))
        capacidad = _num(r.get("CLAVE3", None))
        mix = str(r.get("CLAVE4", "")).strip().upper() if pd.notna(r.get("CLAVE4", None)) else ""
        if minimo is None or maximo is None or capacidad is None or not mix:
            continue
        cap_records.append({
            "CATEGORIA": categoria, "MINIMO": minimo, "MAXIMO": maximo,
            "CAPACIDAD": capacidad, "MIX": mix
        })

    if not cap_records:
        # fallback a los valores conocidos de la plantilla, por si el parseo posicional falla
        cap_records = [
            {"CATEGORIA": "A-4000", "MINIMO": 3900, "MAXIMO": 4000, "CAPACIDAD": 660000, "MIX": "DYE"},
            {"CATEGORIA": "B-3300", "MINIMO": 3000, "MAXIMO": 3300, "CAPACIDAD": 555000, "MIX": "DYE"},
            {"CATEGORIA": "C-2600", "MINIMO": 2500, "MAXIMO": 2600, "CAPACIDAD": 1967000, "MIX": "DYE"},
            {"CATEGORIA": "D-2200", "MINIMO": 2000, "MAXIMO": 2200, "CAPACIDAD": 640000, "MIX": "DYE"},
            {"CATEGORIA": "E-1100", "MINIMO": 1000, "MAXIMO": 1100, "CAPACIDAD": 437000, "MIX": "DYE"},
            {"CATEGORIA": "F-2200", "MINIMO": 2000, "MAXIMO": 2200, "CAPACIDAD": 1212000, "MIX": "BLEACH"},
            {"CATEGORIA": "G-1100", "MINIMO": 1000, "MAXIMO": 1100, "CAPACIDAD": 75000, "MIX": "BLEACH"},
        ]

    df_cap_default = pd.DataFrame(cap_records)

    # ---------- Construcción de params por defecto ----------
    params_default = {
        "MIN_DIFF": obs_minimo_ancho if obs_minimo_ancho is not None else 1.0,
        "MAX_DIFF": obs_maximo_ancho if obs_maximo_ancho is not None else 6.0,
        "MIN_DIFF_BY_TIPO": {"JERSEY": 1.5, "FLEECE": 1.0, "OTRO": obs_minimo_ancho if obs_minimo_ancho is not None else 1.0},
        "MAX_DIFF_BY_TIPO": {"JERSEY": 6.0, "FLEECE": 7.0, "OTRO": obs_maximo_ancho if obs_maximo_ancho is not None else 6.0},
        "MAX_SKU": int(obs_max_skus) if obs_max_skus is not None else 6,
        "MAX_WIDTHS_BY_CAT": dict(DEFAULT_MAX_WIDTHS_BY_CAT),
        "MAX_WIDTHS_DEFAULT": 4,

        "SPLIT_MIN_LBS_DEFAULT": obs_split_minimo if obs_split_minimo is not None else 250.0,
        "SPLIT_MIN_LBS_ANCHO18": 250.0,
        "SCRAP_REMAINDER_BELOW_SPLIT_MIN": 0,

        "RESTRICCIONES_FAMILIA": restr_fam,
        "RESTRICCIONES_COLOR": restr_color,
        "RESTRICCIONES_ANCHO": restr_ancho,
        "REGLAS_ANCHOS_COMBINADOS": [],  # no viene en la nueva plantilla; queda vacío por defecto, editable

        "ALLOWED_PAIRS": list(DEFAULT_ALLOWED_PAIRS),
        "MIX_ALLOWED": set(DEFAULT_ALLOWED_PAIRS) | {(b, a) for a, b in DEFAULT_ALLOWED_PAIRS},

        "RULE_ORDER": "ANCHO18>COMBO_ANCHOS>COLOR_R>FAMILIA>DEFAULT",
        "PRIORITY_ORDER": "",

        "APPLY_RULES_BLEACH": 0,  # MIX-BLEACH-DYE: en BLEACH NO aplican RESTRICCION_FAMILIA/COLOR

        "UPGRADE_CATEGORIA": 1,
        "TRY_ALL_PRIORITIES": 1,

        "ANCHO18_ALLOW_SPILLOVER_2600": 0,
        "ANCHO18_ALLOWED_MAX_DYE": {2200.0, 1100.0},

        "BEAM_WIDTH": 12,
        "SMALL_SPLIT_REUSE_MIN": 100.0,
        "PRIORIZE_LARGE_CATEGORIES": 1,
        "W_FILL": 5.0,
        "W_CAP_LOSS": 3.0,
        "WIDTH_PREF_LIST": [4, 3, 2, 1],
        "W_WIDTH_PREF": 2.0,
        "W_1100_WIDTHS_STRICT": 10.0,

        "WIDTHS_TARGET_ORDER": "4>3>2>1",
        "REQUIRE_WIDTHS_STRICT": 0,
        "ALLOWED_MAXIMO_FOR_3_WIDTHS": {"DYE": {4000.0, 3300.0, 2600.0, 2200.0}, "BLEACH": set()},
        "ALLOWED_MAXIMO_FOR_4_WIDTHS": {"DYE": {4000.0, 3300.0, 2600.0, 2200.0}, "BLEACH": set()},

        # TIPO_TEJIDO (nuevo)
        "TIPO_TEJIDO_ENABLE": 1,
        "TIPO_TEJIDO_CATEGORIAS": ["A-4000", "B-3300"],
        "W_TIPO_TEJIDO_FLEECE": 4.0,

        # Activación individual de cada regla (toggles UI)
        "RULE_TOGGLES": {
            "RESTRICCION_FAMILIA": True,
            "RESTRICCION_COLOR": True,
            "RESTRICCION_ANCHO": True,
            "MIN_MAX_ANCHO": True,
            "MAX_CANTIDAD_ANCHOS": True,
            "MAX_SKUS": True,
            "COMBINACION_PRIORIDAD": True,
            "SPLIT_MINIMO": True,
            "MIX_BLEACH_DYE": True,
            "TIPO_TEJIDO": True,
            "PCT_CARGA": True,
        },
    }

    return reglas_raw, params_default, df_cap_default


def rule_order_options():
    """Todas las combinaciones posibles (permutaciones) de las 4 reglas + DEFAULT al final,
    para el selector de escenarios de ORDEN_REGLAS."""
    from itertools import permutations
    return [">".join(p) + ">DEFAULT" for p in permutations(RULE_TOKENS)]


# ============================================
# loteo_engine.py
# Motor puro (sin Streamlit / sin Colab) del algoritmo de loteo de tintorería NV2.
# Adaptado del script original "ANALYSYS_DATA_2026_Pref_anchos_.txt".
#
# CAMBIOS RESPECTO AL SCRIPT ORIGINAL (acordados con el usuario):
#  - Se ELIMINÓ el modo HUMANO % (OVERSHOOT_ENABLE / UNDERSHOOT_ENABLE / choose_take_humano).
#    Siempre se usa choose_take() (reparto exacto, sin overshoot/undershoot).
#  - SPLIT_MIN_LBS_DEFAULT default = 500 (antes 100).
#  - MAX_WIDTHS ahora es POR CATEGORÍA de tintorería (reemplaza el MAX_WIDTHS global único).
#  - COMBINACION_PRIORIDAD: matriz fija de bloques que se pueden mezclar:
#        VENCIDOS-VENCIDOS, VENCIDOS-AHEAD, AHEAD-AHEAD, AHEAD-AHEAD2, OTROS-AHEAD2
#    (editable desde la UI vía checkboxes, no se parsea texto libre en producción).
#  - TIPO_TEJIDO: nuevo parámetro de scoring. Si la categoría del lote es A-4000 o B-3300,
#    el tejido del SKU semilla (columna TIPO_TEJIDO) es FLEECE, y la FAMILIA de ese SKU NO
#    tiene RESTRICCION_FAMILIA activa, se suma un bono W_TIPO_TEJIDO_FLEECE al score del lote.
#  - %CARGA: nueva columna por fila en DATA (decimal, ej 0.8, 1.0). Reduce el MAXIMO efectivo
#    de capacidad de la categoría para ese lote (ej. 1100*0.80=880) sin cambiar la categoría
#    asignada (sigue siendo E-1100/G-1100). Se aplica usando el %CARGA del SKU semilla del lote.
# ============================================

import pandas as pd
import numpy as np
import re
from itertools import permutations

# ---------------------------- Utils ----------------------------
def norm_str(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def up(x):
    return norm_str(x).upper()

def clean_cols(cols):
    out = []
    for c in cols:
        c = "" if c is None else str(c)
        c = c.replace("\n", " ").replace("\r", " ")
        c = re.sub(r"\s+", " ", c).strip()
        out.append(c)
    return out

# ---------------------------- Readers ----------------------------
def find_header_row(xlsm_path, sheet_name, required_cols, search_rows=80):
    preview = pd.read_excel(xlsm_path, sheet_name=sheet_name, engine="openpyxl", header=None, nrows=search_rows)
    req = set(required_cols)
    for r in range(min(search_rows, len(preview))):
        row_vals = [norm_str(v) for v in preview.iloc[r].tolist()]
        row_set = set([v for v in row_vals if v])
        if req.issubset(row_set):
            return r
    return 0

def read_sheet_autoheader(xlsm_path, sheet_name, required_cols=None, default_header=0):
    hdr = find_header_row(xlsm_path, sheet_name, required_cols) if required_cols else default_header
    df = pd.read_excel(xlsm_path, sheet_name=sheet_name, engine="openpyxl", header=hdr)
    df.columns = clean_cols(df.columns)
    return df, hdr

# ---------------------------- Blocks & Widths ----------------------------
def prioridad_bloque(prio_text: str) -> str:
    """Normaliza prioridad sin absorber AHEAD/AHEAD2 dentro de DUE.

    Precedencia intencional:
      1) PAST DUE / VENCIDO -> VENCIDOS
      2) AHEAD2             -> AHEAD2
      3) AHEAD              -> AHEAD
      4) DUE                -> DUE
      5) resto              -> OTROS
    """
    p = re.sub(r"[\s_\-]+", " ", up(prio_text)).strip()
    if re.search(r"\bPAST\s+DUE\b|\bVENCID[OA]S?\b|\bOVERDUE\b", p):
        return "VENCIDOS"
    if re.search(r"\bAHEAD\s*2\b", p):
        return "AHEAD2"
    if re.search(r"\bAHEAD\b", p):
        return "AHEAD"
    if re.search(r"\bDUE\b", p):
        return "DUE"
    return "OTROS"

def can_mix_blocks(b1, b2, allowed_pairs):
    if b1 == b2:
        return True
    return (b1, b2) in allowed_pairs

def valid_width_group(widths, min_diff, max_diff, max_widths):
    w = [float(x) for x in widths if x is not None and not pd.isna(x) and float(x) != 0.0]
    uw = sorted(set(w))
    if len(uw) <= 1:
        return True
    if len(uw) > int(max_widths):
        return False
    for i in range(len(uw)):
        for j in range(i + 1, len(uw)):
            d = abs(uw[j] - uw[i])
            if d < min_diff or d > max_diff:
                return False
    return True

def get_row_widths(work, idx):
    widths = []
    for c in ["ANCHO.F.C", "ANCHO.F.M"]:
        if c in work.columns:
            v = work.at[idx, c]
            if pd.notna(v) and float(v) != 0.0:
                widths.append(float(v))
    return widths

# ---------------------------- Split chooser ----------------------------
def choose_take(rest, remaining, split_min_lbs, allow_scrap_residue=False, small_reuse_min=100.0):
    """No crea splits nuevos menores al mínimo; reutiliza saldos completos de 100 a 249 lb."""
    try:
        rest = float(rest)
        remaining = float(remaining)
        split_min_lbs = float(split_min_lbs)
        small_reuse_min = float(small_reuse_min)
    except Exception:
        return 0.0
    if rest <= 1e-9 or remaining <= 1e-9:
        return 0.0
    # Un saldo existente >=100 puede utilizarse completo para completar un lote.
    if rest <= remaining + 1e-9:
        return rest if rest + 1e-9 >= small_reuse_min else 0.0
    take = remaining
    # Una división nueva siempre debe respetar el mínimo operativo.
    if take + 1e-9 < split_min_lbs:
        return 0.0
    residue = rest - take
    if residue <= 1e-9 or residue + 1e-9 >= split_min_lbs:
        return take
    # Se permite conservar un remanente reutilizable entre 100 y split_min-1.
    if small_reuse_min - 1e-9 <= residue < split_min_lbs - 1e-9:
        return take
    return take if allow_scrap_residue else 0.0

# ---------------------------- Ranges builder ----------------------------
def build_ranges(df_cap):
    ranges = []
    for _, r in df_cap.iterrows():
        ranges.append({
            "CATEGORIA": norm_str(r["CATEGORIA"]),
            "MINIMO": float(r["MINIMO"]),
            "MAXIMO": float(r["MAXIMO"]),
            "CAPACIDAD": float(r["CAPACIDAD"]),
            "MIX": up(r["MIX"]),
            "RANGO_ID": f"CAP_{norm_str(r['CATEGORIA'])}_{up(r['MIX'])}_{float(r['MAXIMO']):.0f}"
        })
    return sorted(ranges, key=lambda x: x["MAXIMO"], reverse=True)

# ---------------------------- Generación de órdenes posibles para ORDEN_REGLAS ----------------------------
def all_rule_order_options():
    """Devuelve todas las combinaciones posibles (permutaciones) de las 4 reglas
    + DEFAULT siempre al final, para exponer en la UI como selector de escenarios."""
    base_rules = ["ANCHO18", "COMBO_ANCHOS", "COLOR_R", "FAMILIA"]
    opts = []
    for perm in permutations(base_rules):
        opts.append(">".join(perm) + ">DEFAULT")
    return opts

# ---------------------------- Load inputs (DATA + params ya parametrizados) ----------------------------
REQUIRED_DATA_COLS = ["LNK", "TELA.CUERPO", "COLOR", "PRIORIDAD", "ANCHO.F.C", "ANCHO.F.M", "TOTAL", "MIX", "CONSUMO_C"]

def load_data_sheet(xlsm_path):
    """Lee únicamente la hoja DATA con autodetección de fila de encabezado."""
    df_data, hdr_row = read_sheet_autoheader(xlsm_path, "DATA", required_cols=REQUIRED_DATA_COLS, default_header=0)
    miss = [c for c in REQUIRED_DATA_COLS if c not in df_data.columns]
    if miss:
        raise ValueError(f"DATA: faltan columnas obligatorias {miss}. Header detectado en fila {hdr_row+1}.")

    for c in ["ANCHO.F.C", "ANCHO.F.M", "TOTAL", "CONSUMO_C"]:
        df_data[c] = pd.to_numeric(df_data[c], errors="coerce").fillna(0.0)
    for c in ["LNK", "TELA.CUERPO", "COLOR", "PRIORIDAD", "MIX"]:
        df_data[c] = df_data[c].apply(norm_str)
    df_data["MIX"] = df_data["MIX"].apply(up)

    for opt_col, default in [("FAMILIA", ""), ("COLOR_R", ""), ("STYLE", "")]:
        if opt_col not in df_data.columns:
            df_data[opt_col] = default
        else:
            df_data[opt_col] = df_data[opt_col].apply(up)

    if "TONO" in df_data.columns:
        df_data["TONO"] = df_data["TONO"].apply(up)

    # TIPO_TEJIDO (nueva columna): FLEECE / JERSEY / OTRO
    if "TIPO_TEJIDO" not in df_data.columns:
        df_data["TIPO_TEJIDO"] = ""
    else:
        df_data["TIPO_TEJIDO"] = df_data["TIPO_TEJIDO"].apply(up)

    # %CARGA (nueva columna): decimal por fila, ej 0.7, 0.8, 1.0
    carga_col = None
    for cand in ["%CARGA", "PCT_CARGA", "PORCENTAJE_CARGA", "% CARGA"]:
        if cand in df_data.columns:
            carga_col = cand
            break
    if carga_col is None:
        df_data["PCT_CARGA"] = 1.0
    else:
        df_data["PCT_CARGA"] = pd.to_numeric(df_data[carga_col], errors="coerce").fillna(1.0)
        df_data.loc[(df_data["PCT_CARGA"] <= 0) | (df_data["PCT_CARGA"] > 1.0), "PCT_CARGA"] = 1.0

    return df_data, hdr_row

def build_cap_dataframe(cap_rows):
    """cap_rows: lista de dicts {CATEGORIA, MINIMO, MAXIMO, CAPACIDAD, MIX} (viene del data_editor de la UI)."""
    df_cap = pd.DataFrame(cap_rows)
    df_cap["CATEGORIA"] = df_cap["CATEGORIA"].apply(norm_str)
    df_cap["MIX"] = df_cap["MIX"].apply(up)
    for c in ["MINIMO", "MAXIMO", "CAPACIDAD"]:
        df_cap[c] = pd.to_numeric(df_cap[c], errors="coerce")
    if df_cap[["MINIMO", "MAXIMO", "CAPACIDAD"]].isna().any().any():
        raise ValueError("CAPACIDAD TINTORERIA: hay valores MINIMO/MAXIMO/CAPACIDAD inválidos o vacíos.")
    return df_cap

# ---------------------------- Priority helpers ----------------------------
def order_priorities(pris, params):
    pris = [float(x) for x in pris if x is not None]
    po_text = norm_str(params.get("PRIORITY_ORDER", ""))
    if po_text:
        plan = [p.strip() for p in po_text.split(">") if p.strip()]
        rank = {}
        for i, v in enumerate(plan):
            if re.match(r"^\d+(\.\d+)?$", v):
                rank[float(v)] = i
        return sorted(pris, key=lambda x: (rank.get(float(x), 10_000), float(x)))
    return sorted(pris)

def order_by_priorities(base_ranges, prioridades):
    used = set()
    out = []
    for cap in prioridades:
        match = [r for r in base_ranges if abs(float(r["MAXIMO"]) - float(cap)) < 1e-6]
        for r in match:
            if id(r) not in used:
                out.append(r); used.add(id(r))
    for r in base_ranges:
        if id(r) not in used:
            out.append(r)
    return out

# ---------------------------- Reorder rules (ANCHO18 / COMBO_ANCHOS / COLOR_R / FAMILIA) ----------------------------
def reorder_ranges_for_seed(ranges_mix, mixv, work, seed_idx, params):
    base = list(ranges_mix)
    rule_info = {
        "regla_aplicada": "NONE",
        "prioridades": [],
        "match_combo": False,
        "limite_ancho_style": None,
        "origen_prioridad": "MIX",
        "combo_target_width": None,
    }
    if up(mixv) not in ("DYE",) and int(params.get("APPLY_RULES_BLEACH", 0)) != 1:
        return base, rule_info

    fam = up(work.at[seed_idx, "FAMILIA"]) if "FAMILIA" in work.columns else ""
    color_r = up(work.at[seed_idx, "COLOR_R"]) if "COLOR_R" in work.columns else ""
    style = up(work.at[seed_idx, "STYLE"]) if "STYLE" in work.columns else ""

    def f2(x):
        try:
            return float(x)
        except Exception:
            return 0.0

    ancho_c = f2(work.at[seed_idx, "ANCHO.F.C"]) if "ANCHO.F.C" in work.columns else 0.0
    ancho_m = f2(work.at[seed_idx, "ANCHO.F.M"]) if "ANCHO.F.M" in work.columns else 0.0

    restr_fam = params.get("RESTRICCIONES_FAMILIA", {})
    restr_color = params.get("RESTRICCIONES_COLOR", {})
    restr_ancho = params.get("RESTRICCIONES_ANCHO", {})
    reglas_combo = params.get("REGLAS_ANCHOS_COMBINADOS", [])
    rule_order_cfg = norm_str(params.get("RULE_ORDER", ""))
    rule_order = [x.strip().upper() for x in rule_order_cfg.split(">") if x.strip()] or \
        ["ANCHO18", "COMBO_ANCHOS", "COLOR_R", "FAMILIA", "DEFAULT"]

    def ancho_activo_leq_lim(ac, am, lim):
        vals = []
        try:
            if ac is not None and not pd.isna(ac) and float(ac) > 0:
                vals.append(float(ac))
        except Exception:
            pass
        try:
            if am is not None and not pd.isna(am) and float(am) > 0:
                vals.append(float(am))
        except Exception:
            pass
        return (len(vals) > 0) and (min(vals) <= float(lim))

    def try_ancho18():
        if style in restr_ancho:
            lim = restr_ancho[style].get("limite", None)
            prioridades = order_priorities(restr_ancho[style].get("prioridades", []), params)
            if lim is not None and ancho_activo_leq_lim(ancho_c, ancho_m, lim) and len(prioridades) > 0:
                rule_info.update({
                    "regla_aplicada": "ANCHO18",
                    "prioridades": list(prioridades),
                    "limite_ancho_style": lim,
                    "origen_prioridad": "STYLE",
                })
                return order_by_priorities(base, prioridades)
        return None

    def try_combo():
        for regla in reglas_combo:
            a1, a2 = regla["a1"], regla["a2"]
            prioridades = order_priorities(regla["prioridades"], params)
            seed_matches = (abs(ancho_c - a1) < 1e-6 or abs(ancho_m - a1) < 1e-6 or
                            abs(ancho_c - a2) < 1e-6 or abs(ancho_m - a2) < 1e-6)
            if not seed_matches:
                continue
            objetivo = a2 if (abs(ancho_c - a1) < 1e-6 or abs(ancho_m - a1) < 1e-6) else a1
            existe_otro = False
            for idx in work.index:
                if idx == seed_idx:
                    continue
                if float(work.at[idx, "LBS_RESTANTES"]) <= 0:
                    continue
                ac = f2(work.at[idx, "ANCHO.F.C"]); am = f2(work.at[idx, "ANCHO.F.M"])
                if abs(ac - objetivo) < 1e-6 or abs(am - objetivo) < 1e-6:
                    existe_otro = True; break
            if existe_otro and len(prioridades) > 0:
                rule_info.update({
                    "regla_aplicada": "COMBO_ANCHOS",
                    "prioridades": list(prioridades),
                    "match_combo": True,
                    "origen_prioridad": "COMBO",
                    "combo_target_width": float(objetivo),
                })
                return order_by_priorities(base, prioridades)
        return None

    def try_color_r():
        if color_r in restr_color and restr_color[color_r]:
            p = float(restr_color[color_r])
            rule_info.update({
                "regla_aplicada": "COLOR_R",
                "prioridades": [p],
                "origen_prioridad": "COLOR"
            })
            return order_by_priorities(base, [p])
        return None

    def try_familia():
        if fam in restr_fam and len(restr_fam[fam]) > 0:
            prioridades = order_priorities(restr_fam[fam], params)
            rule_info.update({
                "regla_aplicada": "FAMILIA",
                "prioridades": list(prioridades),
                "origen_prioridad": "FAMILIA"
            })
            return order_by_priorities(base, prioridades)
        return None

    for token in rule_order:
        out = None
        if token == "ANCHO18":
            out = try_ancho18()
        elif token == "COMBO_ANCHOS":
            out = try_combo()
        elif token == "COLOR_R":
            out = try_color_r()
        elif token == "FAMILIA":
            out = try_familia()
        elif token == "DEFAULT":
            out = base
        if out is not None:
            return out, rule_info
    return base, rule_info

# ---------------------------- Priority matching for upgrades ----------------------------
def ranges_matching_priority(pri, ranges_try, tol=1e-6, allow_nearest_higher=True):
    pri = float(pri)
    exact = [r for r in ranges_try if abs(float(r["MAXIMO"]) - pri) <= tol]
    if exact:
        return exact
    if not allow_nearest_higher:
        return []
    higher = [r for r in ranges_try if float(r["MAXIMO"]) >= pri - tol]
    if higher:
        higher = sorted(higher, key=lambda r: (float(r["MAXIMO"]) - pri, -float(r["MAXIMO"])))
        return [higher[0]]
    return []

# ---------------------------- Scoring ----------------------------
def score_lote(lote_dict, resumen_rows, params, categoria=None, seed_row=None):
    if lote_dict is None:
        return -1e30
    W_FILL = params.get("W_FILL", 5.0)
    W_CAP_LOSS = params.get("W_CAP_LOSS", 3.0)
    W_WIDTH_PREF = params.get("W_WIDTH_PREF", 2.0)
    W_1100_STRICT = params.get("W_1100_WIDTHS_STRICT", 10.0)
    pref_list = params.get("WIDTH_PREF_LIST", [4, 3, 2, 1])

    total = float(lote_dict.get("TOTAL_LOTE", 0.0))
    maximo = float(lote_dict.get("MAXIMO", 1.0))
    fill = total / maximo if maximo > 1e-9 else 0.0
    cap_loss = max(0.0, (maximo - total) / maximo) if maximo > 1e-9 else 1.0

    anchos = set()
    for r in resumen_rows:
        for w in r.get("ANCHOS_ROW", []):
            if w is not None:
                anchos.add(float(w))
    widths_unique = len(anchos)

    try:
        rank = pref_list.index(widths_unique)
    except ValueError:
        rank = len(pref_list) + abs(widths_unique - pref_list[-1])
    width_pref_score = -float(rank)

    score = (W_FILL * fill) + (-W_CAP_LOSS * cap_loss) + (W_WIDTH_PREF * width_pref_score)
    if abs(maximo - 1100.0) < 1e-6:
        score -= W_1100_STRICT * max(0, widths_unique - 1)

    # --- NUEVO: TIPO_TEJIDO (preferencia FLEECE en categorías grandes) ---
    if int(params.get("TIPO_TEJIDO_ENABLE", 0)) == 1 and categoria is not None and seed_row is not None:
        cats_fleece = set(params.get("TIPO_TEJIDO_CATEGORIAS", ["A-4000", "B-3300"]))
        if categoria in cats_fleece:
            tejido = up(seed_row.get("TIPO_TEJIDO", ""))
            familia = up(seed_row.get("FAMILIA", ""))
            restr_fam = params.get("RESTRICCIONES_FAMILIA", {})
            familia_tiene_restriccion = familia in restr_fam and len(restr_fam.get(familia, [])) > 0
            if tejido == "FLEECE" and not familia_tiene_restriccion:
                score += float(params.get("W_TIPO_TEJIDO_FLEECE", 4.0))
    return score

# ---------------------------- Filtro por objetivo de # de anchos ----------------------------
def filter_ranges_for_width_target(ranges_try, mixv, width_target, params):
    mixu = str(mixv).strip().upper()
    allowed = None
    if width_target == 3:
        allowed_map = params.get("ALLOWED_MAXIMO_FOR_3_WIDTHS", {})
        allowed = allowed_map.get(mixu, set())
    elif width_target == 4:
        allowed_map = params.get("ALLOWED_MAXIMO_FOR_4_WIDTHS", {})
        allowed = allowed_map.get(mixu, set())
    if allowed and len(allowed) > 0:
        return [r for r in ranges_try if float(r["MAXIMO"]) in allowed]
    return list(ranges_try)

# ---------------------------- Intento de lote ----------------------------
def intentar_lote_para_rango(work, seed_idx, rango, capacity_used, params, rule_info,
                              require_two_widths=False, split_min_lbs=None,
                              min_unique_widths=None, max_unique_widths=None):
    tejido_seed = up(work.at[seed_idx, "TIPO_TEJIDO"]) if "TIPO_TEJIDO" in work.columns else "OTRO"
    if tejido_seed not in ("JERSEY", "FLEECE"):
        tejido_seed = "OTRO"
    min_diff = float(params.get("MIN_DIFF_BY_TIPO", {}).get(tejido_seed, params["MIN_DIFF"]))
    max_diff = float(params.get("MAX_DIFF_BY_TIPO", {}).get(tejido_seed, params["MAX_DIFF"]))
    if not bool(params.get("RULE_TOGGLES", {}).get("MIN_MAX_ANCHO", True)):
        min_diff, max_diff = 0.0, float("inf")
    max_sku = params["MAX_SKU"]
    allowed_pairs = params["MIX_ALLOWED"]

    # MAX_WIDTHS por categoría (reemplaza global)
    max_widths_by_cat = params.get("MAX_WIDTHS_BY_CAT", {})
    max_widths = max_widths_by_cat.get(rango["CATEGORIA"], params.get("MAX_WIDTHS_DEFAULT", 4))

    rid = rango["RANGO_ID"]
    cap_total = float(rango["CAPACIDAD"])
    cap_used = float(capacity_used.get(rid, 0.0))
    cap_left_global = max(0.0, cap_total - cap_used)
    if cap_left_global <= 0:
        return None

    # %CARGA: el MAXIMO efectivo del lote se reduce según el % de carga del SKU semilla
    pct_carga_seed = 1.0
    if "PCT_CARGA" in work.columns:
        try:
            pct_carga_seed = float(work.at[seed_idx, "PCT_CARGA"])
            if pct_carga_seed <= 0 or pct_carga_seed > 1.0:
                pct_carga_seed = 1.0
        except Exception:
            pct_carga_seed = 1.0

    max_allowed = min(float(rango["MAXIMO"]) * pct_carga_seed, cap_left_global)

    if float(work.at[seed_idx, "LBS_RESTANTES"]) <= 0:
        return None

    try:
        split_min_lbs = float(split_min_lbs if split_min_lbs is not None else params.get("SPLIT_MIN_LBS_DEFAULT", 250.0))
    except Exception:
        split_min_lbs = float(params.get("SPLIT_MIN_LBS_DEFAULT", 250.0))
    allow_scrap_residue = int(params.get("SCRAP_REMAINDER_BELOW_SPLIT_MIN", 0)) == 1

    lote_rows = []
    lote_lbs = 0.0
    lote_lnks = set()
    lote_blocks = []
    lote_widths = []

    def can_add_row(idx, lbs_to_add):
        if lbs_to_add <= 0:
            return False
        if "TONO" in work.columns:
            seed_tono = up(work.at[seed_idx, "TONO"]) if not pd.isna(work.at[seed_idx, "TONO"]) else ""
            row_tono = up(work.at[idx, "TONO"]) if not pd.isna(work.at[idx, "TONO"]) else ""
            if seed_tono != row_tono:
                return False

        lnk = work.at[idx, "LNK"]
        new_lnks = set(lote_lnks); new_lnks.add(lnk)
        if len(new_lnks) > max_sku:
            return False

        b = work.at[idx, "BLOQUE"]
        # Categorías grandes con semilla VENCIDOS:
        # primero se intenta lote vencido puro; si no alcanza, se habilita DUE y luego AHEAD.
        if (int(params.get("PRIORIDAD_GRANDES_ENABLE", 1)) == 1
                and rango["CATEGORIA"] in set(params.get("PRIORIDAD_GRANDES_CATEGORIAS", ["A-4000", "B-3300"]))
                and lote_blocks and lote_blocks[0] == "VENCIDOS"):
            fallback = list(params.get("PRIORIDAD_GRANDES_FALLBACK", ["DUE", "AHEAD"]))
            etapa = prioridad_grandes_etapa
            permitidos = {"VENCIDOS"} | set(fallback[:etapa])
            if b not in permitidos:
                return False
        for existing_b in lote_blocks:
            if not can_mix_blocks(existing_b, b, allowed_pairs):
                return False

        widths_candidate = list(lote_widths) + get_row_widths(work, idx)
        if not valid_width_group(widths_candidate, min_diff, max_diff, max_widths):
            return False

        if max_unique_widths is not None:
            uwc = sorted(set([float(w) for w in widths_candidate if w is not None and not pd.isna(w) and float(w) != 0.0]))
            if len(uwc) > int(max_unique_widths):
                return False

        if lote_lbs + lbs_to_add > max_allowed + 1e-9:
            return False
        return True

    seed_rest = float(work.at[seed_idx, "LBS_RESTANTES"])
    remaining = max_allowed - lote_lbs
    take = choose_take(seed_rest, remaining, split_min_lbs, allow_scrap_residue=allow_scrap_residue)

    if take <= 0 or not can_add_row(seed_idx, take):
        return None

    lote_rows.append((seed_idx, take, 0.0, 0.0))
    lote_lbs += take
    lote_lnks.add(work.at[seed_idx, "LNK"])
    lote_blocks.append(work.at[seed_idx, "BLOQUE"])
    lote_widths += get_row_widths(work, seed_idx)

    combo_target = rule_info.get("combo_target_width", None) if rule_info else None
    prioridad_grandes_etapa = 0

    while True:
        remaining = max_allowed - lote_lbs
        if remaining <= 1e-6:
            break
        best = None
        best_take = 0.0
        best_score = -1e30

        for idx in work.index:
            rest = float(work.at[idx, "LBS_RESTANTES"])
            if rest <= 0:
                continue
            if any(i == idx for i, *_ in lote_rows):
                continue
            take = choose_take(rest, remaining, split_min_lbs, allow_scrap_residue=allow_scrap_residue)
            if take <= 0:
                continue
            if not can_add_row(idx, take):
                continue

            new_total = lote_lbs + take
            widths_now = set([float(w) for w in lote_widths if w is not None and not pd.isna(w) and float(w) != 0.0])
            widths_add = set([float(w) for w in get_row_widths(work, idx) if w is not None and not pd.isna(w) and float(w) != 0.0])
            new_widths = widths_now.union(widths_add)
            adds_new_width = 1 if len(new_widths) > len(widths_now) else 0

            has_target = 0
            if combo_target is not None:
                for w in widths_add:
                    if abs(float(w) - float(combo_target)) < 1e-6:
                        has_target = 1
                        break

            score = new_total + has_target * 1e-3 + adds_new_width * 1e-4
            if score > best_score:
                best_score = score
                best = idx
                best_take = take

        if best is None:
            break

        lote_rows.append((best, best_take, 0.0, 0.0))
        lote_lbs += best_take
        lote_lnks.add(work.at[best, "LNK"])
        lote_blocks.append(work.at[best, "BLOQUE"])
        lote_widths += get_row_widths(work, best)

    if lote_lbs + 1e-9 < float(rango["MINIMO"]) * pct_carga_seed:
        return None

    if min_unique_widths is not None:
        min_required = int(min_unique_widths)
    elif require_two_widths:
        min_required = 2
    else:
        min_required = None

    if min_required is not None:
        uw = sorted(set([float(w) for w in lote_widths if w is not None and not pd.isna(w) and float(w) != 0.0]))
        if len(uw) < int(min_required):
            return None

    if max_unique_widths is not None:
        uw = sorted(set([float(w) for w in lote_widths if w is not None and not pd.isna(w) and float(w) != 0.0]))
        if len(uw) > int(max_unique_widths):
            return None

    return {
        "RANGO_ID": rango["RANGO_ID"],
        "CATEGORIA": rango["CATEGORIA"],
        "MIX": rango["MIX"],
        "MINIMO": float(rango["MINIMO"]),
        "MAXIMO": float(rango["MAXIMO"]),
        "TOTAL_LOTE": float(lote_lbs),
        "ROWS": lote_rows,
        "REQUIERE_2_ANCHOS": bool(require_two_widths),
        "PCT_CARGA_USADO": pct_carga_seed,
        "PRIORIDAD_GRANDES_ETAPA": prioridad_grandes_etapa,
    }

# ---------------------------- Loteo principal ----------------------------
def run_loteo(df_data, df_cap, params, progress_cb=None):
    """progress_cb(fraction:float, msg:str) opcional para barra de progreso en Streamlit."""
    ranges = build_ranges(df_cap)
    capacity_used = {r["RANGO_ID"]: 0.0 for r in ranges}

    data = df_data.copy()
    data["BLOQUE"] = data["PRIORIDAD"].apply(prioridad_bloque)
    data["LBS_RESTANTES"] = data["TOTAL"].astype(float)
    data["LBS_SCRAP"] = 0.0

    detalle = []
    resumen = []
    lote_id_global = 1

    # Secuencia de negocio: semilla vencida pura; luego DUE; después AHEAD.
    block_order = ["VENCIDOS", "DUE", "AHEAD", "AHEAD2", "OTROS"]

    group_keys = ["TELA.CUERPO", "MIX"]
    if "TONO" in data.columns:
        group_keys.insert(1, "TONO")
    else:
        group_keys.insert(1, "COLOR")

    groups = list(data.groupby(group_keys).groups.items())
    n_groups = max(1, len(groups))

    for gi, (keys, grp_idx) in enumerate(groups):
        if progress_cb:
            progress_cb(gi / n_groups, f"Procesando grupo {gi+1}/{n_groups}")

        work = data.loc[grp_idx].copy()
        if "TONO" in data.columns:
            tela, tono, mixv = keys[0], keys[1], keys[2]
            color = None
        else:
            tela, color, mixv = keys[0], keys[1], keys[2]
            tono = None

        ranges_mix = [r for r in ranges if r["MIX"] == mixv]
        blocked = set()

        while True:
            work["LBS_RESTANTES"] = pd.to_numeric(work["LBS_RESTANTES"], errors="coerce").fillna(0.0)
            if (work["LBS_RESTANTES"] > 0).sum() == 0:
                break
            made_any = False

            for b in block_order:
                if b in blocked:
                    continue
                cand = work[(work["BLOQUE"] == b) & (work["LBS_RESTANTES"] > 0)]
                if len(cand) == 0:
                    blocked.add(b)
                    continue

                beam_w = int(params.get("BEAM_WIDTH", 12))
                top_seeds = cand.sort_values("LBS_RESTANTES", ascending=False).head(beam_w).index.tolist()

                best_lote = None
                best_pack = None
                best_selection_key = None

                for seed_idx in top_seeds:
                    ranges_try, rule_info = reorder_ranges_for_seed(ranges_mix, mixv, work, seed_idx, params)

                    # ANCHO18 es una restricción dura de categoría.
                    if rule_info.get("regla_aplicada") == "ANCHO18" and up(mixv) == "DYE":
                        allowed = set(params.get("ANCHO18_ALLOWED_MAX_DYE", {2200.0, 1100.0}))
                        if int(params.get("ANCHO18_ALLOW_SPILLOVER_2600", 0)) == 1:
                            allowed.add(2600.0)
                        ranges_try = [r for r in ranges_try if float(r["MAXIMO"]) in allowed]

                    # Familia/color/combo definen un techo y permiten esa categoría hacia abajo.
                    pri_list = order_priorities(rule_info.get("prioridades", []), params)
                    if pri_list:
                        techo = max(float(x) for x in pri_list)
                        ranges_try = [r for r in ranges_try if float(r["MAXIMO"]) <= techo + 1e-6]

                    # Cantidad de anchos = máximo de MAX_WIDTHS_BY_CAT, nunca objetivo exacto.
                    candidate_ranges = sorted(ranges_try, key=lambda rr: (-float(rr["MAXIMO"]), str(rr["CATEGORIA"])))
                    best_seed = None
                    best_seed_key = None
                    best_seed_prio = None

                    for r in candidate_ranges:
                        if capacity_used[r["RANGO_ID"]] >= r["CAPACIDAD"] - 1e-6:
                            continue
                        split_min = (params.get("SPLIT_MIN_LBS_ANCHO18", 250)
                                     if rule_info.get("regla_aplicada") == "ANCHO18"
                                     else float(params.get("SPLIT_MIN_LBS_DEFAULT", 250.0)))
                        require_two = rule_info.get("regla_aplicada") == "COMBO_ANCHOS"
                        intento = intentar_lote_para_rango(
                            work, seed_idx, r, capacity_used, params, rule_info,
                            require_two_widths=require_two,
                            split_min_lbs=split_min,
                            min_unique_widths=None,
                            max_unique_widths=None,
                        )
                        if intento is None and require_two:
                            intento = intentar_lote_para_rango(
                                work, seed_idx, r, capacity_used, params, rule_info,
                                require_two_widths=False,
                                split_min_lbs=split_min,
                                min_unique_widths=None,
                                max_unique_widths=None,
                            )
                        if intento is None:
                            continue

                        maximo_cat = float(intento["MAXIMO"])
                        total_cat = float(intento["TOTAL_LOTE"])
                        fill_cat = total_cat / maximo_cat if maximo_cat > 1e-9 else 0.0
                        perdida_cat = max(0.0, maximo_cat - total_cat)
                        local_key = (maximo_cat, fill_cat, -perdida_cat)
                        if best_seed_key is None or local_key > best_seed_key:
                            best_seed_key = local_key
                            best_seed = intento
                            best_seed_prio = max(pri_list) if pri_list else None

                    if best_seed is None:
                        continue

                    resumen_rows = []
                    for idx, _lbs, *_ in best_seed["ROWS"]:
                        resumen_rows.append({
                            "LNK": work.at[idx, "LNK"],
                            "ANCHOS_ROW": get_row_widths(work, idx),
                        })
                    lote_for_score = {
                        "MAXIMO": float(best_seed["MAXIMO"]),
                        "TOTAL_LOTE": float(best_seed["TOTAL_LOTE"]),
                    }
                    seed_row_dict = work.loc[seed_idx].to_dict()
                    sc = score_lote(lote_for_score, resumen_rows, params,
                                    categoria=best_seed["CATEGORIA"], seed_row=seed_row_dict)
                    maximo_cat = float(best_seed["MAXIMO"])
                    total_cat = float(best_seed["TOTAL_LOTE"])
                    fill_cat = total_cat / maximo_cat if maximo_cat > 1e-9 else 0.0
                    perdida_cat = max(0.0, maximo_cat - total_cat)
                    selection_key = (maximo_cat, fill_cat, -perdida_cat, float(sc))
                    if best_selection_key is None or selection_key > best_selection_key:
                        best_selection_key = selection_key
                        best_lote = best_seed
                        best_pack = (best_seed, rule_info, best_seed_prio, float(sc))

                if best_lote is None:
                    blocked.add(b)
                    continue

                lote, rule_info, prioridad_obj, best_score = best_pack
                split_min = params.get("SPLIT_MIN_LBS_ANCHO18", 250) if rule_info.get("regla_aplicada") == "ANCHO18" else float(params.get("SPLIT_MIN_LBS_DEFAULT", 250.0))

                lote_id = f"L{lote_id_global:06d}"
                lote_id_global += 1

                lote_widths = []
                for idx, _lbs, *_ in lote["ROWS"]:
                    lote_widths += get_row_widths(work, idx)
                anchos_lote = sorted(set([float(w) for w in lote_widths if w is not None and not pd.isna(w) and float(w) != 0.0]))
                anchos_lote_str = str(anchos_lote)

                prioridad_final = float(lote["MAXIMO"])
                regla_aplicada_final = rule_info.get("regla_aplicada", "NONE")
                requiere_2_anchos_flag = False
                if regla_aplicada_final == "COMBO_ANCHOS":
                    requiere_2_anchos_flag = bool(lote.get("REQUIERE_2_ANCHOS", False)) and (len(anchos_lote) >= 2)
                    if not requiere_2_anchos_flag:
                        regla_aplicada_final = "COMBO_ANCHOS_FALLBACK"

                for idx, lbs_asig, over_extra, under_saved in lote["ROWS"]:
                    detalle.append({
                        "LOTE_ID": lote_id,
                        "ANCHOS_LOTE": anchos_lote_str,
                        "ANCHOS_CANTIDAD": len(anchos_lote),
                        "CATEGORIA": lote["CATEGORIA"],
                        "MIX": lote["MIX"],
                        "TELA.CUERPO": tela,
                        "COLOR": work.at[idx, "COLOR"],
                        "TONO": work.at[idx, "TONO"] if "TONO" in work.columns else "",
                        "LNK_PRIORIDAD": f"{work.at[idx,'LNK']}|{work.at[idx,'PRIORIDAD']}",
                        "LNK": work.at[idx, "LNK"],
                        "PRIORIDAD": work.at[idx, "PRIORIDAD"],
                        "BLOQUE": work.at[idx, "BLOQUE"],
                        "ANCHO.F.C": float(work.at[idx, "ANCHO.F.C"]),
                        "ANCHO.F.M": float(work.at[idx, "ANCHO.F.M"]),
                        "CONSUMO_C": float(work.at[idx, "CONSUMO_C"]),
                        "FAMILIA": work.at[idx, "FAMILIA"],
                        "COLOR_R": work.at[idx, "COLOR_R"],
                        "STYLE": work.at[idx, "STYLE"],
                        "TIPO_TEJIDO": work.at[idx, "TIPO_TEJIDO"] if "TIPO_TEJIDO" in work.columns else "",
                        "PLANTA_COSTURA": work.at[idx, "PLANTA_COSTURA"] if "PLANTA_COSTURA" in work.columns else "",
                        "CONSTRUCCION": work.at[idx, "CONSTRUCCION"] if "CONSTRUCCION" in work.columns else "",
                        "PCT_CARGA": work.at[idx, "PCT_CARGA"] if "PCT_CARGA" in work.columns else 1.0,
                        "LBS_ASIGNADAS": float(lbs_asig),
                        "LBS_EXTRA_SOBRE_ORDEN": float(max(0.0, over_extra)),
                        "APLICA_REGLA": regla_aplicada_final,
                        "PRIORIDAD_USADA": prioridad_final,
                        "PRIORIDAD_OBJETIVO": prioridad_obj,
                        "ORIGEN_PRIORIDAD": rule_info.get("origen_prioridad", "MIX"),
                        "MATCH_ANCHO": bool(rule_info.get("match_combo", False)),
                        "LIMITE_ANCHO_STYLE": rule_info.get("limite_ancho_style", None),
                        "UPGRADE_CATEGORIA": int(params.get("UPGRADE_CATEGORIA", 0)),
                        "SPLIT_MIN_USADO": float(split_min),
                        "REQUIERE_2_ANCHOS": bool(requiere_2_anchos_flag),
                        "DECISION_SCORE": float(best_score)
                    })

                    prev_rest = float(work.at[idx, "LBS_RESTANTES"])
                    new_rest = prev_rest - float(lbs_asig)
                    work.at[idx, "LBS_RESTANTES"] = max(0.0, new_rest)

                    if int(params.get("SCRAP_REMAINDER_BELOW_SPLIT_MIN", 0)) == 1:
                        rem = float(work.at[idx, "LBS_RESTANTES"])
                        if rem > 1e-9 and rem < 100.0 - 1e-9:
                            work.at[idx, "LBS_SCRAP"] = float(work.at[idx, "LBS_SCRAP"]) + rem
                            work.at[idx, "LBS_RESTANTES"] = 0.0

                det_lote = [d for d in detalle if d["LOTE_ID"] == lote_id]
                lnks = {d["LNK"] for d in det_lote}
                bloques = [d["BLOQUE"] for d in det_lote]
                orden_bloque_dom = {"VENCIDOS": 0, "DUE": 1, "AHEAD": 2, "AHEAD2": 3, "OTROS": 4}
                bloque_dom = min(set(bloques), key=lambda x: (-bloques.count(x), orden_bloque_dom.get(x, 999))) if bloques else ""

                resumen.append({
                    "LOTE_ID": lote_id,
                    "ANCHOS_LOTE": anchos_lote_str,
                    "ANCHOS_CANTIDAD": len(anchos_lote),
                    "CATEGORIA": lote["CATEGORIA"],
                    "MIX": lote["MIX"],
                    "TELA.CUERPO": tela,
                    "COLOR/TONO_KEY": tono if tono is not None else color,
                    "LBS_TOTAL": float(lote["TOTAL_LOTE"]),
                    "MIN_RANGO": float(lote["MINIMO"]),
                    "MAX_RANGO": float(lote["MAXIMO"]),
                    "CAPACIDAD_PERDIDA": float(lote["MAXIMO"] - lote["TOTAL_LOTE"]),
                    "SKU_DISTINTOS": len(lnks),
                    "ANCHOS_UNICOS": len(anchos_lote),
                    "BLOQUE_DOMINANTE": bloque_dom,
                    "BLOQUES_LOTE": ",".join(sorted(set(bloques))),
                    "FALLBACK_PRIORIDAD_GRANDE": int(lote.get("PRIORIDAD_GRANDES_ETAPA", 0)),
                    "REGLA_DOMINANTE": regla_aplicada_final,
                    "PRIORIDAD_FINAL": prioridad_final,
                    "PRIORIDAD_OBJETIVO": prioridad_obj,
                    "COMBO_ANCHOS": (regla_aplicada_final == "COMBO_ANCHOS"),
                    "STYLE_CRITICO": True if rule_info.get("regla_aplicada") == "ANCHO18" else False,
                    "CANT_REGLAS_APLICADAS": 0 if rule_info.get("regla_aplicada") == "NONE" else 1,
                    "UPGRADE_CATEGORIA": int(params.get("UPGRADE_CATEGORIA", 0)),
                    "PCT_CARGA_USADO": lote.get("PCT_CARGA_USADO", 1.0),
                })

                capacity_used[lote["RANGO_ID"]] += float(lote["TOTAL_LOTE"])
                blocked = set()
                made_any = True
                break

            if not made_any:
                break

        data.loc[work.index, "LBS_RESTANTES"] = work["LBS_RESTANTES"]
        data.loc[work.index, "LBS_SCRAP"] = work["LBS_SCRAP"]

    if progress_cb:
        progress_cb(1.0, "Loteo finalizado")

    exced_cols = ["LNK", "TELA.CUERPO", "COLOR", "MIX", "PRIORIDAD", "BLOQUE", "ANCHO.F.C", "ANCHO.F.M", "TOTAL", "LBS_RESTANTES", "LBS_SCRAP"]
    if "TONO" in data.columns:
        exced_cols.insert(3, "TONO")
    for extra_col in ["PLANTA_COSTURA", "CONSTRUCCION", "TIPO_TEJIDO"]:
        if extra_col in data.columns and extra_col not in exced_cols:
            exced_cols.append(extra_col)
    exced = data[data["LBS_RESTANTES"] > 1e-9][exced_cols].copy()

    df_detalle = pd.DataFrame(detalle)
    if len(df_detalle) > 0:
        df_detalle["DOCENAS"] = np.where(df_detalle["CONSUMO_C"] > 0, df_detalle["LBS_ASIGNADAS"] / df_detalle["CONSUMO_C"], np.nan)
    df_resumen = pd.DataFrame(resumen)

    df_param_out = pd.DataFrame([
        ["MIN_DIFF", params["MIN_DIFF"]],
        ["MAX_DIFF", params["MAX_DIFF"]],
        ["MIN_DIFF_BY_TIPO", str(params.get("MIN_DIFF_BY_TIPO", {}))],
        ["MAX_DIFF_BY_TIPO", str(params.get("MAX_DIFF_BY_TIPO", {}))],
        ["MAX_WIDTHS_BY_CAT", str(params.get("MAX_WIDTHS_BY_CAT", {}))],
        ["MAX_SKU", params["MAX_SKU"]],
        ["SPLIT_MIN_LBS_DEFAULT", params.get("SPLIT_MIN_LBS_DEFAULT", 250.0)],
        ["SPLIT_MIN_LBS_ANCHO18", params.get("SPLIT_MIN_LBS_ANCHO18", 250)],
        ["RULE_ORDER", params.get("RULE_ORDER", "")],
        ["PRIORITY_ORDER", params.get("PRIORITY_ORDER", "")],
        ["APPLY_RULES_BLEACH", params.get("APPLY_RULES_BLEACH", 0)],
        ["TRY_ALL_PRIORITIES", params.get("TRY_ALL_PRIORITIES", 1)],
        ["UPGRADE_CATEGORIA", params.get("UPGRADE_CATEGORIA", 0)],
        ["ANCHO18_ALLOW_SPILLOVER_2600", params.get("ANCHO18_ALLOW_SPILLOVER_2600", 0)],
        ["ANCHO18_ALLOWED_MAX_DYE", ",".join(sorted(str(int(x)) for x in params.get("ANCHO18_ALLOWED_MAX_DYE", {2200.0, 1100.0})))],
        ["SCRAP_REMAINDER_BELOW_SPLIT_MIN", params.get("SCRAP_REMAINDER_BELOW_SPLIT_MIN", 0)],
        ["BEAM_WIDTH", params.get("BEAM_WIDTH", 12)],
        ["W_FILL", params.get("W_FILL", 5.0)],
        ["W_CAP_LOSS", params.get("W_CAP_LOSS", 3.0)],
        ["WIDTH_PREF_LIST", ",".join(str(x) for x in params.get("WIDTH_PREF_LIST", [4, 3, 2, 1]))],
        ["W_WIDTH_PREF", params.get("W_WIDTH_PREF", 2.0)],
        ["W_1100_WIDTHS_STRICT", params.get("W_1100_WIDTHS_STRICT", 10.0)],
        ["WIDTHS_TARGET_ORDER", params.get("WIDTHS_TARGET_ORDER", "4>3>2>1")],
        ["REQUIRE_WIDTHS_STRICT", params.get("REQUIRE_WIDTHS_STRICT", 0)],
        ["ALLOWED_MAXIMO_FOR_3_WIDTHS_DYE", ",".join(str(int(x)) for x in sorted(params.get("ALLOWED_MAXIMO_FOR_3_WIDTHS", {}).get("DYE", set()), reverse=True))],
        ["ALLOWED_MAXIMO_FOR_4_WIDTHS_DYE", ",".join(str(int(x)) for x in sorted(params.get("ALLOWED_MAXIMO_FOR_4_WIDTHS", {}).get("DYE", set()), reverse=True))],
        ["ALLOWED_MAXIMO_FOR_3_WIDTHS_BLEACH", ",".join(str(int(x)) for x in sorted(params.get("ALLOWED_MAXIMO_FOR_3_WIDTHS", {}).get("BLEACH", set()), reverse=True))],
        ["ALLOWED_MAXIMO_FOR_4_WIDTHS_BLEACH", ",".join(str(int(x)) for x in sorted(params.get("ALLOWED_MAXIMO_FOR_4_WIDTHS", {}).get("BLEACH", set()), reverse=True))],
        ["TIPO_TEJIDO_ENABLE", params.get("TIPO_TEJIDO_ENABLE", 0)],
        ["W_TIPO_TEJIDO_FLEECE", params.get("W_TIPO_TEJIDO_FLEECE", 4.0)],
        ["PRIORIDAD_GRANDES_ENABLE", params.get("PRIORIDAD_GRANDES_ENABLE", 1)],
        ["PRIORIDAD_GRANDES_CATEGORIAS", ",".join(params.get("PRIORIDAD_GRANDES_CATEGORIAS", ["A-4000", "B-3300"]))],
        ["PRIORIDAD_GRANDES_FALLBACK", ",".join(params.get("PRIORIDAD_GRANDES_FALLBACK", ["DUE", "AHEAD"]))],
        ["COLOR_R_RESTRINGIDOS", ",".join(sorted(params.get("RESTRICCIONES_COLOR", {}).keys()))],
    ], columns=["PARAMETRO", "VALOR"])

    return df_detalle, df_resumen, exced, df_param_out

# ---------------------------- Reportes adicionales ----------------------------
def build_reports(df_data, df_cap, df_detalle, df_resumen):
    df_cap_simple = df_cap[["CATEGORIA", "MIX", "MINIMO", "MAXIMO", "CAPACIDAD"]].copy()
    if len(df_detalle) > 0:
        df_cat_asig = df_detalle.groupby(["CATEGORIA", "MIX"], as_index=False)["LBS_ASIGNADAS"].sum()
    else:
        df_cat_asig = pd.DataFrame({"CATEGORIA": [], "MIX": [], "LBS_ASIGNADAS": []})
    df_cap_cap = (df_cap_simple.merge(df_cat_asig, on=["CATEGORIA", "MIX"], how="left")
                  .fillna({"LBS_ASIGNADAS": 0.0}))
    df_cap_cap["DIFERENCIA"] = df_cap_cap["LBS_ASIGNADAS"] - df_cap_cap["CAPACIDAD"]
    df_cap_cap["FILL_RATE"] = np.where(df_cap_cap["CAPACIDAD"] > 0, df_cap_cap["LBS_ASIGNADAS"] / df_cap_cap["CAPACIDAD"], 0.0)
    df_cap_cap = df_cap_cap.sort_values(["MIX", "CATEGORIA"])

    df_base_blocks = df_data.copy()
    df_base_blocks["BLOQUE"] = df_base_blocks["PRIORIDAD"].apply(prioridad_bloque)
    df_prio_base = (df_base_blocks.groupby(["MIX", "BLOQUE"], as_index=False)["TOTAL"].sum()
                     .rename(columns={"TOTAL": "LBS_BASE"}))

    if len(df_detalle) > 0:
        df_prio_asig = df_detalle.groupby(["MIX", "BLOQUE"], as_index=False)["LBS_ASIGNADAS"].sum()
    else:
        df_prio_asig = pd.DataFrame({"MIX": [], "BLOQUE": [], "LBS_ASIGNADAS": []})

    df_prio_vs_asig = (df_prio_base.merge(df_prio_asig, on=["MIX", "BLOQUE"], how="left")
                        .fillna({"LBS_ASIGNADAS": 0.0}))
    df_prio_vs_asig["LBS_SIN_ASIGNAR"] = df_prio_vs_asig["LBS_BASE"] - df_prio_vs_asig["LBS_ASIGNADAS"]
    order_blocks = ["VENCIDOS", "DUE", "AHEAD", "AHEAD2", "OTROS"]
    df_prio_vs_asig["ORD"] = df_prio_vs_asig["BLOQUE"].apply(lambda x: order_blocks.index(x) if x in order_blocks else 999)
    df_prio_vs_asig = df_prio_vs_asig.sort_values(["MIX", "ORD"]).drop(columns=["ORD"])

    lnk_extra = [c for c in ["TELA.CUERPO", "COLOR", "TONO", "ANCHO.F.C", "ANCHO.F.M", "PRIORIDAD", "TIPO_TEJIDO", "PLANTA_COSTURA", "CONSTRUCCION"] if c in df_data.columns]
    agg_lnk = {"TOTAL": "sum"}
    for c in lnk_extra:
        agg_lnk[c] = (lambda x: " | ".join(dict.fromkeys(str(v) for v in x.dropna())) if x.dtype == object else x.dropna().iloc[0] if len(x.dropna()) else np.nan)
    df_lnk_base = (df_data.groupby(["MIX", "LNK"], as_index=False, dropna=False).agg(agg_lnk)
                   .rename(columns={"TOTAL": "LBS_BASE"}))
    if "LBS_SCRAP" in df_data.columns:
        df_lnk_scrap = (df_data.groupby(["MIX", "LNK"], as_index=False)["LBS_SCRAP"].sum())
    else:
        df_lnk_scrap = pd.DataFrame({"MIX": [], "LNK": [], "LBS_SCRAP": []})

    if len(df_detalle) > 0:
        df_lnk_asig = df_detalle.groupby(["MIX", "LNK"], as_index=False)["LBS_ASIGNADAS"].sum()
    else:
        df_lnk_asig = pd.DataFrame({"MIX": [], "LNK": [], "LBS_ASIGNADAS": []})

    df_lnk_comp = (df_lnk_base.merge(df_lnk_asig, on=["MIX", "LNK"], how="left")
                   .merge(df_lnk_scrap, on=["MIX", "LNK"], how="left")
                   .fillna({"LBS_ASIGNADAS": 0.0, "LBS_SCRAP": 0.0}))
    df_lnk_comp["BALANCE"] = df_lnk_comp["LBS_BASE"] - df_lnk_comp["LBS_ASIGNADAS"] - df_lnk_comp["LBS_SCRAP"]
    df_lnk_comp["ESTADO"] = np.where(df_lnk_comp["BALANCE"].abs() <= 1e-6,
                                      np.where(df_lnk_comp["LBS_SCRAP"] > 1e-6, "COMPLETO (SCRAP)", "COMPLETO"),
                                      "INCOMPLETO")
    df_lnk_comp = df_lnk_comp.sort_values(["MIX", "ESTADO", "BALANCE"], ascending=[True, True, False])

    def resumen_por_lote(df_det, filtro):
        if len(df_det) == 0:
            return pd.DataFrame()
        sub = df_det.query(filtro).copy()
        if len(sub) == 0:
            return pd.DataFrame()
        agg = (sub.groupby("LOTE_ID", as_index=False)
               .agg({
                   "ANCHOS_LOTE": "first",
                   "MIX": "first",
                   "TELA.CUERPO": "first",
                   "COLOR": "first",
                   "TONO": "first" if "TONO" in df_det.columns else (lambda x: ""),
                   "FAMILIA": "first",
                   "STYLE": "first",
                   "COLOR_R": "first",
                   "PRIORIDAD_USADA": "first",
                   "PRIORIDAD_OBJETIVO": "first",
                   "UPGRADE_CATEGORIA": "first",
                   "LBS_ASIGNADAS": "sum"
               }))
        return agg

    rep_ancho18 = resumen_por_lote(df_detalle, "APLICA_REGLA == 'ANCHO18'")
    rep_combo = resumen_por_lote(df_detalle, "APLICA_REGLA == 'COMBO_ANCHOS'")
    rep_color = resumen_por_lote(df_detalle, "APLICA_REGLA == 'COLOR_R'")
    rep_fam = resumen_por_lote(df_detalle, "APLICA_REGLA == 'FAMILIA'")

    if len(df_resumen) > 0:
        rep_maestro = df_resumen[[
            "LOTE_ID", "ANCHOS_LOTE", "MIX", "REGLA_DOMINANTE", "PRIORIDAD_FINAL", "PRIORIDAD_OBJETIVO",
            "COMBO_ANCHOS", "STYLE_CRITICO", "CANT_REGLAS_APLICADAS",
            "ANCHOS_UNICOS", "LBS_TOTAL", "CAPACIDAD_PERDIDA", "UPGRADE_CATEGORIA"
        ]].copy()
    else:
        rep_maestro = pd.DataFrame()

    if len(df_detalle) > 0:
        overs = (df_detalle.groupby(["MIX", "LNK"], as_index=False)
                 .agg({"LBS_EXTRA_SOBRE_ORDEN": "sum", "LBS_ASIGNADAS": "sum"}))
        tono_col = "TONO" if "TONO" in df_detalle.columns else "LNK"
        decision_log = df_detalle[[
            "LOTE_ID", "MIX", "TELA.CUERPO", "COLOR", tono_col,
            "FAMILIA", "STYLE", "COLOR_R",
            "CATEGORIA", "PRIORIDAD", "BLOQUE", "APLICA_REGLA", "PRIORIDAD_USADA", "PRIORIDAD_OBJETIVO",
            "LBS_ASIGNADAS", "LBS_EXTRA_SOBRE_ORDEN", "ANCHOS_LOTE", "DECISION_SCORE", "LNK"
        ]].copy().sort_values(["LOTE_ID", "LNK"])
    else:
        overs = pd.DataFrame({"MIX": [], "LNK": [], "LBS_EXTRA_SOBRE_ORDEN": [], "LBS_ASIGNADAS": []})
        decision_log = pd.DataFrame()

    # Diagnóstico explicable de no asignación. Es una causa probable basada en reglas activas,
    # no una afirmación absoluta del historial de búsqueda del optimizador.
    def causa_probable(row):
        if float(row.get("BALANCE", 0) or 0) <= 1e-6:
            return "ASIGNADO_COMPLETO"
        lbs = float(row.get("BALANCE", 0) or 0)
        ancho_vals = [float(row.get(c, 0) or 0) for c in ["ANCHO.F.C", "ANCHO.F.M"] if pd.notna(row.get(c, None)) and float(row.get(c, 0) or 0) > 0]
        tejido = up(row.get("TIPO_TEJIDO", ""))
        min_split = 0.0
        try:
            min_split = float(df_param_context.get("SPLIT_MIN_LBS_DEFAULT", 0))
        except Exception:
            min_split = 0.0
        if min_split and lbs < min_split:
            return "SALDO_MENOR_SPLIT_MINIMO"
        if not ancho_vals:
            return "ANCHO_NO_INFORMADO"
        return "SIN_COMBINACION_FACTIBLE_CON_REGLAS_ACTUALES"

    # Contexto mínimo para el diagnóstico, alimentado desde atributos agregados por run_loteo.
    df_param_context = getattr(df_data, "attrs", {}).get("LOTEO_PARAMS", {})
    if len(df_lnk_comp):
        df_lnk_comp["CAUSA_PROBABLE"] = df_lnk_comp.apply(causa_probable, axis=1)
        df_causas = (df_lnk_comp[df_lnk_comp["BALANCE"] > 1e-6]
                     .groupby("CAUSA_PROBABLE", as_index=False)
                     .agg(LNK_AFECTADOS=("LNK", "nunique"), LBS_NO_ASIGNADAS=("BALANCE", "sum"))
                     .sort_values("LBS_NO_ASIGNADAS", ascending=False))
        total_no_asig = float(df_causas["LBS_NO_ASIGNADAS"].sum()) if len(df_causas) else 0.0
        df_causas["PCT_LBS_NO_ASIGNADAS"] = np.where(total_no_asig > 0, df_causas["LBS_NO_ASIGNADAS"] / total_no_asig, 0.0)
    else:
        df_causas = pd.DataFrame(columns=["CAUSA_PROBABLE", "LNK_AFECTADOS", "LBS_NO_ASIGNADAS", "PCT_LBS_NO_ASIGNADAS"])

    # Auditoría objetiva de lotes ya creados contra reglas finales.
    audit_rows = []
    if len(df_resumen):
        for _, lr in df_resumen.iterrows():
            lote_id = lr.get("LOTE_ID", "")
            sub = df_detalle[df_detalle["LOTE_ID"] == lote_id] if len(df_detalle) else pd.DataFrame()
            fallas = []
            total_lote = float(lr.get("LBS_TOTAL", 0) or 0)
            minimo = float(lr.get("MIN_RANGO", 0) or 0)
            maximo = float(lr.get("MAX_RANGO", 0) or 0)
            if total_lote < minimo - 1e-6: fallas.append("MINIMO_CATEGORIA")
            if total_lote > maximo + 1e-6: fallas.append("MAXIMO_CATEGORIA")
            if len(sub):
                if sub["LNK"].nunique() > int(df_param_context.get("MAX_SKU", 9999)): fallas.append("MAX_SKUS")
                if "PCT_CARGA" in sub.columns:
                    pct = float(pd.to_numeric(sub["PCT_CARGA"], errors="coerce").fillna(1).iloc[0])
                    if total_lote > maximo * pct + 1e-6: fallas.append("PCT_CARGA")
                anchos = sorted(set(pd.to_numeric(pd.concat([sub.get("ANCHO.F.C", pd.Series(dtype=float)), sub.get("ANCHO.F.M", pd.Series(dtype=float))]), errors="coerce").dropna()))
                anchos = [float(x) for x in anchos if float(x) > 0]
                max_anchos = int(df_param_context.get("MAX_WIDTHS_BY_CAT", {}).get(lr.get("CATEGORIA", ""), df_param_context.get("MAX_WIDTHS_DEFAULT", 9999)))
                if len(anchos) > max_anchos: fallas.append("MAX_CANTIDAD_ANCHOS")
                tejido = up(sub["TIPO_TEJIDO"].iloc[0]) if "TIPO_TEJIDO" in sub.columns and len(sub) else "OTRO"
                if tejido not in ("JERSEY", "FLEECE"): tejido = "OTRO"
                min_d = float(df_param_context.get("MIN_DIFF_BY_TIPO", {}).get(tejido, df_param_context.get("MIN_DIFF", 0)))
                max_d = float(df_param_context.get("MAX_DIFF_BY_TIPO", {}).get(tejido, df_param_context.get("MAX_DIFF", 999)))
                for i, a in enumerate(anchos):
                    for b in anchos[i+1:]:
                        if abs(b-a) < min_d - 1e-9 or abs(b-a) > max_d + 1e-9: fallas.append("DIFERENCIA_ANCHOS")
                if "TONO" in sub.columns and sub["TONO"].fillna("").astype(str).nunique() > 1: fallas.append("TONO_MIXTO")
                bloques_lote = list(sub["BLOQUE"].dropna().astype(str).unique()) if "BLOQUE" in sub.columns else []
                allowed = df_param_context.get("MIX_ALLOWED", set())
                for i, b1 in enumerate(bloques_lote):
                    for b2 in bloques_lote[i+1:]:
                        if allowed and (b1, b2) not in allowed and (b2, b1) not in allowed:
                            fallas.append("COMBINACION_PRIORIDAD")
            audit_rows.append({"LOTE_ID": lote_id, "CATEGORIA": lr.get("CATEGORIA", ""), "MIX": lr.get("MIX", ""),
                               "LBS_TOTAL": total_lote, "ANCHOS_CANTIDAD": lr.get("ANCHOS_CANTIDAD", lr.get("ANCHOS_UNICOS", 0)),
                               "ESTADO_VALIDACION": "FALLA" if fallas else "OK",
                               "REGLAS_FALLIDAS": ", ".join(sorted(set(fallas)))})
    df_auditoria = pd.DataFrame(audit_rows)

    return {
        "CAPACIDAD_X_CATEG": df_cap_cap,
        "PRIORIDAD_VS_ASIG": df_prio_vs_asig,
        "LNK_COMPLETITUD": df_lnk_comp,
        "REGLA_STYLE_ANCHO18": rep_ancho18,
        "REGLA_COMBINACION_ANCHOS": rep_combo,
        "REGLA_COLOR_R": rep_color,
        "REGLA_FAMILIA": rep_fam,
        "REPORTE_REGLAS_MIX": rep_maestro,
        "OVERSHOOT_SUMMARY": overs,
        "DECISION_LOG": decision_log,
        "TOP_CAUSAS_NO_ASIGNACION": df_causas,
        "AUDITORIA_REGLAS_LOTES": df_auditoria
    }

# ---------------------------- Formatting / export ----------------------------
def format_workbook(path_xlsx, font_name="Arial", font_size=9):
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.chart import BarChart, DoughnutChart, Reference
    from openpyxl.chart.label import DataLabelList

    wb = load_workbook(path_xlsx)
    navy, blue, sky, orange, light, red, white = "082B54", "0072CE", "65B5E8", "F28E2B", "EAF0F6", "C00000", "FFFFFF"
    thin = Side(style="thin", color="D9E2F2")
    pct_tokens = ("PCT", "PORC", "FILL_RATE", "%")

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        ws.sheet_view.showGridLines = False
        ws.row_dimensions[1].height = 30
        headers = {str(c.value).strip(): c.column for c in ws[1] if c.value is not None}
        for cell in ws[1]:
            cell.fill = PatternFill("solid", fgColor=navy)
            cell.font = Font(name=font_name, size=10, bold=True, color=white)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(bottom=Side(style="medium", color=blue))
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.font = Font(name=font_name, size=font_size)
                cell.border = Border(bottom=thin)
                cell.alignment = Alignment(vertical="center")
                header = str(ws.cell(1, cell.column).value or "").upper()
                if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                    cell.number_format = "0.0%" if any(t in header for t in pct_tokens) else '#,##0;[Red]-#,##0;0'
        if ws.max_row >= 2 and ws.max_column >= 1:
            ws.conditional_formatting.add(f"A2:{get_column_letter(ws.max_column)}{ws.max_row}",
                CellIsRule(operator="lessThan", formula=["0"], font=Font(color=red)))
        for col_idx in range(1, ws.max_column + 1):
            values = [str(ws.cell(r, col_idx).value or "") for r in range(1, min(ws.max_row, 500) + 1)]
            width = min(max(10, max((len(v) for v in values), default=0) + 2), 45)
            ws.column_dimensions[get_column_letter(col_idx)].width = width
        for r in range(2, ws.max_row + 1):
            if r % 2 == 0:
                for c in range(1, ws.max_column + 1):
                    ws.cell(r, c).fill = PatternFill("solid", fgColor="F7FAFD")

        # Texto de COLOR con relleno aproximado cuando el nombre es reconocible.
        if ws.title == "DETALLE_LOTES" and "COLOR" in headers:
            cmap = {"BLACK":"000000", "WHITE":"FFFFFF", "RED":"E53935", "BLUE":"1E88E5", "NAVY":"082B54",
                    "GREEN":"43A047", "YELLOW":"FDD835", "ORANGE":"FB8C00", "PURPLE":"8E24AA", "PINK":"EC407A",
                    "GREY":"9E9E9E", "GRAY":"9E9E9E", "BROWN":"795548", "BEIGE":"D7CCC8", "NATURAL":"E8DFC8"}
            ci = headers["COLOR"]
            for r in range(2, ws.max_row + 1):
                txt = str(ws.cell(r, ci).value or "").upper()
                hit = next((hexv for name, hexv in cmap.items() if name in txt), None)
                if hit:
                    ws.cell(r, ci).fill = PatternFill("solid", fgColor=hit)
                    ws.cell(r, ci).font = Font(name=font_name, size=font_size, color="FFFFFF" if hit in {"000000","082B54","795548","8E24AA"} else "000000")

    # Hoja ejecutiva con gráficos vinculados a tablas del libro.
    if "DASHBOARD" in wb.sheetnames:
        del wb["DASHBOARD"]
    dash = wb.create_sheet("DASHBOARD", 0)
    dash.sheet_view.showGridLines = False
    dash["B2"] = "ELCATEX | Resumen ejecutivo de loteo"
    dash["B2"].font = Font(name=font_name, size=20, bold=True, color=navy)
    dash.column_dimensions["B"].width = 24
    for c in range(3, 12): dash.column_dimensions[get_column_letter(c)].width = 14

    if "CAPACIDAD_X_CATEG" in wb.sheetnames:
        src = wb["CAPACIDAD_X_CATEG"]
        h = {str(c.value): c.column for c in src[1]}
        if all(x in h for x in ["CATEGORIA", "CAPACIDAD", "LBS_ASIGNADAS"]):
            chart = BarChart(); chart.type = "col"; chart.style = 10; chart.title = "Capacidad vs asignado"
            chart.y_axis.title = "Libras"; chart.x_axis.title = "Categoría"; chart.height = 8; chart.width = 15
            chart.add_data(Reference(src, min_col=h["CAPACIDAD"], max_col=h["LBS_ASIGNADAS"], min_row=1, max_row=src.max_row), titles_from_data=True)
            chart.set_categories(Reference(src, min_col=h["CATEGORIA"], min_row=2, max_row=src.max_row))
            dash.add_chart(chart, "B5")
    if "TOP_CAUSAS_NO_ASIGNACION" in wb.sheetnames:
        src = wb["TOP_CAUSAS_NO_ASIGNACION"]
        h = {str(c.value): c.column for c in src[1]}
        if src.max_row > 1 and all(x in h for x in ["CAUSA_PROBABLE", "LBS_NO_ASIGNADAS"]):
            chart = BarChart(); chart.type = "bar"; chart.style = 10; chart.title = "Top causas probables de no asignación"
            chart.height = 8; chart.width = 15
            chart.add_data(Reference(src, min_col=h["LBS_NO_ASIGNADAS"], min_row=1, max_row=src.max_row), titles_from_data=True)
            chart.set_categories(Reference(src, min_col=h["CAUSA_PROBABLE"], min_row=2, max_row=src.max_row))
            dash.add_chart(chart, "J5")
    if "LNK_COMPLETITUD" in wb.sheetnames:
        src = wb["LNK_COMPLETITUD"]
        h = {str(c.value): c.column for c in src[1]}
        if "ESTADO" in h:
            counts = {}
            for r in range(2, src.max_row + 1): counts[str(src.cell(r, h["ESTADO"]).value)] = counts.get(str(src.cell(r, h["ESTADO"]).value), 0) + 1
            dash["B23"], dash["C23"] = "ESTADO", "LNK"
            for i, (k, v) in enumerate(counts.items(), 24): dash.cell(i,2,k); dash.cell(i,3,v)
            if counts:
                chart = DoughnutChart(); chart.title = "Completitud LNK"; chart.height = 7; chart.width = 10
                chart.add_data(Reference(dash, min_col=3, min_row=23, max_row=23+len(counts)), titles_from_data=True)
                chart.set_categories(Reference(dash, min_col=2, min_row=24, max_row=23+len(counts)))
                chart.dataLabels = DataLabelList(); chart.dataLabels.showPercent = True
                dash.add_chart(chart, "B25")
    wb.save(path_xlsx)

# ============================================
# app.py — App Streamlit de Loteo de Tintorería (NV2)
# ============================================
import io
import json
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go


st.set_page_config(page_title="Elcatex | Optimización de Loteo", page_icon="🧵", layout="wide", initial_sidebar_state="expanded")
LOGO_B64 = "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"
st.markdown("""<style>
html, body, [class*="css"] {font-family: Arial, Calibri, sans-serif;}
[data-testid="stAppViewContainer"] {background:#f5f8fc;}
[data-testid="stSidebar"] {background:#082b54;}
[data-testid="stSidebar"] * {color:#ffffff;}
.block-container {padding-top:1.2rem; max-width:1550px;}
.elc-header {background:linear-gradient(110deg,#05284d,#0068b9);padding:18px 24px;border-radius:14px;color:white;display:flex;align-items:center;gap:22px;box-shadow:0 8px 24px rgba(5,40,77,.18);}
.elc-header img {height:56px;background:white;border-radius:8px;padding:7px;}
.elc-header h1 {font-size:26px;margin:0;}
.elc-header p {margin:3px 0 0;color:#dbeeff;}
div[data-testid="stMetric"] {background:white;border-left:4px solid #0072ce;padding:14px;border-radius:10px;box-shadow:0 3px 12px rgba(0,0,0,.06);}
.stTabs [data-baseweb="tab-list"] {gap:6px;}
.stTabs [data-baseweb="tab"] {background:white;border-radius:8px 8px 0 0;padding:10px 14px;}
.stButton>button {border-radius:8px;font-weight:700;}
h1,h2,h3 {color:#082b54;}
</style>""", unsafe_allow_html=True)
st.markdown(f"""<div class="elc-header"><img src="data:image/png;base64,{LOGO_B64}"><div><h1>Centro de Optimización de Loteo</h1><p>Planeación Textil · Reglas operativas · Simulación y control ejecutivo</p></div></div>""", unsafe_allow_html=True)
st.caption("ELCATEX · Motor de decisión para lotes de tintorería")

# ---------------------------- Estado ----------------------------
for key, default in [
    ("df_data", None), ("df_fam", None), ("reglas_raw", None),
    ("params", None), ("df_cap", None), ("excel_path", None),
    ("resultado", None),
]:
    if key not in st.session_state:
        st.session_state[key] = default


def reset_params_from_excel():
    reglas_raw, params_default, df_cap_default = parse_reglas_operativas(st.session_state["excel_path"])
    st.session_state["reglas_raw"] = reglas_raw
    st.session_state["params"] = params_default
    st.session_state["df_cap"] = df_cap_default


# ---------------------------- 1. Carga de datos ----------------------------
st.header("1. Cargar Excel (DATA + REGLAS_OPERATIVAS + FAMILIA)")
uploaded = st.file_uploader("Sube el archivo .xlsx", type=["xlsx", "xlsm"])

if uploaded is not None:
    if st.session_state["excel_path"] != uploaded.name or st.session_state["df_data"] is None:
        # Guardar a disco temporalmente (pandas/openpyxl necesitan ruta o buffer)
        tmp_path = f"/tmp/{uploaded.name}"
        with open(tmp_path, "wb") as f:
            f.write(uploaded.getbuffer())
        st.session_state["excel_path"] = tmp_path
        try:
            df_data, hdr_row = load_data_sheet(tmp_path)
            st.session_state["df_data"] = df_data
            st.success(f"✅ Hoja DATA leída correctamente. Encabezado detectado en fila {hdr_row+1}. {len(df_data)} filas, {len(df_data.columns)} columnas.")
        except Exception as e:
            st.error(f"❌ Error leyendo DATA: {e}")
            st.session_state["df_data"] = None

        try:
            reset_params_from_excel()
            st.success("✅ REGLAS_OPERATIVAS parseado correctamente y cargado como configuración default.")
        except Exception as e:
            st.error(f"❌ Error parseando REGLAS_OPERATIVAS: {e}")

        try:
            df_fam = pd.read_excel(tmp_path, sheet_name="FAMILIA", engine="openpyxl")
            st.session_state["df_fam"] = df_fam
        except Exception:
            st.session_state["df_fam"] = pd.DataFrame()

if st.session_state["df_data"] is not None:
    st.subheader("Filtros del escenario")
    df_base = st.session_state["df_data"].copy()
    fc1, fc2, fc3 = st.columns(3)
    filter_specs = [("PLANTA_COSTURA", fc1), ("PRIORIDAD", fc2), ("CONSTRUCCION", fc3)]
    selected_filters = {}
    for col_name, container in filter_specs:
        with container:
            if col_name in df_base.columns:
                opts = sorted([str(x) for x in df_base[col_name].dropna().unique()])
                selected_filters[col_name] = st.multiselect(col_name.replace("_", " ").title(), opts, default=[])
            else:
                selected_filters[col_name] = []
                st.caption(f"{col_name}: columna no disponible")
    df_filtered = df_base.copy()
    for col_name, vals in selected_filters.items():
        if vals and col_name in df_filtered.columns:
            df_filtered = df_filtered[df_filtered[col_name].astype(str).isin(vals)]
    st.session_state["df_data_filtrada"] = df_filtered
    st.caption(f"Escenario activo: {len(df_filtered):,} de {len(df_base):,} filas · {df_filtered['TOTAL'].sum():,.0f} lbs")
    with st.expander("👀 Preview de DATA filtrada", expanded=False):
        st.dataframe(st.session_state["df_data_filtrada"].head(50), use_container_width=True)
        faltantes = [c for c in ["TONO", "TIPO_TEJIDO", "PCT_CARGA"] if c not in st.session_state["df_data"].columns]
        if faltantes:
            st.warning(f"⚠️ Columnas opcionales no encontradas (se usará default): {faltantes}")

if st.session_state["params"] is None:
    st.info("Sube un archivo Excel para continuar.")
    st.stop()

params = st.session_state["params"]
df_cap = st.session_state["df_cap"]

# ---------------------------- 2. Panel de configuración ----------------------------
st.header("2. Configuración")

col_reset, col_profile_exp, col_profile_imp = st.columns(3)
with col_reset:
    if st.button("🔄 Restaurar valores del Excel"):
        reset_params_from_excel()
        st.rerun()
with col_profile_exp:
    profile_json = json.dumps({k: (list(v) if isinstance(v, set) else v) for k, v in params.items()}, default=str, indent=2)
    st.download_button("💾 Exportar perfil (JSON)", data=profile_json, file_name="perfil_loteo.json", mime="application/json")
with col_profile_imp:
    profile_file = st.file_uploader("📂 Cargar perfil (JSON)", type=["json"], key="profile_uploader")
    if profile_file is not None:
        loaded = json.load(profile_file)
        st.session_state["params"].update(loaded)
        st.success("Perfil cargado. Revisa los valores abajo.")
        params = st.session_state["params"]

tabs = st.tabs([
    "RESTRICCION_FAMILIA", "RESTRICCION_COLOR", "RESTRICCION_ANCHO",
    "ANCHOS (MIN/MAX/CANTIDAD)", "MAX_SKUS", "COMBINACION_PRIORIDAD",
    "SPLIT_MINIMO", "MIX-BLEACH-DYE", "TIPO_TEJIDO", "%CARGA",
    "ORDEN_REGLAS", "CAPACIDAD TINTORERIA", "Avanzado"
])

# --- RESTRICCION_FAMILIA ---
with tabs[0]:
    st.markdown("La restricción de familia aplica a las familias seleccionadas y se hará en las categorías seleccionadas (MAXIMO) y hacia abajo.")
    on = st.checkbox("Activar RESTRICCION_FAMILIA", value=params["RULE_TOGGLES"]["RESTRICCION_FAMILIA"], key="t_fam")
    params["RULE_TOGGLES"]["RESTRICCION_FAMILIA"] = on
    fam_rows = [{"FAMILIA": k, "MAXIMOS_PERMITIDOS": ",".join(str(int(x)) for x in v)} for k, v in params["RESTRICCIONES_FAMILIA"].items()]
    df_fam_edit = st.data_editor(pd.DataFrame(fam_rows), num_rows="dynamic", use_container_width=True, key="ed_fam")
    new_restr_fam = {}
    for _, r in df_fam_edit.iterrows():
        fam = str(r.get("FAMILIA", "")).strip().upper()
        if not fam:
            continue
        try:
            vals = [float(x.strip()) for x in str(r.get("MAXIMOS_PERMITIDOS", "")).split(",") if x.strip()]
        except Exception:
            vals = []
        if vals:
            new_restr_fam[fam] = vals
    params["RESTRICCIONES_FAMILIA"] = new_restr_fam if on else {}

# --- RESTRICCION_COLOR ---
with tabs[1]:
    st.markdown("La restricción de color aplica a todos los estilos (TODOS), en la categoría seleccionada (MAXIMO) y hacia abajo.")
    on = st.checkbox("Activar RESTRICCION_COLOR", value=params["RULE_TOGGLES"]["RESTRICCION_COLOR"], key="t_color")
    params["RULE_TOGGLES"]["RESTRICCION_COLOR"] = on
    default_cap = params["RESTRICCIONES_COLOR"].get("RESTRICCION", params["RESTRICCIONES_COLOR"].get("TODOS", 2600.0))
    cap_val = st.number_input("MAXIMO para registros marcados como RESTRICCION en COLOR_R", value=float(default_cap), step=100.0, key="color_cap")
    if on and st.session_state["df_data"] is not None and "COLOR_R" in st.session_state["df_data"].columns:
        # COLOR_R es un marcador operativo. Solo valores que indiquen RESTRICCION
        # deben limitarse a 2600 y hacia abajo; NORMAL y demás valores quedan libres.
        marcadores = sorted({
            up(v) for v in st.session_state["df_data"]["COLOR_R"].dropna().unique()
            if "RESTRICC" in up(v)
        })
        params["RESTRICCIONES_COLOR"] = {m: cap_val for m in marcadores}
        if marcadores:
            st.caption(f"La regla se aplicará únicamente a COLOR_R marcado como restricción: {marcadores}")
        else:
            st.caption("No se encontraron registros COLOR_R marcados como RESTRICCION; la regla no limitará categorías.")
    else:
        params["RESTRICCIONES_COLOR"] = {}

# --- RESTRICCION_ANCHO ---
with tabs[2]:
    st.markdown("Por estilo (STYLE): límite de ancho (pulgadas) y categoría (MAXIMO) asociada, hacia abajo.")
    on = st.checkbox("Activar RESTRICCION_ANCHO", value=params["RULE_TOGGLES"]["RESTRICCION_ANCHO"], key="t_ancho")
    params["RULE_TOGGLES"]["RESTRICCION_ANCHO"] = on
    ancho_rows = [{"STYLE": k, "LIMITE_ANCHO": v["limite"], "MAXIMO_CATEGORIA": ",".join(str(int(x)) for x in v["prioridades"])} for k, v in params["RESTRICCIONES_ANCHO"].items()]
    df_ancho_edit = st.data_editor(pd.DataFrame(ancho_rows), num_rows="dynamic", use_container_width=True, key="ed_ancho")
    new_restr_ancho = {}
    for _, r in df_ancho_edit.iterrows():
        style = str(r.get("STYLE", "")).strip().upper()
        if not style:
            continue
        try:
            lim = float(r.get("LIMITE_ANCHO"))
        except Exception:
            continue
        try:
            caps = [float(x.strip()) for x in str(r.get("MAXIMO_CATEGORIA", "")).split(",") if x.strip()]
        except Exception:
            caps = []
        new_restr_ancho[style] = {"limite": lim, "prioridades": caps}
    params["RESTRICCIONES_ANCHO"] = new_restr_ancho if on else {}

# --- MINIMO/MAXIMO/CANTIDAD ANCHO ---
with tabs[3]:
    on = st.checkbox("Activar reglas de MIN/MAX/CANTIDAD de anchos", value=params["RULE_TOGGLES"]["MIN_MAX_ANCHO"], key="t_minmax")
    params["RULE_TOGGLES"]["MIN_MAX_ANCHO"] = on
    st.markdown("**Tolerancia de diferencia de anchos por tipo de tejido**")
    rows_diff = []
    for tejido in ["JERSEY", "FLEECE", "OTRO"]:
        rows_diff.append({"TIPO_TEJIDO": tejido, "MIN_DIFF": params.get("MIN_DIFF_BY_TIPO", {}).get(tejido, params["MIN_DIFF"]), "MAX_DIFF": params.get("MAX_DIFF_BY_TIPO", {}).get(tejido, params["MAX_DIFF"])})
    df_diff = st.data_editor(pd.DataFrame(rows_diff), hide_index=True, use_container_width=True, key="ed_diff_tipo")
    params["MIN_DIFF_BY_TIPO"] = {str(r["TIPO_TEJIDO"]).upper(): float(r["MIN_DIFF"]) for _, r in df_diff.iterrows()}
    params["MAX_DIFF_BY_TIPO"] = {str(r["TIPO_TEJIDO"]).upper(): float(r["MAX_DIFF"]) for _, r in df_diff.iterrows()}
    params["MIN_DIFF"] = params["MIN_DIFF_BY_TIPO"].get("OTRO", params["MIN_DIFF"])
    params["MAX_DIFF"] = params["MAX_DIFF_BY_TIPO"].get("OTRO", params["MAX_DIFF"])

    st.markdown("**MAXIMO CANTIDAD ANCHOS por categoría de tintorería** (reemplaza el máximo global):")
    params["RULE_TOGGLES"]["MAX_CANTIDAD_ANCHOS"] = st.checkbox("Activar MAXIMO CANTIDAD ANCHOS por categoría", value=params["RULE_TOGGLES"]["MAX_CANTIDAD_ANCHOS"], key="t_maxcant")
    cat_widths_rows = [{"CATEGORIA": k, "MAX_ANCHOS": v} for k, v in params["MAX_WIDTHS_BY_CAT"].items()]
    df_catw = st.data_editor(pd.DataFrame(cat_widths_rows), use_container_width=True, key="ed_catw")
    if params["RULE_TOGGLES"]["MAX_CANTIDAD_ANCHOS"]:
        params["MAX_WIDTHS_BY_CAT"] = {str(r["CATEGORIA"]): int(r["MAX_ANCHOS"]) for _, r in df_catw.iterrows() if str(r["CATEGORIA"]).strip()}
    else:
        params["MAX_WIDTHS_BY_CAT"] = {k: 6 for k in params["MAX_WIDTHS_BY_CAT"]}  # sin restricción real

# --- MAX SKUS ---
with tabs[4]:
    on = st.checkbox("Activar MAXIMO SKUS", value=params["RULE_TOGGLES"]["MAX_SKUS"], key="t_sku")
    params["RULE_TOGGLES"]["MAX_SKUS"] = on
    val = st.number_input("Máximo de SKUs (LNK) distintos por lote", value=int(params["MAX_SKU"]), min_value=1, step=1)
    params["MAX_SKU"] = val if on else 9999

# --- COMBINACION_PRIORIDAD ---
with tabs[5]:
    st.markdown("Matriz de bloques que pueden mezclarse en un mismo lote (PAST DUE+DUE=VENCIDOS, AHEAD, AHEAD2, OTROS):")
    on = st.checkbox("Activar COMBINACION_PRIORIDAD", value=params["RULE_TOGGLES"]["COMBINACION_PRIORIDAD"], key="t_combo")
    params["RULE_TOGGLES"]["COMBINACION_PRIORIDAD"] = on
    blocks = ["VENCIDOS", "DUE", "AHEAD", "AHEAD2", "OTROS"]
    default_pairs = set(params["ALLOWED_PAIRS"])
    selected_pairs = []
    st.caption("Marca los pares que se pueden mezclar (la diagonal, mismo bloque, siempre se permite).")
    cols = st.columns(len(blocks))
    pair_state = {}
    for i, b1 in enumerate(blocks):
        for j, b2 in enumerate(blocks):
            if j < i:
                continue
            key = f"pair_{b1}_{b2}"
            default_checked = (b1, b2) in default_pairs or (b2, b1) in default_pairs or b1 == b2
            checked = st.checkbox(f"{b1} ↔ {b2}", value=default_checked, key=key)
            if checked:
                selected_pairs.append((b1, b2))
    if on:
        params["ALLOWED_PAIRS"] = selected_pairs
        params["MIX_ALLOWED"] = set(selected_pairs) | {(b, a) for a, b in selected_pairs}
    else:
        all_pairs = [(a, b) for a in blocks for b in blocks]
        params["MIX_ALLOWED"] = set(all_pairs)

# --- SPLIT_MINIMO ---
with tabs[6]:
    on = st.checkbox("Activar SPLIT_MINIMO", value=params["RULE_TOGGLES"]["SPLIT_MINIMO"], key="t_split")
    params["RULE_TOGGLES"]["SPLIT_MINIMO"] = on
    val = st.number_input("Split mínimo (lbs) — evita splits más pequeños", value=float(params["SPLIT_MIN_LBS_DEFAULT"]), step=50.0)
    params["SPLIT_MIN_LBS_DEFAULT"] = val if on else 0.0
    params["SPLIT_MIN_LBS_ANCHO18"] = st.number_input("Split mínimo para regla ANCHO18 (lbs)", value=float(params["SPLIT_MIN_LBS_ANCHO18"]), step=50.0)

# --- MIX-BLEACH-DYE ---
with tabs[7]:
    st.markdown("BLEACH solo tiene categorías F-2200 y G-1100. Por defecto, en BLEACH NO aplican RESTRICCION_FAMILIA ni RESTRICCION_COLOR.")
    on = st.checkbox("¿Aplicar RESTRICCION_FAMILIA / RESTRICCION_COLOR también en BLEACH?", value=bool(params["APPLY_RULES_BLEACH"]), key="t_bleach")
    params["RULE_TOGGLES"]["MIX_BLEACH_DYE"] = True
    params["APPLY_RULES_BLEACH"] = 1 if on else 0

# --- TIPO_TEJIDO ---
with tabs[8]:
    st.markdown("En categorías grandes (A-4000, B-3300) preferir tejido FLEECE en el scoring, EXCEPTO si la familia tiene RESTRICCION_FAMILIA activa.")
    on = st.checkbox("Activar preferencia TIPO_TEJIDO (FLEECE)", value=params["RULE_TOGGLES"]["TIPO_TEJIDO"], key="t_tejido")
    params["RULE_TOGGLES"]["TIPO_TEJIDO"] = on
    params["TIPO_TEJIDO_ENABLE"] = 1 if on else 0
    params["W_TIPO_TEJIDO_FLEECE"] = st.number_input("Peso del bono FLEECE en el score del lote (W_TIPO_TEJIDO_FLEECE)", value=float(params["W_TIPO_TEJIDO_FLEECE"]), step=0.5)
    cats_sel = st.multiselect("Categorías donde aplica la preferencia FLEECE", options=list(params["MAX_WIDTHS_BY_CAT"].keys()), default=params["TIPO_TEJIDO_CATEGORIAS"])
    params["TIPO_TEJIDO_CATEGORIAS"] = cats_sel
    if st.session_state["df_data"] is not None and "TIPO_TEJIDO" not in st.session_state["df_data"].columns:
        st.warning("⚠️ No se encontró la columna TIPO_TEJIDO en DATA. Esta regla no tendrá efecto hasta que la columna exista.")

# --- %CARGA ---
with tabs[9]:
    st.markdown("Columna por fila en DATA con el % de carga (decimal, ej. 0.7, 0.8, 1.0). Reduce el MAXIMO efectivo de capacidad de la categoría para ese lote sin cambiar la categoría asignada.")
    on = st.checkbox("Activar %CARGA", value=params["RULE_TOGGLES"]["PCT_CARGA"], key="t_carga")
    params["RULE_TOGGLES"]["PCT_CARGA"] = on
    if not on and st.session_state["df_data"] is not None:
        st.session_state["df_data"]["PCT_CARGA"] = 1.0
    if st.session_state["df_data"] is not None and "PCT_CARGA" in st.session_state["df_data"].columns:
        st.caption("Distribución de % de carga encontrada en DATA:")
        st.dataframe(st.session_state["df_data"]["PCT_CARGA"].value_counts().rename("conteo"), use_container_width=True)

# --- ORDEN_REGLAS ---
with tabs[10]:
    st.markdown("Orden de aplicación de las reglas (ANCHO18 > COMBO_ANCHOS > COLOR_R > FAMILIA > DEFAULT). Selecciona un escenario; puedes correr varios y comparar.")
    options = rule_order_options()
    current = params.get("RULE_ORDER", options[0])
    if current not in options:
        options = [current] + options
    idx = options.index(current) if current in options else 0
    chosen = st.selectbox("Orden de reglas (escenario)", options=options, index=idx)
    params["RULE_ORDER"] = chosen
    st.caption("Tip: guarda distintos perfiles (JSON) con cada orden para comparar escenarios y luego ejecutar el loteo varias veces.")

# --- CAPACIDAD TINTORERIA ---
with tabs[11]:
    st.markdown("Tabla editable de capacidades por categoría de tintorería (equivalente a la hoja CAPACIDADES_TINTO original).")
    df_cap_edit = st.data_editor(df_cap, num_rows="dynamic", use_container_width=True, key="ed_cap")
    st.session_state["df_cap"] = df_cap_edit

# --- Avanzado ---
with tabs[12]:
    st.markdown("Parámetros avanzados del motor original (no presentes en REGLAS_OPERATIVAS, mantienen los defaults del script).")
    c1, c2, c3 = st.columns(3)
    with c1:
        params["BEAM_WIDTH"] = st.number_input("BEAM_WIDTH", value=int(params["BEAM_WIDTH"]), min_value=1, step=1)
        params["W_FILL"] = st.number_input("W_FILL", value=float(params["W_FILL"]), step=0.5)
        params["W_CAP_LOSS"] = st.number_input("W_CAP_LOSS", value=float(params["W_CAP_LOSS"]), step=0.5)
    with c2:
        params["W_WIDTH_PREF"] = st.number_input("W_WIDTH_PREF", value=float(params["W_WIDTH_PREF"]), step=0.5)
        params["W_1100_WIDTHS_STRICT"] = st.number_input("W_1100_WIDTHS_STRICT", value=float(params["W_1100_WIDTHS_STRICT"]), step=0.5)
        params["UPGRADE_CATEGORIA"] = 1 if st.checkbox("UPGRADE_CATEGORIA", value=bool(params["UPGRADE_CATEGORIA"])) else 0
    with c3:
        params["TRY_ALL_PRIORITIES"] = 1 if st.checkbox("TRY_ALL_PRIORITIES", value=bool(params["TRY_ALL_PRIORITIES"])) else 0
        params["REQUIRE_WIDTHS_STRICT"] = 1 if st.checkbox("REQUIRE_WIDTHS_STRICT", value=bool(params["REQUIRE_WIDTHS_STRICT"])) else 0
        params["WIDTHS_TARGET_ORDER"] = st.text_input("WIDTHS_TARGET_ORDER", value=params["WIDTHS_TARGET_ORDER"])

    wpl = st.text_input("WIDTH_PREF_LIST (coma-separado)", value=",".join(str(x) for x in params["WIDTH_PREF_LIST"]))
    try:
        params["WIDTH_PREF_LIST"] = [int(x.strip()) for x in wpl.split(",") if x.strip()]
    except Exception:
        pass

    st.markdown("**Filtro de categorías grandes para objetivo de 3/4 anchos (por MIX):**")
    c1, c2 = st.columns(2)
    with c1:
        dye3 = st.text_input("ALLOWED_MAXIMO_FOR_3_WIDTHS_DYE", value=",".join(str(int(x)) for x in params["ALLOWED_MAXIMO_FOR_3_WIDTHS"]["DYE"]))
        dye4 = st.text_input("ALLOWED_MAXIMO_FOR_4_WIDTHS_DYE", value=",".join(str(int(x)) for x in params["ALLOWED_MAXIMO_FOR_4_WIDTHS"]["DYE"]))
    with c2:
        bl3 = st.text_input("ALLOWED_MAXIMO_FOR_3_WIDTHS_BLEACH", value=",".join(str(int(x)) for x in params["ALLOWED_MAXIMO_FOR_3_WIDTHS"]["BLEACH"]))
        bl4 = st.text_input("ALLOWED_MAXIMO_FOR_4_WIDTHS_BLEACH", value=",".join(str(int(x)) for x in params["ALLOWED_MAXIMO_FOR_4_WIDTHS"]["BLEACH"]))

    def _parse_set(s):
        out = set()
        for x in s.split(","):
            x = x.strip()
            if x:
                try:
                    out.add(float(x))
                except Exception:
                    pass
        return out

    params["ALLOWED_MAXIMO_FOR_3_WIDTHS"] = {"DYE": _parse_set(dye3), "BLEACH": _parse_set(bl3)}
    params["ALLOWED_MAXIMO_FOR_4_WIDTHS"] = {"DYE": _parse_set(dye4), "BLEACH": _parse_set(bl4)}

st.session_state["params"] = params

# ---------------------------- 3. Ejecución ----------------------------
st.header("3. Ejecutar Loteo")
if st.button("▶️ Ejecutar Loteo", type="primary"):
    if st.session_state["df_data"] is None:
        st.error("Primero carga un archivo válido.")
    else:
        try:
            df_cap_final = build_cap_dataframe(st.session_state["df_cap"].to_dict("records"))
            progress_bar = st.progress(0.0, text="Iniciando...")

            def cb(frac, msg):
                progress_bar.progress(min(1.0, frac), text=msg)

            with st.spinner("Ejecutando algoritmo de loteo..."):
                st.session_state.get("df_data_filtrada", st.session_state["df_data"]).attrs["LOTEO_PARAMS"] = params
                df_detalle, df_resumen, df_exced, df_param_out = run_loteo(
                    st.session_state.get("df_data_filtrada", st.session_state["df_data"]), df_cap_final, params, progress_cb=cb
                )
                reports = build_reports(st.session_state.get("df_data_filtrada", st.session_state["df_data"]), df_cap_final, df_detalle, df_resumen)

            st.session_state["excel_bytes"] = None
            st.session_state["resultado"] = {
                "df_detalle": df_detalle, "df_resumen": df_resumen, "df_exced": df_exced,
                "df_param_out": df_param_out, "reports": reports, "df_cap_final": df_cap_final
            }
            st.success(f"✅ Loteo completado: {len(df_resumen)} lotes creados.")
        except Exception as e:
            st.exception(e)

# ---------------------------- 4 & 5. Resultados y gráficos ----------------------------
if st.session_state["resultado"] is not None:
    res = st.session_state["resultado"]
    df_detalle = res["df_detalle"]
    df_resumen = res["df_resumen"]
    df_exced = res["df_exced"]
    reports = res["reports"]

    st.header("4. KPIs")
    total_data = float(st.session_state.get("df_data_filtrada", st.session_state["df_data"])["TOTAL"].sum())
    total_asignado = float(df_detalle["LBS_ASIGNADAS"].sum()) if len(df_detalle) else 0.0
    pct_asignado = (total_asignado / total_data * 100) if total_data > 0 else 0.0
    n_lotes = len(df_resumen)
    skus_sin_asignar = df_exced["LNK"].nunique() if len(df_exced) else 0
    lbs_excedentes = float(df_exced["LBS_RESTANTES"].sum()) if len(df_exced) else 0.0
    fill_avg = float(reports["CAPACIDAD_X_CATEG"]["FILL_RATE"].mean()) * 100 if len(reports["CAPACIDAD_X_CATEG"]) else 0.0

    pct_sin_asignar = (lbs_excedentes / total_data * 100) if total_data > 0 else 0.0
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Total libras", f"{total_data:,.0f}")
    k2.metric("Libras asignadas", f"{total_asignado:,.0f}", f"{pct_asignado:.1f}% del total")
    k3.metric("Libras sin asignar", f"{lbs_excedentes:,.0f}", f"{pct_sin_asignar:.1f}% del total", delta_color="inverse")
    k4.metric("Lotes creados", f"{n_lotes:,}")
    k5.metric("Fill rate promedio", f"{fill_avg:.1f}%")

    st.subheader("Resumen ejecutivo")
    ec1, ec2 = st.columns(2)
    with ec1:
        if len(reports["PRIORIDAD_VS_ASIG"]):
            dfx = reports["PRIORIDAD_VS_ASIG"].copy()
            fig_exec = px.bar(dfx, x="BLOQUE", y=["LBS_ASIGNADAS", "LBS_SIN_ASIGNAR"], color_discrete_sequence=["#0072CE", "#D9E2F2"], barmode="stack", title="Cobertura por prioridad")
            fig_exec.update_layout(template="plotly_white", legend_title_text="", margin=dict(t=55,l=10,r=10,b=10))
            st.plotly_chart(fig_exec, use_container_width=True)
    with ec2:
        if len(df_resumen):
            dfr = df_resumen.groupby("CATEGORIA", as_index=False).agg(LBS_TOTAL=("LBS_TOTAL","sum"), CAPACIDAD_PERDIDA=("CAPACIDAD_PERDIDA","sum"))
            fig_loss = px.bar(dfr, x="CATEGORIA", y="CAPACIDAD_PERDIDA", color="LBS_TOTAL", color_continuous_scale=[[0,"#9ecae1"],[1,"#005a9c"]], title="Oportunidad por capacidad no utilizada")
            fig_loss.update_layout(template="plotly_white", coloraxis_colorbar_title="Lbs loteadas", margin=dict(t=55,l=10,r=10,b=10))
            st.plotly_chart(fig_loss, use_container_width=True)
    if len(df_exced):
        dims = [c for c in ["PLANTA_COSTURA","PRIORIDAD","CONSTRUCCION","TIPO_TEJIDO","MIX"] if c in df_exced.columns]
        if dims:
            dim = st.selectbox("Analizar excedentes por", dims, index=0)
            dfe = df_exced.groupby(dim, as_index=False)["LBS_RESTANTES"].sum().nlargest(15, "LBS_RESTANTES")
            fig_exc = px.bar(dfe, x="LBS_RESTANTES", y=dim, orientation="h", title=f"Top excedentes por {dim}", color="LBS_RESTANTES", color_continuous_scale="Blues")
            fig_exc.update_layout(template="plotly_white", yaxis={"categoryorder":"total ascending"})
            st.plotly_chart(fig_exc, use_container_width=True)

    st.header("5. Reportes")
    report_tabs = st.tabs([
        "DETALLE_LOTES", "RESUMEN_LOTES", "EXCEDENTES", "PARAMETROS",
        "CAPACIDAD_X_CATEG", "PRIORIDAD_VS_ASIG", "LNK_COMPLETITUD",
        "REGLA_STYLE_ANCHO18", "REGLA_COMBINACION_ANCHOS", "REGLA_COLOR_R",
        "REGLA_FAMILIA", "REPORTE_REGLAS_MIX", "OVERSHOOT_SUMMARY", "DECISION_LOG",
        "TOP_CAUSAS_NO_ASIGNACION", "AUDITORIA_REGLAS_LOTES"
    ])
    all_reports = {
        "DETALLE_LOTES": df_detalle, "RESUMEN_LOTES": df_resumen, "EXCEDENTES": df_exced,
        "PARAMETROS": res["df_param_out"],
        **reports
    }
    for tab, (name, df) in zip(report_tabs, all_reports.items()):
        with tab:
            st.dataframe(df, use_container_width=True)
            st.download_button(f"⬇️ Descargar {name}.csv", data=df.to_csv(index=False).encode("utf-8"),
                                file_name=f"{name}.csv", mime="text/csv", key=f"dl_{name}")

    st.header("6. Tablero ejecutivo")
    ELC_BLUE, ELC_NAVY, ELC_SKY, ELC_ORANGE, ELC_GRAY = "#0072CE", "#082B54", "#65B5E8", "#F28E2B", "#D9E2F2"
    def estilo_ejecutivo(fig, titulo, altura=430):
        fig.update_layout(
            title=dict(text=titulo, x=0.02, xanchor="left", font=dict(size=19, color=ELC_NAVY)),
            template="plotly_white", height=altura, margin=dict(l=35, r=25, t=70, b=45),
            font=dict(family="Arial", size=12, color="#243447"),
            paper_bgcolor="white", plot_bgcolor="white",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            hoverlabel=dict(bgcolor="white", font_size=12, font_family="Arial")
        )
        fig.update_xaxes(showgrid=False, linecolor="#D9E2F2")
        fig.update_yaxes(gridcolor="#EAF0F6", zeroline=False)
        return fig

    g1, g2 = st.columns(2)
    with g1:
        if len(reports["CAPACIDAD_X_CATEG"]):
            dc = reports["CAPACIDAD_X_CATEG"].copy()
            dc["ETIQUETA"] = dc["CATEGORIA"] + " · " + dc["MIX"]
            fig = go.Figure()
            fig.add_bar(x=dc["ETIQUETA"], y=dc["CAPACIDAD"], name="Capacidad", marker_color=ELC_GRAY,
                        hovertemplate="%{x}<br>Capacidad: %{y:,.0f} lb<extra></extra>")
            fig.add_bar(x=dc["ETIQUETA"], y=dc["LBS_ASIGNADAS"], name="Asignado", marker_color=ELC_BLUE,
                        text=dc["FILL_RATE"].map(lambda x: f"{x:.0%}"), textposition="outside",
                        hovertemplate="%{x}<br>Asignado: %{y:,.0f} lb<extra></extra>")
            fig.update_layout(barmode="group")
            st.plotly_chart(estilo_ejecutivo(fig, "Utilización de capacidad por categoría"), use_container_width=True)
    with g2:
        if len(reports["PRIORIDAD_VS_ASIG"]):
            dp = reports["PRIORIDAD_VS_ASIG"].copy()
            orden = ["VENCIDOS", "DUE", "AHEAD", "AHEAD2", "OTROS"]
            dp["BLOQUE"] = pd.Categorical(dp["BLOQUE"], categories=orden, ordered=True)
            dp = dp.sort_values(["MIX", "BLOQUE"])
            fig = go.Figure()
            fig.add_bar(x=dp["BLOQUE"].astype(str), y=dp["LBS_ASIGNADAS"], name="Asignadas", marker_color=ELC_BLUE,
                        hovertemplate="%{x}<br>Asignadas: %{y:,.0f} lb<extra></extra>")
            fig.add_bar(x=dp["BLOQUE"].astype(str), y=dp["LBS_SIN_ASIGNAR"], name="Pendientes", marker_color=ELC_ORANGE,
                        hovertemplate="%{x}<br>Pendientes: %{y:,.0f} lb<extra></extra>")
            fig.update_layout(barmode="stack")
            st.plotly_chart(estilo_ejecutivo(fig, "Cobertura por bloque de prioridad"), use_container_width=True)

    g3, g4 = st.columns(2)
    with g3:
        if len(df_resumen):
            dw = df_resumen["ANCHOS_UNICOS"].value_counts().sort_index().rename_axis("ANCHOS").reset_index(name="LOTES")
            fig = px.bar(dw, x="ANCHOS", y="LOTES", text="LOTES", color="ANCHOS",
                         color_continuous_scale=[[0, ELC_SKY], [1, ELC_NAVY]])
            fig.update_traces(textposition="outside", hovertemplate="%{x} anchos<br>%{y} lotes<extra></extra>")
            fig.update_layout(coloraxis_showscale=False)
            st.plotly_chart(estilo_ejecutivo(fig, "Composición de lotes por cantidad de anchos"), use_container_width=True)
    with g4:
        if len(reports["LNK_COMPLETITUD"]):
            ds = reports["LNK_COMPLETITUD"]["ESTADO"].value_counts().rename_axis("ESTADO").reset_index(name="LNK")
            fig = px.pie(ds, names="ESTADO", values="LNK", hole=.62,
                         color="ESTADO", color_discrete_map={"COMPLETO":ELC_BLUE,"COMPLETO (SCRAP)":ELC_ORANGE,"INCOMPLETO":ELC_GRAY})
            fig.update_traces(textposition="inside", textinfo="percent+label", hovertemplate="%{label}: %{value} LNK<br>%{percent}<extra></extra>")
            fig.add_annotation(text=f"{ds['LNK'].sum():,.0f}<br><span style='font-size:12px'>LNK</span>", x=.5, y=.5, showarrow=False, font=dict(size=22,color=ELC_NAVY))
            st.plotly_chart(estilo_ejecutivo(fig, "Completitud de LNK"), use_container_width=True)

    g5, g6 = st.columns(2)
    with g5:
        if len(df_resumen):
            dr = df_resumen["REGLA_DOMINANTE"].fillna("SIN REGLA").value_counts().head(10).sort_values().rename_axis("REGLA").reset_index(name="LOTES")
            fig = px.bar(dr, x="LOTES", y="REGLA", orientation="h", text="LOTES", color="LOTES",
                         color_continuous_scale=[[0, ELC_SKY], [1, ELC_NAVY]])
            fig.update_traces(textposition="outside", hovertemplate="%{y}: %{x} lotes<extra></extra>")
            fig.update_layout(coloraxis_showscale=False)
            st.plotly_chart(estilo_ejecutivo(fig, "Reglas dominantes", 450), use_container_width=True)
    with g6:
        if len(df_exced):
            de = df_exced.groupby("MIX", as_index=False)["LBS_RESTANTES"].sum().sort_values("LBS_RESTANTES", ascending=False)
            fig = px.bar(de, x="MIX", y="LBS_RESTANTES", text_auto=",.0f", color="MIX",
                         color_discrete_sequence=[ELC_ORANGE, ELC_BLUE, ELC_SKY])
            fig.update_traces(textposition="outside", hovertemplate="%{x}: %{y:,.0f} lb<extra></extra>")
            st.plotly_chart(estilo_ejecutivo(fig, "Libras pendientes por proceso"), use_container_width=True)

    st.header("7. Descarga del Excel completo")
    if "excel_bytes" not in st.session_state:
        st.session_state["excel_bytes"] = None

    if st.button("📦 Generar Excel completo"):
        try:
            out_path = "/tmp/RESULTADOS_LOTES.xlsx"
            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                used_names = set()
                for name, df in all_reports.items():
                    sheet_name = name[:31]
                    # evitar nombres de hoja duplicados tras el truncado a 31 caracteres
                    base = sheet_name
                    i = 1
                    while sheet_name in used_names:
                        suffix = f"_{i}"
                        sheet_name = base[:31 - len(suffix)] + suffix
                        i += 1
                    used_names.add(sheet_name)
                    df_safe = df.copy()
                    # openpyxl no acepta tz-aware datetimes ni objetos no serializables; forzamos a str si hace falta
                    for col in df_safe.columns:
                        if df_safe[col].dtype == object:
                            df_safe[col] = df_safe[col].apply(lambda v: v if (v is None or isinstance(v, (str, int, float, bool))) else str(v))
                    df_safe.to_excel(writer, index=False, sheet_name=sheet_name)
            format_workbook(out_path, font_name="Arial", font_size=9)
            with open(out_path, "rb") as f:
                st.session_state["excel_bytes"] = f.read()
            st.success("✅ Excel generado. Usa el botón de abajo para descargarlo.")
        except Exception as e:
            st.session_state["excel_bytes"] = None
            st.error(f"❌ Error generando el Excel: {e}")

    if st.session_state["excel_bytes"] is not None:
        st.download_button(
            "⬇️ Descargar RESULTADOS_LOTES.xlsx",
            data=st.session_state["excel_bytes"],
            file_name="RESULTADOS_LOTES.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_excel_final",
        )
