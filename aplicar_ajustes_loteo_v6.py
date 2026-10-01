from pathlib import Path
import re, sys

SRC = Path(sys.argv[1] if len(sys.argv) > 1 else 'app_loteo_elcatex_v5_1_color_lnk.py')
OUT = Path(sys.argv[2] if len(sys.argv) > 2 else 'app_loteo_elcatex_v6_optimizacion.py')
if not SRC.exists():
    raise SystemExit(f'No se encontró {SRC}. Coloque este parche junto al archivo o indique la ruta como primer argumento.')
s = SRC.read_text(encoding='utf-8')
original = s
changes=[]

def sub_once(pattern, repl, label, flags=0):
    global s
    ns,n = re.subn(pattern,repl,s,count=1,flags=flags)
    if n != 1:
        raise RuntimeError(f'No se pudo aplicar de forma única: {label} (coincidencias={n})')
    s=ns; changes.append(label)

# 1) Beam width razonable para ampliar exploración sin búsqueda explosiva.
s, n = re.subn(r'params\.get\("BEAM_WIDTH",\s*3\)', 'params.get("BEAM_WIDTH", 12)', s)
if n: changes.append(f'BEAM_WIDTH runtime=12 ({n})')
s, n = re.subn(r'("BEAM_WIDTH"\s*:\s*)3\b', r'\g<1>12', s)
if n: changes.append(f'BEAM_WIDTH defaults=12 ({n})')
s, n = re.subn(r'(params\.get\("BEAM_WIDTH",\s*)3(\))', r'\g<1>12\2', s)
if n: changes.append(f'BEAM_WIDTH UI=12 ({n})')

# 2) Cantidad de anchos = máximo por categoría, nunca cantidad obligatoria.
sub_once(
 r'order_text\s*=\s*norm_str\(params\.get\("WIDTHS_TARGET_ORDER",\s*"4>3>2>1"\)\)\s*\n\s*targets\s*=\s*\[int\(x\) for x in order_text\.split\(">"\) if x\.strip\(\)\.isdigit\(\)\]\s*\n\s*req_strict\s*=\s*int\(params\.get\("REQUIRE_WIDTHS_STRICT",\s*1\)\)\s*==\s*1',
 'order_text = norm_str(params.get("WIDTHS_TARGET_ORDER", "4>3>2>1"))\n                    # El número de anchos es un límite superior por categoría, no un objetivo exacto.\n                    targets = [None]\n                    req_strict = False',
 'anchos como máximo, no exactos', re.M)
sub_once(
 r'candidate_ranges_all\s*=\s*filter_ranges_for_width_target\(ranges_try,\s*mixv,\s*target,\s*params\)',
 'candidate_ranges_all = sorted(ranges_try, key=lambda rr: (-float(rr["MAXIMO"]), str(rr["CATEGORIA"])))',
 'categorías grandes ordenadas primero')

# 3) Prioridad fuerte, pero no absoluta, para categorías grandes entre candidatos factibles.
sub_once(
 r'score\s*=\s*\(W_FILL \* fill\) \+ \(-W_CAP_LOSS \* cap_loss\) \+ \(W_WIDTH_PREF \* width_pref_score\)',
 'score = (W_FILL * fill) + (-W_CAP_LOSS * cap_loss) + (W_WIDTH_PREF * width_pref_score)\n    # Front-load: entre lotes técnicamente válidos, favorecer la categoría de mayor capacidad.\n    # A-4000 recibe el mayor bono, luego B-3300, C-2600, etc.\n    W_LARGE_CATEGORY = float(params.get("W_LARGE_CATEGORY", 25.0))\n    score += W_LARGE_CATEGORY * (maximo / 4000.0)',
 'bono categorías grandes')

# 4) Sustituir split chooser: no crear splits <250; remanentes completos 100-249 sí se reutilizan.
choose_new = '''def choose_take(rest, remaining, split_min_lbs, allow_scrap_residue=False, small_reuse_min=100.0):
    """Selecciona libras sin crear splits menores al mínimo.

    - Una porción nueva debe ser >= split_min_lbs (normalmente 250).
    - Un remanente YA existente entre 100 y 249 puede entrar completo como filler.
    - Al recortar una fila, se permite dejar un remanente reutilizable entre 100 y split_min-1.
    - No se crean lotes independientes con estos remanentes; solo completan lotes factibles.
    """
    try:
        rest = float(rest); remaining = float(remaining); split_min_lbs = float(split_min_lbs)
        small_reuse_min = float(small_reuse_min)
    except Exception:
        return 0.0
    if rest <= 1e-9 or remaining <= 1e-9:
        return 0.0
    # Consumo completo: permite reutilizar saldos 100-249 como filler.
    if rest <= remaining + 1e-9:
        return rest if rest + 1e-9 >= small_reuse_min else 0.0
    take = remaining
    # Una división nueva nunca puede ser menor al split operativo.
    if take + 1e-9 < split_min_lbs:
        return 0.0
    residue = rest - take
    if residue <= 1e-9:
        return take
    # Residuo normal o residuo pequeño reutilizable.
    if residue + 1e-9 >= split_min_lbs:
        return take
    if small_reuse_min - 1e-9 <= residue < split_min_lbs - 1e-9:
        return take
    # Menor de 100: solo si explícitamente se autorizó scrap.
    return take if allow_scrap_residue else 0.0
'''
sub_once(
 r'def choose_take\(rest,\s*remaining,\s*split_min_lbs,\s*allow_scrap_residue=False\):.*?(?=\n# ----------------------------)',
 choose_new.rstrip(), 'reutilización de remanentes 100-249', re.S)

# 5) Parámetros visibles y auditables.
anchor='"BEAM_WIDTH": 12'
if anchor in s and '"SMALL_SPLIT_REUSE_MIN"' not in s:
    s=s.replace(anchor, anchor + ',\n        "SMALL_SPLIT_REUSE_MIN": 100.0,\n        "PRIORIZE_LARGE_CATEGORIES": 1,\n        "W_LARGE_CATEGORY": 25.0', 1)
    changes.append('parámetros auditables')
# Pasar el mínimo de reutilización a choose_take en todas las llamadas.
s, n = re.subn(
 r'choose_take\(([^\n]*?),\s*allow_scrap_residue=allow_scrap_residue\)',
 r'choose_take(\1, allow_scrap_residue=allow_scrap_residue, small_reuse_min=float(params.get("SMALL_SPLIT_REUSE_MIN", 100.0)))', s)
if n: changes.append(f'small_reuse_min conectado ({n})')

# Controles de consistencia.
if s == original: raise RuntimeError('No se aplicó ningún cambio.')
if 'targets = [None]' not in s: raise RuntimeError('No quedó desactivada la exigencia exacta de anchos.')
if 'W_LARGE_CATEGORY' not in s: raise RuntimeError('No quedó activa la prioridad de categorías grandes.')
if 'small_reuse_min' not in s: raise RuntimeError('No quedó activa la reutilización de remanentes.')
compile(s, str(OUT), 'exec')
OUT.write_text(s, encoding='utf-8')
print('OK:', OUT)
for x in changes: print('-', x)
