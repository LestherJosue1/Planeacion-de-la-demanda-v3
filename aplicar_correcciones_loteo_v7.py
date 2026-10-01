from pathlib import Path
import re
import sys

SOURCE_DEFAULT = "app_loteo_elcatex_v5_1_color_lnk.py"
OUTPUT_DEFAULT = "app_loteo_elcatex_v7_categorias_grandes.py"

src = Path(sys.argv[1] if len(sys.argv) > 1 else SOURCE_DEFAULT)
out = Path(sys.argv[2] if len(sys.argv) > 2 else OUTPUT_DEFAULT)
if not src.exists():
    raise SystemExit(
        f"No se encontró {src}. Coloque este aplicador junto al código actual o ejecute: "
        f"python {Path(__file__).name} ruta_codigo_actual.py {OUTPUT_DEFAULT}"
    )

text = src.read_text(encoding="utf-8")
original = text
log = []


def replace_once(pattern, replacement, label, flags=0):
    global text
    updated, count = re.subn(pattern, replacement, text, count=1, flags=flags)
    if count != 1:
        raise RuntimeError(f"No se pudo aplicar '{label}'. Coincidencias encontradas: {count}")
    text = updated
    log.append(label)


# -----------------------------------------------------------------------------
# 1. Parámetros seguros por defecto
# -----------------------------------------------------------------------------
text, n = re.subn(
    r'params\.get\("REQUIRE_WIDTHS_STRICT",\s*1\)',
    'params.get("REQUIRE_WIDTHS_STRICT", 0)',
    text,
)
if n:
    log.append(f"REQUIRE_WIDTHS_STRICT default=0 ({n} apariciones)")

text, n = re.subn(
    r'("REQUIRE_WIDTHS_STRICT"\s*:\s*)1\b',
    r'\g<1>0',
    text,
)
if n:
    log.append(f"REQUIRE_WIDTHS_STRICT parámetro=0 ({n} apariciones)")

# BEAM_WIDTH 12 permite explorar más semillas, pero ya no decide la categoría.
text, n = re.subn(
    r'params\.get\("BEAM_WIDTH",\s*3\)',
    'params.get("BEAM_WIDTH", 12)',
    text,
)
if n:
    log.append(f"BEAM_WIDTH default=12 ({n} apariciones)")
text, n = re.subn(r'("BEAM_WIDTH"\s*:\s*)3\b', r'\g<1>12', text)
if n:
    log.append(f"BEAM_WIDTH parámetro=12 ({n} apariciones)")

# -----------------------------------------------------------------------------
# 2. Los anchos son un máximo por categoría, no un target mínimo/exacto
# -----------------------------------------------------------------------------
replace_once(
    r'''order_text\s*=\s*norm_str\(params\.get\("WIDTHS_TARGET_ORDER",\s*"4>3>2>1"\)\)\s*\n\s*targets\s*=\s*\[int\(x\)\s+for\s+x\s+in\s+order_text\.split\(">"\)\s+if\s+x\.strip\(\)\.isdigit\(\)\]\s*\n\s*req_strict\s*=\s*int\(params\.get\("REQUIRE_WIDTHS_STRICT",\s*[01]\)\)\s*==\s*1''',
    '''order_text = norm_str(params.get("WIDTHS_TARGET_ORDER", "4>3>2>1"))
                    # La cantidad de anchos es un LIMITE MAXIMO por categoría.
                    # WIDTHS_TARGET_ORDER queda solo como preferencia informativa/scoring.
                    targets = [None]
                    req_strict = False''',
    "anchos interpretados como máximo",
    re.MULTILINE,
)

# No filtrar categorías por un supuesto objetivo exacto de 3/4 anchos.
replace_once(
    r'candidate_ranges_all\s*=\s*filter_ranges_for_width_target\(ranges_try,\s*mixv,\s*target,\s*params\)',
    'candidate_ranges_all = sorted(ranges_try, key=lambda rr: (-float(rr["MAXIMO"]), str(rr["CATEGORIA"])))',
    "sin filtro obligatorio por target de anchos",
)

# En la llamada principal no colocar target como mínimo. El máximo ya se valida
# dentro de intentar_lote_para_rango mediante MAX_WIDTHS_BY_CAT[categoría].
text, n1 = re.subn(r'min_unique_widths\s*=\s*target', 'min_unique_widths=None', text)
text, n2 = re.subn(
    r'max_unique_widths\s*=\s*\(target\s+if\s+req_strict\s+else\s+None\)',
    'max_unique_widths=None',
    text,
)
if not n1:
    raise RuntimeError("No se encontró min_unique_widths=target en el motor actual.")
log.append(f"target eliminado como mínimo de anchos ({n1})")
if n2:
    log.append(f"target eliminado como máximo exacto ({n2})")

# -----------------------------------------------------------------------------
# 3. Selección lexicográfica: categoría grande > fill > menor pérdida > score
# -----------------------------------------------------------------------------
# Se conserva score_lote como desempate, no como mecanismo que pueda degradar
# una categoría grande factible a una pequeña.
text, n = re.subn(
    r'best_score\s*=\s*-1e30',
    'best_score = -1e30\n                  best_selection_key = None',
    text,
    count=1,
)
if n != 1:
    raise RuntimeError("No se encontró la inicialización principal de best_score.")
log.append("clave lexicográfica inicializada")

old_selection = r'''sc\s*=\s*score_lote\(lote_for_score,\s*resumen_rows,\s*params,\s*categoria=lote\["CATEGORIA"\],\s*seed_row=seed_row_dict\)\s*\n\s*if\s+sc\s*>\s*best_score:\s*\n\s*best_score\s*=\s*sc\s*\n\s*best_lote\s*=\s*lote\s*\n\s*best_pack\s*=\s*\(lote,\s*rule_info,\s*prioridad_obj,\s*best_score\)'''
new_selection = '''sc = score_lote(lote_for_score, resumen_rows, params, categoria=lote["CATEGORIA"], seed_row=seed_row_dict)
                          maximo_cat = float(lote["MAXIMO"])
                          total_cat = float(lote["TOTAL_LOTE"])
                          fill_cat = (total_cat / maximo_cat) if maximo_cat > 1e-9 else 0.0
                          perdida_cat = max(0.0, maximo_cat - total_cat)
                          # Orden duro: categoría mayor, luego mejor llenado, menor pérdida y score.
                          selection_key = (maximo_cat, fill_cat, -perdida_cat, float(sc))
                          if best_selection_key is None or selection_key > best_selection_key:
                              best_selection_key = selection_key
                              best_score = sc
                              best_lote = lote
                              best_pack = (lote, rule_info, prioridad_obj, best_score)'''
replace_once(old_selection, new_selection, "selección lexicográfica de categoría", re.MULTILINE)

# -----------------------------------------------------------------------------
# 4. Evitar el break prematuro al hallar el primer rango principal válido
# -----------------------------------------------------------------------------
# En el bucle primario se guardan candidatos de todos los rangos. El bloque de
# selección posterior decide cuál consumir. El reemplazo se limita al patrón
# exacto que asignaba lote/prioridad/found y hacía break.
primary_pattern = r'''if\s+intento\s+is\s+not\s+None:\s*\n\s*lote\s*=\s*intento\s*\n\s*prioridad_obj\s*=\s*float\(pri\)\s+if\s+pri\s+is\s+not\s+None\s+else\s+None\s*\n\s*found\s*=\s*True\s*\n\s*break'''
primary_replacement = '''if intento is not None:
                                      # No consumir ni detenerse en la primera categoría válida.
                                      # Conservar el mejor candidato local con el mismo orden duro.
                                      pri_tmp = float(pri) if pri is not None else None
                                      max_tmp = float(intento["MAXIMO"])
                                      total_tmp = float(intento["TOTAL_LOTE"])
                                      fill_tmp = total_tmp / max_tmp if max_tmp > 1e-9 else 0.0
                                      loss_tmp = max(0.0, max_tmp - total_tmp)
                                      local_key = (max_tmp, fill_tmp, -loss_tmp)
                                      if lote is None or local_key > lote_local_key:
                                          lote = intento
                                          lote_local_key = local_key
                                          prioridad_obj = pri_tmp
                                      found = True'''
text, n = re.subn(primary_pattern, primary_replacement, text, count=1, flags=re.MULTILINE)
if n != 1:
    raise RuntimeError("No se encontró el break primario de primera categoría válida.")
log.append("eliminado break de primera categoría válida")

# Inicializar clave local antes del recorrido de targets/rangos.
replace_once(
    r'lote\s*=\s*None\s*\n\s*prioridad_obj\s*=\s*None\s*\n\s*order_text',
    'lote = None\n                    lote_local_key = None\n                    prioridad_obj = None\n\n                    order_text',
    "clave local de comparación",
    re.MULTILINE,
)

# Los breaks exteriores dependientes de found impedirían revisar prioridades.
# Se retira únicamente el primero que aparece inmediatamente después del bloque
# principal de rangos; los fallbacks se conservan.
text, n = re.subn(
    r'\n\s*if\s+found:\s*\n\s*break\s*\n\s*if\s+lote\s+is\s+not\s+None:\s*\n\s*break',
    '\n                          # Continuar revisando todas las prioridades y rangos factibles.\n                          pass',
    text,
    count=1,
)
if n:
    log.append("evaluación completa de prioridades/rangos")
else:
    raise RuntimeError("No se encontró el cierre prematuro del bucle principal.")

# -----------------------------------------------------------------------------
# 5. Remanentes reutilizables 100-249 sin crear split nuevo menor de 250
# -----------------------------------------------------------------------------
# Reemplaza choose_take conservando la firma funcional, con argumento opcional.
choose_take_new = '''def choose_take(rest, remaining, split_min_lbs, allow_scrap_residue=False, small_reuse_min=100.0):
    """No crea splits nuevos menores al mínimo; reutiliza saldos completos 100-249."""
    try:
        rest = float(rest)
        remaining = float(remaining)
        split_min_lbs = float(split_min_lbs)
        small_reuse_min = float(small_reuse_min)
    except Exception:
        return 0.0
    if rest <= 1e-9 or remaining <= 1e-9:
        return 0.0
    # Un saldo ya existente de 100-249 puede completar un lote, pero no iniciarlo.
    if rest <= remaining + 1e-9:
        return rest if rest + 1e-9 >= small_reuse_min else 0.0
    take = remaining
    # Nunca generar una porción nueva menor al split operativo.
    if take + 1e-9 < split_min_lbs:
        return 0.0
    residue = rest - take
    if residue <= 1e-9 or residue + 1e-9 >= split_min_lbs:
        return take
    if small_reuse_min - 1e-9 <= residue < split_min_lbs - 1e-9:
        return take
    return take if allow_scrap_residue else 0.0
'''
text, n = re.subn(
    r'def\s+choose_take\(rest,\s*remaining,\s*split_min_lbs,\s*allow_scrap_residue=False\):.*?(?=\n#\s*-{5,})',
    choose_take_new.rstrip(),
    text,
    count=1,
    flags=re.DOTALL,
)
if n != 1:
    raise RuntimeError("No se pudo sustituir choose_take de manera segura.")
log.append("reutilización de saldos 100-249")

# -----------------------------------------------------------------------------
# 6. Validaciones estáticas
# -----------------------------------------------------------------------------
required_fragments = [
    'targets = [None]',
    'req_strict = False',
    'selection_key = (maximo_cat, fill_cat, -perdida_cat, float(sc))',
    'local_key = (max_tmp, fill_tmp, -loss_tmp)',
    'small_reuse_min=100.0',
]
for fragment in required_fragments:
    if fragment not in text:
        raise RuntimeError(f"Validación fallida: falta {fragment}")

if text == original:
    raise RuntimeError("No se aplicaron cambios.")

compile(text, str(out), "exec")
out.write_text(text, encoding="utf-8")

print(f"OK: {out}")
for item in log:
    print(f"- {item}")
