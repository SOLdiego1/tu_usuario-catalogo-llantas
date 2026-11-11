"""
search_engine.py
Motor de indexación, búsqueda inteligente y desambiguación de productos
por Código CG1, Referencia comercial y Código de Fabricante.
"""

import re
import unicodedata

def clean_str(s):
    if s is None:
        return ""
    nfkd_form = unicodedata.normalize('NFKD', str(s))
    only_ascii = "".join([c for c in nfkd_form if not unicodedata.combining(c)])
    return only_ascii.strip().lower()

def search_products(products, query, selected_sheet=None, only_in_stock=False):
    """
    Búsqueda inteligente con priorización estricta:
    1. Coincidencia Exacta por Código CG1
    2. Coincidencia Exacta por Código de Fabricante
    3. Coincidencia Exacta por Dimensión/Token de Referencia (evita mezclar 500-12 con 500-12.5)
    4. Coincidencia Parcial en Referencia
    5. Coincidencia Parcial en Descripción y Marca
    """
    q_raw = str(query).strip()
    if not q_raw:
        return []

    q_clean = clean_str(q_raw)
    q_no_punct = re.sub(r'[^a-z0-9]', '', q_clean)

    escaped_q = re.escape(q_clean)
    exact_token_pattern = re.compile(rf'(?:^|[\s,;()\"\'\[\]/])({escaped_q})(?:$|[\s,;()\"\'\[\]/])', re.IGNORECASE)

    exact_cg1_matches = []
    exact_fab_matches = []
    exact_ref_matches = []
    partial_matches = []

    for p in products:
        if selected_sheet and selected_sheet != "Todas las pestañas":
            if p['source_sheet'] != selected_sheet:
                continue

        if only_in_stock and p['total_fisico'] <= 0:
            continue

        cg1 = p['cod_cg1']
        fab = p['fabricante_code']
        ref = p['referencia']
        desc = p['descripcion']
        marca = p['marca']

        cg1_clean = clean_str(cg1)
        fab_clean = clean_str(fab)
        ref_clean = clean_str(ref)
        desc_clean = clean_str(desc)

        # 1. Coincidencia exacta CG1
        if q_clean == cg1_clean or (q_no_punct and q_no_punct == re.sub(r'[^a-z0-9]', '', cg1_clean)):
            exact_cg1_matches.append({
                'product': p,
                'match_type': 'Código CG1 Exacto',
                'priority': 1,
                'score': 100,
                'match_detail': f"Coincidencia exacta con Código CG1: {cg1}"
            })
            continue

        # 2. Coincidencia exacta Fabricante
        if fab_clean and (q_clean == fab_clean or (q_no_punct and q_no_punct == re.sub(r'[^a-z0-9]', '', fab_clean))):
            exact_fab_matches.append({
                'product': p,
                'match_type': 'Código Fabricante Exacto',
                'priority': 2,
                'score': 95,
                'match_detail': f"Coincidencia exacta con Código Fabricante: {fab}"
            })
            continue

        # 3. Coincidencia exacta Token de Referencia
        if exact_token_pattern.search(ref_clean) or exact_token_pattern.search(desc_clean):
            exact_ref_matches.append({
                'product': p,
                'match_type': 'Referencia Exacta',
                'priority': 3,
                'score': 90,
                'match_detail': f"Coincidencia exacta con medida: {ref}"
            })
            continue

        # 4. Parcial en Referencia
        if q_clean in ref_clean:
            score = 75 if re.search(rf'\b{escaped_q}', ref_clean) else 65
            partial_matches.append({
                'product': p,
                'match_type': 'Coincidencia Parcial (Referencia)',
                'priority': 4,
                'score': score,
                'match_detail': f"Contiene '{q_raw}' en: {ref}"
            })
            continue

        # 5. Parcial en Descripción / Marca
        if q_clean in desc_clean or (len(q_clean) >= 4 and q_clean in clean_str(marca)):
            partial_matches.append({
                'product': p,
                'match_type': 'Coincidencia Parcial (Descripción/Marca)',
                'priority': 5,
                'score': 50,
                'match_detail': f"Contiene '{q_raw}' en descripción/marca"
            })
            continue

    results = exact_cg1_matches + exact_fab_matches + exact_ref_matches + partial_matches
    results.sort(key=lambda x: (x['priority'], -x['product']['total_fisico'], -x['score']))
    return results

def find_supplier_info(supplier_inventory, cod_cg1="", cod_fab="", referencia=""):
    if not supplier_inventory:
        return []

    cg1_clean = clean_str(cod_cg1)
    fab_clean = clean_str(cod_fab)
    ref_clean = clean_str(referencia)

    matches = []
    for item in supplier_inventory:
        i_cg1 = clean_str(item.get('cod_cg1'))
        i_fab = clean_str(item.get('cod_fabricante'))
        i_ref = clean_str(item.get('referencia'))

        if cg1_clean and i_cg1 and cg1_clean == i_cg1:
            matches.append(item)
            continue
        if fab_clean and i_fab and fab_clean == i_fab:
            matches.append(item)
            continue
        if ref_clean and len(ref_clean) >= 5 and (ref_clean in i_ref or i_ref in ref_clean):
            matches.append(item)
            continue

    return matches
