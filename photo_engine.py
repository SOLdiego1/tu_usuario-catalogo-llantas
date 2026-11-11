"""
photo_engine.py
Motor de indexación y emparejamiento de fotos por diseño de llantas.
Localiza las 3 fotos correspondientes al diseño de cada producto, o retorna SFR.
"""

import os
import re
import unicodedata
import openpyxl

_opt_dir = os.path.join(os.path.dirname(__file__), "FOTOS_OPTIMIZADAS")
_orig_dir = os.path.join(os.path.dirname(__file__), "FOTOS DE LLANTAS")
DEFAULT_PHOTOS_DIR = _opt_dir if os.path.exists(_opt_dir) else _orig_dir

def clean_tok(text):
    if not text:
        return ""
    nfkd = unicodedata.normalize('NFKD', str(text))
    s = "".join([c for c in nfkd if not unicodedata.combining(c)]).lower()
    s = re.sub(r'[\(\)\[\]\,]', ' ', s)
    return s.strip()

def normalize_tyre_dimension(text):
    """
    Normaliza medidas de llantas para empatar nomenclaturas comerciales:
    ej: 600-12 -> 6.00-12, 750-16 -> 7.50-16, 500-12 -> 5.00-12, 1100-20 -> 11.00-20
    """
    if not text:
        return ""
    s = clean_tok(text)
    s = re.sub(r'^(ll\.|neu|llanta|neumatico)\s*', '', s).strip()
    
    def dot_replace(m):
        digits = m.group(1)
        sep = m.group(2)
        if len(digits) == 3 and digits[1:] in ['00', '25', '50', '75']:
            return f"{digits[0]}.{digits[1:]}{sep}"
        elif len(digits) == 4 and digits[2:] in ['00']:
            return f"{digits[:2]}.{digits[2:]}{sep}"
        return m.group(0)

    s = re.sub(r'\b(\d{3,4})([-\/xX])', dot_replace, s)
    m = re.findall(r'\b\d+(?:\.\d+)?(?:l|r|c)?[\/\-xX]\d+(?:\.\d+)?(?:r\d+(?:\.\d+)?)?', s)
    if m:
        return m[0].replace('x', '-').replace('X', '-')
    return s.split()[0] if s else ""

def normalize_brand(brand):
    b = clean_tok(brand)
    brands = [
        'advance', 'aitc', 'titan', 'goodyear', 'continental', 'general',
        'alliance', 'carlstar', 'atf', 'forerunner', 'galaxy', 'bkt',
        'deestone', 'trelleborg', 'michelin', 'firestone', 'chaoyang',
        'petlas', 'speedways', 'westlake', 'powerstone', 'sailun', 'double coin'
    ]
    for main_b in brands:
        if main_b in b:
            return main_b
    return b

class PhotoManager:
    def __init__(self, photos_base_dir=DEFAULT_PHOTOS_DIR):
        self.base_dir = photos_base_dir
        self.photo_folders = []
        self.listado_entries = []
        self._load_and_index()

    def _load_and_index(self):
        if not os.path.exists(self.base_dir):
            json_idx = os.path.join(os.path.dirname(__file__), "index_fotos.json")
            if os.path.exists(json_idx):
                import json
                try:
                    with open(json_idx, "r", encoding="utf-8") as f:
                        data = json.load(f)
                    for item in data:
                        # Priorizar CDN si la ruta local no existe
                        item['photos'] = item.get('photos_cdn', item.get('photos_local', []))
                        self.photo_folders.append(item)
                except Exception as e:
                    print(f"Aviso al cargar index_fotos.json: {e}")
            return

        # 1. Leer LISTADO LLANTAS FOTOS.xlsx si existe
        listado_file = os.path.join(self.base_dir, "LISTADO LLANTAS FOTOS.xlsx")
        if os.path.exists(listado_file):
            try:
                wb = openpyxl.load_workbook(listado_file, data_only=True)
                ws = wb['Hoja1']
                for r in ws.iter_rows(min_row=2, values_only=True):
                    if len(r) > 4 and (r[2] or r[3]):
                        self.listado_entries.append({
                            'linea': clean_tok(r[1]),
                            'referencia': str(r[2]).strip() if r[2] else "",
                            'marca': str(r[3]).strip() if r[3] else "",
                            'diseno': str(r[4]).strip() if r[4] else "",
                            'tipo': str(r[5]).strip() if len(r) > 5 and r[5] else ""
                        })
                wb.close()
            except Exception as e:
                print(f"Aviso: no se pudo leer LISTADO LLANTAS FOTOS.xlsx: {e}")

        # 2. Indexar todas las subcarpetas con imágenes
        for root, dirs, files in os.walk(self.base_dir):
            imgs = [f for f in files if f.lower().endswith(('.png', '.jpg', '.jpeg', '.webp', '.bmp'))]
            if imgs:
                rel = os.path.relpath(root, self.base_dir)
                parts = rel.split(os.sep)
                cat = parts[0] if len(parts) > 1 else "SIN CLASIFICAR"
                folder_name = parts[-1]

                # Ordenar fotos numéricamente (Foto 1, 2, 3...)
                def img_sort_key(fname):
                    nums = re.findall(r'\d+', fname)
                    return int(nums[0]) if nums else 999

                imgs_sorted = sorted(imgs, key=img_sort_key)
                full_imgs = [os.path.join(root, img) for img in imgs_sorted]

                dim = normalize_tyre_dimension(folder_name)
                brand = normalize_brand(folder_name)

                self.photo_folders.append({
                    'folder_name': folder_name,
                    'rel_path': rel,
                    'abs_path': root,
                    'category': cat,
                    'brand': brand,
                    'dimension': dim,
                    'clean_folder': clean_tok(folder_name),
                    'photos': full_imgs[:3],  # exactamente hasta 3 fotos por diseño
                    'all_photos': full_imgs,
                    'num_photos': len(full_imgs)
                })

    def get_photos_for_product(self, product_or_marca=None, referencia="", diseno="", linea="", **kwargs):
        """
        Retorna la ficha de fotos (con exactamente 3 fotos por diseño) si existe,
        o None si no existe foto registrada (indicando SFR.).
        Acepta un diccionario de producto como primer argumento, o argumentos nombrados.
        """
        if not self.photo_folders:
            return None

        if isinstance(product_or_marca, dict):
            marca = product_or_marca.get('marca', '')
            referencia = product_or_marca.get('referencia', '')
            diseno = product_or_marca.get('diseno', '')
            linea = product_or_marca.get('linea', product_or_marca.get('tipo', ''))
        else:
            marca = product_or_marca or kwargs.get('marca', '')
            referencia = referencia or kwargs.get('referencia', '')
            diseno = diseno or kwargs.get('diseno', '')
            linea = linea or kwargs.get('linea', '')

        b_norm = normalize_brand(marca)
        dim_norm = normalize_tyre_dimension(referencia)
        d_clean = clean_tok(diseno)
        ref_clean = clean_tok(referencia)

        # Tier 1: Marca + Dimensión + Token de Diseño
        if b_norm and dim_norm and d_clean and len(d_clean) >= 2:
            for pf in self.photo_folders:
                if pf['brand'] == b_norm and (dim_norm in pf['dimension'] or pf['dimension'] in dim_norm):
                    if d_clean in pf['clean_folder']:
                        return pf

        # Tier 2: Marca + Dimensión
        if b_norm and dim_norm:
            for pf in self.photo_folders:
                if pf['brand'] == b_norm and (dim_norm in pf['dimension'] or pf['dimension'] in dim_norm):
                    return pf

        # Tier 3: Dimensión + Diseño
        if dim_norm and d_clean and len(d_clean) >= 3:
            for pf in self.photo_folders:
                if (dim_norm in pf['dimension'] or pf['dimension'] in dim_norm) and d_clean in pf['clean_folder']:
                    return pf

        # Tier 4: Cruce con LISTADO LLANTAS FOTOS.xlsx
        for le in self.listado_entries:
            le_b = normalize_brand(le['marca'])
            le_dim = normalize_tyre_dimension(le['referencia'])
            if (b_norm and le_b == b_norm) and (dim_norm and le_dim == dim_norm):
                # Buscar carpeta que coincida con esta fila
                for pf in self.photo_folders:
                    if (dim_norm in pf['dimension'] or pf['dimension'] in dim_norm) and (b_norm == pf['brand']):
                        return pf

        # Si no se encuentra foto registrada
        return None

    def get_all_designs(self):
        """Retorna todas las carpetas de diseño indexadas con sus fotos para el catálogo visual."""
        return self.photo_folders
