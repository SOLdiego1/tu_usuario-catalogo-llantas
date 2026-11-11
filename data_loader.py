"""
data_loader.py
Motor de lectura, procesamiento y extracción de inventarios y listas de precios
desde archivos Excel de OneDrive para Llantas y Neumáticos.
"""

import os
import glob
from datetime import datetime
import openpyxl

SEARCH_DIRS = [
    os.path.dirname(__file__),
    os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "prueba diego"))
]

def get_latest_excel_file(directories=SEARCH_DIRS):
    """
    Busca el archivo Excel (.xlsx / .xlsm) más reciente en los directorios de OneDrive configurados.
    Ignora archivos temporales (~$) y archivos de fotos (LISTADO LLANTAS FOTOS).
    """
    candidates = []
    for d in directories:
        if os.path.exists(d):
            patterns = [os.path.join(d, "*.xlsx"), os.path.join(d, "*.xlsm")]
            for pat in patterns:
                for f in glob.glob(pat):
                    base = os.path.basename(f)
                    if not base.startswith("~$") and "LISTADO" not in base.upper():
                        candidates.append(f)

    if not candidates:
        return None
    candidates.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    return candidates[0]

def get_file_sync_metadata(file_path):
    """
    Calcula metadatos de sincronización con OneDrive:
    - Fecha y hora de última modificación
    - Días y horas transcurridas
    - Alerta si han pasado más de 3 días sin actualizar
    """
    if not file_path or not os.path.exists(file_path):
        return {
            'exists': False,
            'file_name': "No encontrado",
            'file_path': file_path or "",
            'last_modified': None,
            'last_modified_str': "Sin información",
            'days_ago': 0,
            'hours_ago': 0,
            'status': "MISSING",
            'status_label': "Archivo no encontrado",
            'status_color': "red"
        }
    
    mtime = os.path.getmtime(file_path)
    dt_mod = datetime.fromtimestamp(mtime)
    now = datetime.now()
    delta = now - dt_mod
    
    days_ago = delta.days
    hours_ago = int(delta.seconds / 3600)
    
    if days_ago > 3:
        status = "OUTDATED"
        status_label = f"⚠️ Requiere actualización: Hace {days_ago} días no se actualiza en OneDrive (> 3 días)"
        status_color = "warning"
    else:
        status = "UP_TO_DATE"
        if days_ago == 0:
            tiempo_str = f"hace {hours_ago} horas" if hours_ago > 0 else "hace unos minutos"
        elif days_ago == 1:
            tiempo_str = "hace 1 día"
        else:
            tiempo_str = f"hace {days_ago} días"
        status_label = f"🟢 Al día: Actualizado {tiempo_str}"
        status_color = "success"

    return {
        'exists': True,
        'file_name': os.path.basename(file_path),
        'file_path': file_path,
        'file_size_mb': round(os.path.getsize(file_path) / (1024 * 1024), 2),
        'last_modified': dt_mod,
        'last_modified_str': dt_mod.strftime("%d/%m/%Y %I:%M %p"),
        'days_ago': days_ago,
        'hours_ago': hours_ago,
        'status': status,
        'status_label': status_label,
        'status_color': status_color
    }

def clean_code(val):
    if val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s

def normalize_text(text):
    if text is None:
        return ""
    return str(text).strip()

def format_cop(val):
    """Formatea valores monetarios en pesos colombianos ($ 1.250.000)."""
    if val is None or val == "" or str(val).strip() == "" or val == "Sin información":
        return "Sin información"
    try:
        num = float(val)
        if num == 0:
            return "$ 0"
        return f"$ {int(round(num)):,}".replace(",", ".")
    except (ValueError, TypeError):
        return str(val)

def parse_excel_inventory(file_path):
    """
    Lee y procesa todas las hojas del Excel respetando encabezados,
    estructuras independientes, listas de precios, existencias y columnas adicionales.
    """
    if not file_path or not os.path.exists(file_path):
        raise FileNotFoundError(f"Archivo no encontrado: {file_path}")

    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    all_sheets = wb.sheetnames

    def find_sheet(target):
        target_norm = target.strip().lower()
        for name in all_sheets:
            if name.strip().lower() == target_norm:
                return name
        return None

    products = []

    # =========================================================================
    # 1. PESTAÑA: LLANTAS
    # =========================================================================
    s_llantas = find_sheet('LLANTAS')
    if s_llantas:
        ws = wb[s_llantas]
        r8 = list(ws.iter_rows(min_row=8, max_row=8, values_only=True))[0]
        r9 = list(ws.iter_rows(min_row=9, max_row=9, values_only=True))[0]
        r10 = list(ws.iter_rows(min_row=10, max_row=10, values_only=True))[0]

        # Mapeo de columnas BQ:CD (índices 68 a 81)
        extra_headers_llantas = {}
        for c in range(68, 82):
            date_part = normalize_text(r8[c]) if c < len(r8) else ""
            brand_part = normalize_text(r9[c]) if c < len(r9) else ""
            code_part = normalize_text(r10[c]) if c < len(r10) else ""
            col_letter = openpyxl.utils.get_column_letter(c+1)
            parts = [p for p in [code_part, brand_part, date_part] if p]
            label = " - ".join(parts) if parts else f"Col {col_letter}"
            extra_headers_llantas[c] = {
                'col_letter': col_letter,
                'label': label,
                'code': code_part,
                'brand': brand_part,
                'date': date_part
            }

        physical_sedes_map = {
            9: "BOD TITAN", 10: "ALMAG", 11: "PASAMAR", 12: "CALI",
            13: "PEREIRA", 14: "PASTO", 15: "BOGOTA", 16: "VILLAVICENCIO",
            17: "MEDELLIN", 18: "CARTAGENA", 19: "BARRANQUILLA", 20: "BUCARAMANGA"
        }

        consign_map = {}
        for c in range(23, 61):
            name = normalize_text(r9[c]) if c < len(r9) else ""
            if name:
                consign_map[c] = name

        for row_idx, row in enumerate(ws.iter_rows(min_row=11, values_only=True), start=11):
            cg1 = clean_code(row[1]) if len(row) > 1 else ""
            ref = normalize_text(row[2]) if len(row) > 2 else ""
            if not cg1 and not ref:
                continue

            linea = normalize_text(row[0]) if len(row) > 0 else ""
            pr = clean_code(row[3]) if len(row) > 3 else ""
            marca = normalize_text(row[4]) if len(row) > 4 else ""
            diseno = normalize_text(row[5]) if len(row) > 5 else ""
            tipo = normalize_text(row[6]) if len(row) > 6 else ""
            tt = normalize_text(row[7]) if len(row) > 7 else ""
            aplica_iva = normalize_text(row[8]) if len(row) > 8 else ""

            # Sedes físicas con stock > 0
            sedes_stock = {}
            total_fisico = 0
            for col_idx, sede_name in physical_sedes_map.items():
                if col_idx < len(row):
                    val = row[col_idx]
                    if val is not None and isinstance(val, (int, float)) and val > 0:
                        qty = int(val) if val == int(val) else val
                        sedes_stock[sede_name] = qty
                        total_fisico += qty

            # Consignaciones
            consignaciones = {}
            for col_idx, c_name in consign_map.items():
                if col_idx < len(row):
                    val = row[col_idx]
                    if val is not None and isinstance(val, (int, float)) and val > 0:
                        consignaciones[c_name] = int(val) if val == int(val) else val

            comprometido = 0
            if len(row) > 21 and row[21] is not None and isinstance(row[21], (int, float)):
                comprometido = int(row[21]) if row[21] == int(row[21]) else row[21]

            transito = 0
            if len(row) > 61 and row[61] is not None and isinstance(row[61], (int, float)):
                transito = int(row[61]) if row[61] == int(row[61]) else row[61]

            pedido_fabr = 0
            if len(row) > 62 and row[62] is not None and isinstance(row[62], (int, float)):
                pedido_fabr = int(row[62]) if row[62] == int(row[62]) else row[62]

            # Precios en LLANTAS: BL=63 Cliente, BM=64 Distribuidor, BN=65 Oferta, BO=66 Oferta con IVA, BP=67 Bono
            p_cli = row[63] if len(row) > 63 else None
            p_dist = row[64] if len(row) > 64 else None
            p_of = row[65] if len(row) > 65 else None
            p_of_iva = row[66] if len(row) > 66 else None
            bono = row[67] if len(row) > 67 else None

            iva_multiplier = 1.19 if aplica_iva.upper() == "SI" else 1.0
            p_cli_con_iva = (p_cli * iva_multiplier) if (p_cli is not None and isinstance(p_cli, (int, float))) else None
            p_dist_con_iva = (p_dist * iva_multiplier) if (p_dist is not None and isinstance(p_dist, (int, float))) else None

            info_adicional = []
            for col_idx, meta in extra_headers_llantas.items():
                if col_idx < len(row):
                    val = row[col_idx]
                    if val is not None and str(val).strip() != "" and str(val).strip() != "0":
                        info_adicional.append({
                            'col': meta['col_letter'],
                            'label': meta['label'],
                            'code': meta['code'],
                            'brand': meta['brand'],
                            'date': meta['date'],
                            'value': val
                        })

            products.append({
                'source_sheet': 'LLANTAS',
                'row_index': row_idx,
                'cod_cg1': cg1,
                'referencia': ref,
                'descripcion': f"{ref} {marca} {diseno}".strip(),
                'marca': marca or "Sin información",
                'diseno': diseno or "Sin información",
                'tipo': tipo or "Sin información",
                'tipo_valvula': tt or "Sin información",
                'pr': pr or "Sin información",
                'aplica_iva': aplica_iva or "Sin información",
                'cant_caja': "Sin información",
                'linea': linea or "Sin información",
                'precios': {
                    'precio_cliente': p_cli,
                    'precio_cliente_iva': p_cli_con_iva,
                    'precio_distribuidor': p_dist,
                    'precio_distribuidor_iva': p_dist_con_iva,
                    'precio_oferta': p_of,
                    'precio_oferta_iva': p_of_iva,
                    'bono': bono
                },
                'stock_fisico': sedes_stock,
                'total_fisico': total_fisico,
                'comprometido': comprometido,
                'consignaciones': consignaciones,
                'transito': transito,
                'pedido_fabrica': pedido_fabr,
                'info_adicional': info_adicional,
                'fabricante_code': ""
            })

    # =========================================================================
    # 2. PESTAÑAS: NEU ESPEC & NEU COMERC
    # =========================================================================
    for s_name in ['NEU ESPEC', 'NEU COMERC']:
        real_sheet = find_sheet(s_name)
        if not real_sheet:
            continue
        ws = wb[real_sheet]
        r1 = list(ws.iter_rows(min_row=1, max_row=1, values_only=True))[0]
        r2 = list(ws.iter_rows(min_row=2, max_row=2, values_only=True))[0]
        r3 = list(ws.iter_rows(min_row=3, max_row=3, values_only=True))[0]

        neu_sedes_map = {
            6: "CALI", 7: "PEREIRA", 8: "PASTO", 9: "BOGOTA",
            10: "VILLAVICENCIO", 11: "MEDELLIN", 12: "CARTAGENA",
            13: "BARRANQUILLA", 14: "BUCARAMANGA"
        }

        neu_consign_map = {}
        for c in range(15, 54):
            if c < len(r2) and r2[c]:
                neu_consign_map[c] = normalize_text(r2[c])

        order_cols = {}
        for c in range(55, 59):
            if c < len(r2) and r2[c]:
                ord_name = normalize_text(r2[c])
                brand = normalize_text(r3[c]) if c < len(r3) else ""
                dt = normalize_text(r1[c]) if c < len(r1) else ""
                order_cols[c] = f"{ord_name} ({brand} - {dt})".strip()

        for row_idx, row in enumerate(ws.iter_rows(min_row=5, values_only=True), start=5):
            cg1 = clean_code(row[0]) if len(row) > 0 else ""
            ref = normalize_text(row[1]) if len(row) > 1 else ""
            if not cg1 and not ref:
                continue

            valvula = normalize_text(row[2]) if len(row) > 2 else ""
            tr = normalize_text(row[3]) if len(row) > 3 else ""
            marca = normalize_text(row[4]) if len(row) > 4 else ""
            cant_caja = clean_code(row[5]) if len(row) > 5 else ""

            sedes_stock = {}
            total_fisico = 0
            for col_idx, s_name_lbl in neu_sedes_map.items():
                if col_idx < len(row):
                    val = row[col_idx]
                    if val is not None and isinstance(val, (int, float)) and val > 0:
                        qty = int(val) if val == int(val) else val
                        sedes_stock[s_name_lbl] = qty
                        total_fisico += qty

            consignaciones = {}
            for col_idx, c_name in neu_consign_map.items():
                if col_idx < len(row):
                    val = row[col_idx]
                    if val is not None and isinstance(val, (int, float)) and val > 0:
                        consignaciones[c_name] = int(val) if val == int(val) else val

            pedido_fab = 0
            if len(row) > 54 and row[54] is not None and isinstance(row[54], (int, float)):
                pedido_fab = int(row[54]) if row[54] == int(row[54]) else row[54]

            # En NEU ESPEC / NEU COMERC no hay columnas BQ:CD
            info_adicional = []

            # Precios: BH=59 Cliente Sin IVA, BI=60 Con IVA, BJ=61 Dist Sin IVA, BK=62 Con IVA, BL=63 Oferta Sin IVA, BM=64 Con IVA
            p_cli_sin = row[59] if len(row) > 59 else None
            p_cli_con = row[60] if len(row) > 60 else None
            p_dist_sin = row[61] if len(row) > 61 else None
            p_dist_con = row[62] if len(row) > 62 else None
            p_of_sin = row[63] if len(row) > 63 else None
            p_of_con = row[64] if len(row) > 64 else None

            products.append({
                'source_sheet': s_name,
                'row_index': row_idx,
                'cod_cg1': cg1,
                'referencia': ref,
                'descripcion': f"{ref} {marca}".strip(),
                'marca': marca or "Sin información",
                'diseno': "Sin información",
                'tipo': s_name,
                'tipo_valvula': f"{valvula} {tr}".strip() or "Sin información",
                'pr': "Sin información",
                'aplica_iva': "SI",
                'cant_caja': cant_caja or "Sin información",
                'linea': s_name,
                'precios': {
                    'precio_cliente': p_cli_sin,
                    'precio_cliente_iva': p_cli_con,
                    'precio_distribuidor': p_dist_sin,
                    'precio_distribuidor_iva': p_dist_con,
                    'precio_oferta': p_of_sin,
                    'precio_oferta_iva': p_of_con,
                    'bono': None
                },
                'stock_fisico': sedes_stock,
                'total_fisico': total_fisico,
                'comprometido': 0,
                'consignaciones': consignaciones,
                'transito': 0,
                'pedido_fabrica': pedido_fab,
                'info_adicional': info_adicional,
                'fabricante_code': ""
            })

    # =========================================================================
    # 3. PESTAÑA: LP GY AUTO-CAM
    # =========================================================================
    s_lp_gy_ac = find_sheet('LP GY AUTO-CAM')
    if s_lp_gy_ac:
        ws = wb[s_lp_gy_ac]
        gy_ac_sedes = {3: "CALI", 4: "BOGOTA", 5: "VILLAVICENCIO", 6: "MEDELLIN", 7: "BARRANQUILLA"}
        for row_idx, row in enumerate(ws.iter_rows(min_row=5, values_only=True), start=5):
            cod_gy = clean_code(row[0]) if len(row) > 0 else ""
            cg1 = clean_code(row[1]) if len(row) > 1 else ""
            desc = normalize_text(row[2]) if len(row) > 2 else ""
            if not cod_gy and not cg1 and not desc:
                continue

            sedes_stock = {}
            total_fisico = 0
            for col_idx, s_name_lbl in gy_ac_sedes.items():
                if col_idx < len(row):
                    val = row[col_idx]
                    if val is not None and isinstance(val, (int, float)) and val > 0:
                        qty = int(val) if val == int(val) else val
                        sedes_stock[s_name_lbl] = qty
                        total_fisico += qty

            products.append({
                'source_sheet': 'LP GY AUTO-CAM',
                'row_index': row_idx,
                'cod_cg1': cg1,
                'referencia': desc,
                'descripcion': desc,
                'marca': "GOODYEAR",
                'diseno': "Sin información",
                'tipo': "Auto-Camioneta",
                'tipo_valvula': "Sin información",
                'pr': "Sin información",
                'aplica_iva': "SI",
                'cant_caja': "Sin información",
                'linea': "GOODYEAR AUTO-CAMIONETA",
                'precios': {
                    'precio_cliente': row[8] if len(row) > 8 else None,
                    'precio_cliente_iva': row[9] if len(row) > 9 else None,
                    'precio_distribuidor': row[10] if len(row) > 10 else None,
                    'precio_distribuidor_iva': row[11] if len(row) > 11 else None,
                    'precio_ingenio': row[12] if len(row) > 12 else None,
                    'precio_oferta': row[14] if len(row) > 14 else None,
                    'precio_oferta_iva': row[15] if len(row) > 15 else None,
                    'bono': None
                },
                'stock_fisico': sedes_stock,
                'total_fisico': total_fisico,
                'comprometido': 0,
                'consignaciones': {},
                'transito': 0,
                'pedido_fabrica': 0,
                'info_adicional': [],
                'fabricante_code': cod_gy
            })

    # =========================================================================
    # 4. PESTAÑA: LP CONTI CAMION
    # =========================================================================
    s_lp_conti = find_sheet('LP CONTI CAMION')
    if s_lp_conti:
        ws = wb[s_lp_conti]
        conti_sedes = {8: "CALI", 9: "BOGOTA"}
        for row_idx, row in enumerate(ws.iter_rows(min_row=3, values_only=True), start=3):
            material = clean_code(row[0]) if len(row) > 0 else ""
            cg1 = clean_code(row[1]) if len(row) > 1 else ""
            segmento = normalize_text(row[2]) if len(row) > 2 else ""
            eje = normalize_text(row[3]) if len(row) > 3 else ""
            desc = normalize_text(row[4]) if len(row) > 4 else ""
            marca = normalize_text(row[6]) if len(row) > 6 else "CONTINENTAL"
            diseno = normalize_text(row[7]) if len(row) > 7 else ""
            if not material and not cg1 and not desc:
                continue

            sedes_stock = {}
            total_fisico = 0
            for col_idx, s_name_lbl in conti_sedes.items():
                if col_idx < len(row):
                    val = row[col_idx]
                    if val is not None and isinstance(val, (int, float)) and val > 0:
                        qty = int(val) if val == int(val) else val
                        sedes_stock[s_name_lbl] = qty
                        total_fisico += qty

            products.append({
                'source_sheet': 'LP CONTI CAMION',
                'row_index': row_idx,
                'cod_cg1': cg1 if cg1 != "0" else "",
                'referencia': desc,
                'descripcion': f"{desc} {diseno}".strip(),
                'marca': marca,
                'diseno': diseno or "Sin información",
                'tipo': f"{segmento} {eje}".strip() or "Camión",
                'tipo_valvula': "Sin información",
                'pr': "Sin información",
                'aplica_iva': "SI",
                'cant_caja': "Sin información",
                'linea': "CONTINENTAL CAMION",
                'precios': {
                    'precio_cliente': row[10] if len(row) > 10 else None,
                    'precio_cliente_iva': row[11] if len(row) > 11 else None,
                    'precio_distribuidor': row[12] if len(row) > 12 else None,
                    'precio_distribuidor_iva': row[13] if len(row) > 13 else None,
                    'precio_oferta': row[14] if len(row) > 14 else None,
                    'precio_oferta_iva': row[15] if len(row) > 15 else None,
                    'bono': None
                },
                'stock_fisico': sedes_stock,
                'total_fisico': total_fisico,
                'comprometido': 0,
                'consignaciones': {},
                'transito': 0,
                'pedido_fabrica': 0,
                'info_adicional': [],
                'fabricante_code': material
            })

    # =========================================================================
    # 5. PESTAÑA: LP GY CAMION
    # =========================================================================
    s_lp_gy_cam = find_sheet('LP GY CAMION')
    if s_lp_gy_cam:
        ws = wb[s_lp_gy_cam]
        gy_cam_sedes = {4: "CALI", 5: "BOGOTA", 6: "VILLAVICENCIO", 7: "MEDELLIN"}
        for row_idx, row in enumerate(ws.iter_rows(min_row=5, values_only=True), start=5):
            cod_gy = clean_code(row[0]) if len(row) > 0 else ""
            cg1 = clean_code(row[1]) if len(row) > 1 else ""
            pos = normalize_text(row[2]) if len(row) > 2 else ""
            desc = normalize_text(row[3]) if len(row) > 3 else ""
            if not cod_gy and not cg1 and not desc:
                continue

            if "GOODYEAR" in cod_gy or "CÓDIGO" in cod_gy or "PRECIO" in str(row[8]):
                continue

            sedes_stock = {}
            total_fisico = 0
            for col_idx, s_name_lbl in gy_cam_sedes.items():
                if col_idx < len(row):
                    val = row[col_idx]
                    if val is not None and isinstance(val, (int, float)) and val > 0:
                        qty = int(val) if val == int(val) else val
                        sedes_stock[s_name_lbl] = qty
                        total_fisico += qty

            products.append({
                'source_sheet': 'LP GY CAMION',
                'row_index': row_idx,
                'cod_cg1': cg1,
                'referencia': desc,
                'descripcion': f"{desc} {pos}".strip(),
                'marca': "GOODYEAR",
                'diseno': pos or "Sin información",
                'tipo': pos or "Camión",
                'tipo_valvula': "Sin información",
                'pr': "Sin información",
                'aplica_iva': "SI",
                'cant_caja': "Sin información",
                'linea': "GOODYEAR CAMION",
                'precios': {
                    'precio_cliente': row[8] if len(row) > 8 else None,
                    'precio_cliente_iva': row[9] if len(row) > 9 else None,
                    'precio_distribuidor': row[10] if len(row) > 10 else None,
                    'precio_distribuidor_iva': row[11] if len(row) > 11 else None,
                    'precio_ingenio': row[12] if len(row) > 12 else None,
                    'precio_oferta': row[14] if len(row) > 14 else None,
                    'bono': None
                },
                'stock_fisico': sedes_stock,
                'total_fisico': total_fisico,
                'comprometido': 0,
                'consignaciones': {},
                'transito': 0,
                'pedido_fabrica': 0,
                'info_adicional': [],
                'fabricante_code': cod_gy
            })

    # =========================================================================
    # 6. PESTAÑA: LP CONTI AUTO-CAM
    # =========================================================================
    s_lp_conti_ac = find_sheet('LP CONTI AUTO-CAM ')
    if s_lp_conti_ac:
        ws = wb[s_lp_conti_ac]
        conti_ac_sedes = {6: "CALI", 7: "BOGOTA"}
        for row_idx, row in enumerate(ws.iter_rows(min_row=6, values_only=True), start=6):
            mat = clean_code(row[0]) if len(row) > 0 else ""
            cg1 = clean_code(row[1]) if len(row) > 1 else ""
            desc = normalize_text(row[2]) if len(row) > 2 else ""
            marca = normalize_text(row[3]) if len(row) > 3 else "CONTINENTAL"
            medida = normalize_text(row[4]) if len(row) > 4 else ""
            diseno = normalize_text(row[5]) if len(row) > 5 else ""
            if not mat and not cg1 and not desc:
                continue

            sedes_stock = {}
            total_fisico = 0
            for col_idx, s_name_lbl in conti_ac_sedes.items():
                if col_idx < len(row):
                    val = row[col_idx]
                    if val is not None and isinstance(val, (int, float)) and val > 0:
                        qty = int(val) if val == int(val) else val
                        sedes_stock[s_name_lbl] = qty
                        total_fisico += qty

            products.append({
                'source_sheet': 'LP CONTI AUTO-CAM',
                'row_index': row_idx,
                'cod_cg1': cg1,
                'referencia': desc or medida,
                'descripcion': f"{desc} {diseno}".strip(),
                'marca': marca,
                'diseno': diseno or "Sin información",
                'tipo': "Auto-Camioneta",
                'tipo_valvula': "Sin información",
                'pr': "Sin información",
                'aplica_iva': "SI",
                'cant_caja': "Sin información",
                'linea': f"{marca} AUTO-CAMIONETA",
                'precios': {
                    'precio_cliente': row[8] if len(row) > 8 else None,
                    'precio_cliente_iva': row[9] if len(row) > 9 else None,
                    'precio_distribuidor': row[10] if len(row) > 10 else None,
                    'precio_distribuidor_iva': row[11] if len(row) > 11 else None,
                    'precio_ingenio': row[12] if len(row) > 12 else None,
                    'precio_oferta': None,
                    'bono': None
                },
                'stock_fisico': sedes_stock,
                'total_fisico': total_fisico,
                'comprometido': 0,
                'consignaciones': {},
                'transito': 0,
                'pedido_fabrica': 0,
                'info_adicional': [],
                'fabricante_code': mat
            })

    # =========================================================================
    # 7. OTRAS HOJAS: E-S, OFERTA, OBSOL
    # =========================================================================
    s_es = find_sheet('E-S')
    if s_es:
        ws = wb[s_es]
        es_sedes = {2: "CALI", 3: "BOGOTA", 4: "VILLAVICENCIO", 5: "MEDELLIN", 6: "BARRANQUILLA"}
        for row_idx, row in enumerate(ws.iter_rows(min_row=5, values_only=True), start=5):
            cg1 = clean_code(row[0]) if len(row) > 0 else ""
            ref = normalize_text(row[1]) if len(row) > 1 else ""
            if not cg1 and not ref:
                continue

            sedes_stock = {}
            total_fisico = 0
            for col_idx, s_name_lbl in es_sedes.items():
                if col_idx < len(row):
                    val = row[col_idx]
                    if val is not None and isinstance(val, (int, float)) and val > 0:
                        qty = int(val) if val == int(val) else val
                        sedes_stock[s_name_lbl] = qty
                        total_fisico += qty

            products.append({
                'source_sheet': 'E-S',
                'row_index': row_idx,
                'cod_cg1': cg1,
                'referencia': ref,
                'descripcion': ref,
                'marca': "Sin información",
                'diseno': "Sin información",
                'tipo': "Entradas y Salidas / Saldos",
                'tipo_valvula': "Sin información",
                'pr': "Sin información",
                'aplica_iva': "SI",
                'cant_caja': "Sin información",
                'linea': "E-S",
                'precios': {
                    'precio_cliente': row[7] if len(row) > 7 else None,
                    'precio_cliente_iva': row[8] if len(row) > 8 else None,
                    'precio_distribuidor': row[7] if len(row) > 7 else None,
                    'precio_distribuidor_iva': row[8] if len(row) > 8 else None,
                    'precio_oferta': None,
                    'bono': None
                },
                'stock_fisico': sedes_stock,
                'total_fisico': total_fisico,
                'comprometido': 0,
                'consignaciones': {},
                'transito': 0,
                'pedido_fabrica': 0,
                'info_adicional': [],
                'fabricante_code': ""
            })

    s_of = find_sheet('OFERTA')
    if s_of:
        ws = wb[s_of]
        for row_idx, row in enumerate(ws.iter_rows(min_row=10, values_only=True), start=10):
            cg1 = clean_code(row[1]) if len(row) > 1 else ""
            ref = normalize_text(row[2]) if len(row) > 2 else ""
            if not cg1 and not ref:
                continue

            products.append({
                'source_sheet': 'OFERTA',
                'row_index': row_idx,
                'cod_cg1': cg1,
                'referencia': ref,
                'descripcion': f"{ref} {normalize_text(row[4])} {normalize_text(row[5])}".strip(),
                'marca': normalize_text(row[4]) or "Sin información",
                'diseno': normalize_text(row[5]) or "Sin información",
                'tipo': normalize_text(row[6]) or "Oferta Especial",
                'tipo_valvula': normalize_text(row[7]) or "Sin información",
                'pr': clean_code(row[3]) or "Sin información",
                'aplica_iva': normalize_text(row[8]) or "Sin información",
                'cant_caja': "Sin información",
                'linea': normalize_text(row[0]) or "OFERTA",
                'precios': {
                    'precio_cliente': None,
                    'precio_distribuidor': None,
                    'precio_oferta': row[9] if len(row) > 9 else None,
                    'precio_oferta_iva': row[10] if len(row) > 10 else None,
                    'bono': row[11] if len(row) > 11 else None
                },
                'stock_fisico': {},
                'total_fisico': 0,
                'comprometido': 0,
                'consignaciones': {},
                'transito': 0,
                'pedido_fabrica': 0,
                'info_adicional': [],
                'fabricante_code': ""
            })

    s_obsol = find_sheet('OBSOL')
    if s_obsol:
        ws = wb[s_obsol]
        obsol_sedes = {5: "CALI", 6: "MEDELLIN", 7: "CARTAGENA", 8: "BARRANQUILLA", 9: "BOGOTA", 10: "VILLAVICENCIO"}
        for row_idx, row in enumerate(ws.iter_rows(min_row=13, values_only=True), start=13):
            cg1 = clean_code(row[0]) if len(row) > 0 else ""
            ref = normalize_text(row[1]) if len(row) > 1 else ""
            if not cg1 and not ref:
                continue

            sedes_stock = {}
            total_fisico = 0
            for col_idx, s_name_lbl in obsol_sedes.items():
                if col_idx < len(row):
                    val = row[col_idx]
                    if val is not None and isinstance(val, (int, float)) and val > 0:
                        qty = int(val) if val == int(val) else val
                        sedes_stock[s_name_lbl] = qty
                        total_fisico += qty

            products.append({
                'source_sheet': 'OBSOL',
                'row_index': row_idx,
                'cod_cg1': cg1,
                'referencia': ref,
                'descripcion': f"{ref} {normalize_text(row[3])}".strip(),
                'marca': normalize_text(row[3]) or "Sin información",
                'diseno': "Sin información",
                'tipo': normalize_text(row[4]) or "Obsolescencia",
                'tipo_valvula': "Sin información",
                'pr': clean_code(row[2]) or "Sin información",
                'aplica_iva': "SI",
                'cant_caja': "Sin información",
                'linea': "OBSOLETOS",
                'precios': {
                    'precio_cliente': None,
                    'precio_distribuidor': None,
                    'precio_oferta': row[12] if len(row) > 12 else None,
                    'precio_oferta_iva': row[13] if len(row) > 13 else None,
                    'bono': row[14] if len(row) > 14 else None
                },
                'stock_fisico': sedes_stock,
                'total_fisico': total_fisico,
                'comprometido': 0,
                'consignaciones': {},
                'transito': 0,
                'pedido_fabrica': 0,
                'info_adicional': [],
                'fabricante_code': ""
            })

    # =========================================================================
    # 8. HOJAS AUXILIARES DE PROVEEDORES (TITAN, CONTINENTAL, GOODYEAR)
    # =========================================================================
    supplier_inventory = []

    s_inv_titan = find_sheet('INV TITAN OCT 1')
    if s_inv_titan:
        ws = wb[s_inv_titan]
        for row in ws.iter_rows(min_row=5, values_only=True):
            cod_titan = clean_code(row[0]) if len(row) > 0 else ""
            cod_agro = clean_code(row[1]) if len(row) > 1 else ""
            desc = normalize_text(row[2]) if len(row) > 2 else ""
            inv = row[9] if len(row) > 9 else 0
            if (cod_titan or cod_agro or desc) and inv:
                supplier_inventory.append({
                    'proveedor': 'TITAN USA / CARLSTAR',
                    'cod_fabricante': cod_titan,
                    'cod_cg1': cod_agro if cod_agro != "-" else "",
                    'referencia': desc,
                    'disponible': inv,
                    'transito': 0,
                    'eta': "Disponible en planta/bodega proveedor"
                })

    s_inv_cst = find_sheet('INV CST CONTI OCT 22') or find_sheet('INV CST CONTI OCT 7')
    if s_inv_cst:
        ws = wb[s_inv_cst]
        for row in ws.iter_rows(min_row=9, values_only=True):
            cod = clean_code(row[2]) if len(row) > 2 else ""
            desc = normalize_text(row[3]) if len(row) > 3 else ""
            inv_cobo = row[10] if len(row) > 10 else 0
            trans_oct = row[11] if len(row) > 11 else 0
            eta = normalize_text(row[12]) if len(row) > 12 else ""
            if cod or desc:
                supplier_inventory.append({
                    'proveedor': 'CONTINENTAL CST',
                    'cod_fabricante': cod,
                    'cod_cg1': "",
                    'referencia': desc,
                    'disponible': inv_cobo or 0,
                    'transito': trans_oct or 0,
                    'eta': eta or "Consultar fecha"
                })

    s_inv_gy = find_sheet('INV GY FEB 10')
    if s_inv_gy:
        ws = wb[s_inv_gy]
        for row in ws.iter_rows(min_row=5, values_only=True):
            mat1 = clean_code(row[0]) if len(row) > 0 else ""
            desc1 = normalize_text(row[1]) if len(row) > 1 else ""
            cant1 = normalize_text(row[3]) if len(row) > 3 else ""
            if mat1 or desc1:
                supplier_inventory.append({
                    'proveedor': 'GOODYEAR',
                    'cod_fabricante': mat1,
                    'cod_cg1': "",
                    'referencia': desc1,
                    'disponible': cant1 or "En inventario planta",
                    'transito': 0,
                    'eta': "Inmediata"
                })

    wb.close()
    return products, supplier_inventory
