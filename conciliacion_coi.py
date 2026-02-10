import pandas as pd
import numpy as np
import re
import math

# --- CONFIGURACIÓN ---
FILE_ODOO = 'Reporte_Contable_Final.xlsx'
FILE_COI = 'COI_Final_SumaCorrecta.xlsx'
FILE_OUTPUT = 'Analisis_Comparativo_Diciembre_V18_7.xlsx'

# --- 1. DICCIONARIO DE CONTROL PARA CHECK ABUELAS ---
CHECK_ABUELAS_LIST = [
    '1110-000-000', '1120-000-000', '1121-000-000', '1122-000-000',
    '1140-000-000', '1150-000-000', '1170-000-000', '1180-000-000',
    '1190-000-000', '1200-000-000', '1201-000-000', '1215-000-000',
    '1310-006-000','1360-000-000','1360-002-000', '2150-000-000', '2160-000-000',
    '2170-000-000','2180-000-000', '1310-004-000', '2130-000-000', '4100-000-000',
    '6200-000-000','7200-000-000',
    '2110-000-000', '2120-000-000', '2190-000-000', 'SUMA-BANCOS-TOTAL',
    'SUMA-CLIENTES-NACIONALES', 'SUMA-CLIENTES-EXTRANJEROS'
]

# --- 2. MAPA MAESTRO ---
HEADER_MAP = {
    '101': '1110-000-000', '102': 'SUMA-BANCOS-TOTAL', '102.01': '1120-000-000', 
    '102.02.01':'1121-001-000',
    '102.02':'SUMA-BANCOS-USD','105.01.00':'1150-001-000', '102.02.02':'1122-000-000',
    '104': '1140-000-000',
    '105': '1150-000-000', '105.01': 'SUMA-CLIENTES-NACIONALES', '105.02': 'SUMA-CLIENTES-EXTRANJEROS',
    '107': '1170-000-000', '107.02': '1170-002-000','109': '1210-000-000', '113': '1180-000-000', 
    '115': '1190-000-000', '118': '1200-000-000', '119': '1201-000-000', '154':'1310-003-000','155':'1310-005-000',
    '120': '1215-000-000', '120.02.01':'1215-002-000','114': '1220-000-000', '153': '1310-006-000',
    '156':'1310-004-000', '160':'1310-007-000', 
    '171': '1360-000-000',  '171.03': '1360-002-000','201': '2110-000-000', '201.01':'2110-001-000',
    '201.03.01':'2115-000-000', '205.02.06' :'2120-001-013', 
    '205': '2120-000-000', '205.02': '2120-001-000','205.02.01':'2120-001-001','201.01.01':'2110-001-001',
    '205.02.09':'2120-001-019', 
    '206': '2190-000-000', '206.01.01': '2190-001-000', '213': '2140-000-000', '216': '2150-000-000', 
    '210': '2160-000-000', '211': '2170-000-000', '208': '2180-000-000', 
    '209': '2181-000-000', '251': '2130-000-000', '401': '4100-000-000', '401.04.01':'4100-002-000',
    '402': '4200-000-000', '402.02.01':'4200-002-000','501': '5000-000-000',  '501.08':'5100-000-000', '501.08.08':'5200-000-000','602': '6100-000-000', 
    '603': '6200-000-000', '603.82': '6200-055-000', '604.59': '6200-034-000',
    '702': '7100-000-000', '701': '7200-000-000',  '701.05': '7200-005-000',
    '704': '7300-000-000', '703': '7400-000-000'
}

COI_TO_ODOO_SECTION = {
    '1110': '101', '1120': '102', '1140': '104', '1150': '105', '1170': '107',
    '1180': '113', '1190': '115', '1191': '115', '1192': '115','2180':'205',
    '2120':'205', '2115':'205', '2181':'205', '2190':'205', '2160':'205',
    '1200': '118', '1201': '119', '1210': '109', '1215': '120', '1220': '114',
    '1310': '153', '1360': '171', '2110': '201', '2130': '251', '1460':'183',
    '4100': '401', '4200': '402', '5000': '501', '5100': '501', '5200': '501',
    '6100': '602', '6200': '603', '7100': '702', '7200': '701', '7300': '704', '7400': '703'
}

# --- 3. SUMAS VIRTUALES ---
VIRTUAL_COI_SUMS = {
    'SUMA-BANCOS-TOTAL': ['1120-000-000', '1121-000-000', '1122-000-000'],
    'SUMA-BANCOS-USD': ['1121-001-000', '1122-000-000'],
    'SUMA-CLIENTES-NACIONALES': ['1150-001-000', '1150-004-000', '1150-005-000', '1150-006-000'],
    'SUMA-CLIENTES-EXTRANJEROS': ['1150-002-000', '1150-003-000'],
    'SUMA-PROVEEDORES-NACIONALES': ['1150-002-000', '2115-000-000']
}

# --- HELPERS ---
def normalize_code(code):
    if not code or pd.isna(code): return ""
    s = str(code).upper().strip()
    if s.startswith("SUMA-") or s.startswith("COI-"): return s
    return s.replace('.', '').replace('-', '')

def extract_key(desc):
    if not desc or pd.isna(desc): return None
    match = re.search(r'(\d{4}[-\.]\d{3}[-\.]\d{3})', str(desc))
    if match: return match.group(1).replace('.', '-').strip()
    return None

def clean_money(val):
    if pd.isna(val) or val == '': return 0.0
    try: return float(str(val).replace('$','').replace(',','').replace(' ',''))
    except: return 0.0

def is_abuela_format(key):
    if not key: return False
    return normalize_code(key) in [normalize_code(x) for x in CHECK_ABUELAS_LIST]

def safe_write_money(ws, row, col, val, fmt):
    if val is None or pd.isna(val) or (isinstance(val, float) and (math.isnan(val) or math.isinf(val))):
        ws.write(row, col, "", fmt)
    else:
        ws.write(row, col, float(val), fmt)

# --- PROCESO ---
def generar_analisis_v18_7():
    print("--- Ejecutando Versión 18.7: Ajuste de Sumas Virtuales y Estatus Estructural ---")
    df_odoo = pd.read_excel(FILE_ODOO, engine='openpyxl').fillna('')
    df_coi = pd.read_excel(FILE_COI, engine='openpyxl').fillna('')

    df_coi['Saldo'] = df_coi['Saldo'].apply(clean_money)
    df_coi['Cuenta'] = df_coi['Cuenta'].astype(str).str.strip()
    
    coi_lookup = {}
    for _, r in df_coi[df_coi['Cuenta'] != ''].iterrows():
        norm = normalize_code(r['Cuenta'])
        coi_lookup[norm] = {'Cuenta_Orig': r['Cuenta'], 'Descripcion': r['Descripcion'], 'Saldo': r['Saldo']}

    for v_key, comps in VIRTUAL_COI_SUMS.items():
        total = sum((coi_lookup.get(normalize_code(c), {}).get('Saldo', 0.0)) for c in comps)
        coi_lookup[normalize_code(v_key)] = {'Cuenta_Orig': v_key, 'Descripcion': f"GRUPO {v_key}", 'Saldo': total}

    coi_restante = {k: v['Saldo'] for k, v in coi_lookup.items()}
    df_odoo['Saldo_L'] = df_odoo['Saldo'].apply(clean_money)
    df_odoo['Cta_S'] = df_odoo['Cuenta'].astype(str).str.strip()
    
    rows_final = []
    cuentas_coi_usadas = set()
    anchors = {}

    for idx, row in df_odoo.iterrows():
        cta_o = row['Cta_S']
        desc_o = str(row['Descripcion']).strip()
        saldo_o = row['Saldo_L']

        if not desc_o or desc_o == "" or (saldo_o == 0 and cta_o == ""):
            continue
        if "recibo" in desc_o.lower() and "pendiente" in desc_o.lower(): continue
        if desc_o.startswith("Suma") and cta_o not in HEADER_MAP and not cta_o.startswith("206"): continue

        target_map = HEADER_MAP.get(cta_o)
        extracted = extract_key(row['Descripcion'])
        key_norm = normalize_code(target_map) if target_map else normalize_code(extracted)
        is_h = cta_o in HEADER_MAP

        for pref_c, cta_ref in COI_TO_ODOO_SECTION.items():
            if cta_o == cta_ref: anchors[pref_c] = float(idx) + 0.6

        # Lógica de estatus por estructura
        status_init = "NO EN COI"
        if not target_map and not extracted:
            status_init = "NO EN COI POR ESTRUCTURA"

        item = {
            'Orden': float(idx), 'Odoo_Cta': cta_o, 'Odoo_Desc': row['Descripcion'], 
            'Odoo_Saldo': row['Saldo_L'], 'COI_Cta': target_map or "", 'COI_Desc': '', 'COI_Saldo': None,
            'Diff': None, 'Status': status_init, 'Is_Header': is_h
        }
        
        if key_norm in coi_lookup:
            c = coi_lookup[key_norm]
            item['COI_Cta'], item['COI_Desc'], item['COI_Saldo'] = c['Cuenta_Orig'], c['Descripcion'], coi_restante.get(key_norm)
            coi_restante[key_norm] = None 
            if not is_h and not str(key_norm).startswith("SUMA"): cuentas_coi_usadas.add(key_norm)
            diff = abs(item['Odoo_Saldo'] or 0.0) - abs(item['COI_Saldo'] or 0.0)
            item['Diff'] = diff if abs(diff) > 0.01 else None
            item['Status'] = "OK" if abs(diff) < 0.1 else "DIFERENCIA"
        
        rows_final.append(item)

    # Inserción de huérfanas (Modificado para que las SUMA no se oculten)
    for norm_key, data in coi_lookup.items():
        if norm_key not in cuentas_coi_usadas and not any(normalize_code(v) == norm_key for v in HEADER_MAP.values()):
            if abs(data['Saldo']) > 0.01:
                prefix = normalize_code(data['Cuenta_Orig'])[:4]
                pos = anchors.get(prefix, 99999)
                rows_final.append({
                    'Orden': pos, 'Odoo_Cta': '', 'Odoo_Desc': f"--- [COI] {data['Descripcion']} ---", 
                    'Odoo_Saldo': None, 'COI_Cta': data['Cuenta_Orig'], 'COI_Desc': data['Descripcion'], 
                    'COI_Saldo': data['Saldo'], 'Diff': None, 'Status': "NO EN ELISA", 'Is_Header': False
                })

    df_fin = pd.DataFrame(rows_final).sort_values('Orden')
    writer = pd.ExcelWriter(FILE_OUTPUT, engine='xlsxwriter')
    wb, ws = writer.book, writer.book.add_worksheet('Conciliacion')
    
    def get_set(bg, bold=False):
        p_r = {'bold': bold, 'border': 1}; p_m = {'num_format': '$ #,##0.00', 'bold': bold, 'border': 1}
        if bg: p_r['bg_color'] = p_m['bg_color'] = bg
        return wb.add_format(p_r), wb.add_format(p_m)

    f_ok_r, f_ok_m = get_set('#C6EFCE', True); f_bad_r, f_bad_m = get_set('#FFC7CE', True)
    f_audit_r, f_audit_m = get_set("#F5CC27", True); f_std_r, f_std_m = get_set(None, False)
    f_hdr = wb.add_format({'bg_color': '#D9D9D9', 'bold': True, 'border': 1, 'align': 'center'})
    f_ab_error = wb.add_format({'bg_color': '#9C0006', 'font_color': '#FFFFFF', 'bold': True, 'align': 'center'})
    f_i_ok = wb.add_format({'bold': True, 'font_color': '#006100', 'align': 'center'})

    for c, h in enumerate(['Odoo Cta', 'Odoo Desc', 'Odoo Saldo', 'COI Cta', 'COI Desc', 'COI Saldo', 'Diff', 'Estatus', 'Check Abuelas']):
        ws.write(0, c, h, f_hdr)

    for i, r in enumerate(df_fin.to_dict('records')):
        row_idx = i + 1
        st, is_h, cta_coi = r['Status'], r['Is_Header'], str(r['COI_Cta'])
        is_ab = is_abuela_format(cta_coi)

        if "NO EN COI" in st or st == "NO EN ELISA": fr, fm = f_audit_r, f_audit_m
        elif is_h or is_ab: fr, fm = (f_ok_r, f_ok_m) if st == "OK" else (f_bad_r, f_bad_m)
        else: fr, fm = f_std_r, f_std_m
        
        ws.write(row_idx, 0, r['Odoo_Cta'], fr); ws.write(row_idx, 1, r['Odoo_Desc'], fr)
        safe_write_money(ws, row_idx, 2, r['Odoo_Saldo'], fm)
        ws.write(row_idx, 3, r['COI_Cta'], fr); ws.write(row_idx, 4, r['COI_Desc'], fr)
        safe_write_money(ws, row_idx, 5, r['COI_Saldo'], fm)
        safe_write_money(ws, row_idx, 6, r['Diff'], fm); ws.write(row_idx, 7, st, fr)

        if is_ab:
            n_ab = normalize_code(cta_coi)
            if "SUMA" in n_ab:
                ws.write(row_idx, 8, "OK (SUMA)", f_i_ok)
            else:
                prefix = n_ab[:4]
                sal_ab = (coi_lookup.get(n_ab, {}).get('Saldo', 0.0))
                sum_h = sum((v['Saldo'] or 0.0) for k, v in coi_lookup.items() if k.startswith(prefix) and k != n_ab and k.endswith('000') and not k.startswith("SUMA"))
                if abs(sum_h - sal_ab) < 0.1: ws.write(row_idx, 8, "OK", f_i_ok)
                else: ws.write(row_idx, 8, f"ERR: {sum_h-sal_ab:,.2f}", f_ab_error)

    ws.set_column('B:B', 50); ws.set_column('E:E', 40); ws.set_column('C:I', 15); writer.close()
    print("Versión 18.7 finalizada.")

if __name__ == "__main__":
    generar_analisis_v18_7()