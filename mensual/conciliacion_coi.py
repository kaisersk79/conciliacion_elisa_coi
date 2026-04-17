import pandas as pd
import numpy as np
import re
from collections import Counter

# --- CONFIGURACIÓN ---
FILE_ELISA = 'Reporte_Contable_Final.xlsx' 
FILE_COI = 'COI_Limpio_Base.xlsx'          
FILE_OUTPUT = 'Conciliacion_Elisa_vs_COI.xlsx'

# --- 1. SUMAS VIRTUALES ---
SUMAS_VIRTUALES = {
    'SUMA-BANCOS-TOTAL': ['1120-000-000', '1121-000-000', '1122-000-000'],
    'SUMA-BANCOS-USD': ['1121-001-000', '1122-000-000'],
    'SUMA-CLIENTES-NACIONALES': ['1150-001-000', '1150-004-000', '1150-005-000', '1150-006-000'],
    'SUMA-CLIENTES-EXTRANJEROS': ['1150-002-000', '1150-003-000'],
    'SUMA-PROVEEDORES-NACIONALES': ['2110-001-000', '2115-000-000']
}

# --- 2. MAPA MAESTRO ---
MAPA_ELISA_COI = {
    '101': '1110-000-000', '102': 'SUMA-BANCOS-TOTAL', '102.01': '1120-000-000',
    '102.02': 'SUMA-BANCOS-USD', '102.02.01': '1121-001-000', '102.02.02': '1122-000-000',
    '104': '1140-000-000', '105': '1150-000-000', '105.01': 'SUMA-CLIENTES-NACIONALES',
    '105.02': 'SUMA-CLIENTES-EXTRANJEROS', '107': '1170-000-000', '107.02': '1170-002-000', 
    '109': '1210-000-000', '113': '1180-000-000', '114': '1220-000-000', 
    '115': '1190-000-000', '118': '1200-000-000', '119': '1201-000-000', 
    '120': '1215-000-000', '120.02.01': '1215-002-000', '153': '1310-006-000', 
    '154': '1310-003-000', '155': '1310-005-000', '156': '1310-004-000', 
    '160': '1310-007-000', '171': '1360-000-000', '171.03': '1360-002-000', 
    '201': '2110-000-000', '201.01': '2110-001-000', '201.01.01': '2110-001-001', 
    '201.03.01': '2115-000-000', '205': '2120-000-000', '205.02': '2120-001-000', 
    '205.02.01': '2120-001-001', '205.02.06': '2120-001-013', '205.02.09': '2120-001-019', 
    '206': '2190-000-000', '206.01.01': '2190-001-000', '208': '2180-000-000', 
    '209': '2181-000-000', '210': '2160-000-000', '211': '2170-000-000', 
    '213': '2140-000-000', '216': '2150-000-000', '251': '2130-000-000', 
    '401': '4100-000-000', '401.04.01': '4100-002-000', '402': '4200-000-000', 
    '402.02.01': '4200-002-000', '501': '5000-000-000', '501.08': '5100-000-000', 
    '501.08.08': '5200-000-000', '602': '6100-000-000', '603': '6200-000-000', 
    '603.82': '6200-055-000', '604.59': '6200-034-000', '701': '7200-000-000', 
    '701.05': '7200-005-000', '702': '7100-000-000', '703': '7400-000-000', 
    '704': '7300-000-000'
}

# --- 3. EXCEPCIONES Y OMISIONES ---
EXCEPCIONES = {
    '107.05.99': '2120-001-013', # Samuel Villa
    '105.01.00': '1150-001-001'  # Corregimos Público en General
}

CUENTAS_IGNORADAS = [
    '105.02.00'  # Carpeta Extranjeros
]

def safe_num(val):
    try:
        if pd.isna(val): return 0.0
        return float(val)
    except: return 0.0

def cruzar_archivos():
    print("--- Cruzando Elisa vs COI (Limpieza Visual y Formato Condicional) ---")
    
    try:
        df_elisa = pd.read_excel(FILE_ELISA).fillna('')
        df_coi = pd.read_excel(FILE_COI).fillna('')
    except Exception as e:
        print(f"Error al leer archivos: {e}")
        return

    diccionario_coi = {}
    for _, row in df_coi.iterrows():
        cta = str(row['Cuenta']).strip()
        diccionario_coi[cta] = {
            'Descripcion': row['Descripcion'],
            'Saldo': safe_num(row['Saldo_Final'])
        }

    for nombre_virtual, lista_cuentas in SUMAS_VIRTUALES.items():
        total_virtual = sum([diccionario_coi.get(cta, {}).get('Saldo', 0.0) for cta in lista_cuentas])
        diccionario_coi[nombre_virtual] = {
            'Descripcion': f'[AGRUPACIÓN COI] {nombre_virtual.replace("SUMA-", "")}',
            'Saldo': total_virtual
        }

    datos_procesados = []
    lista_destinos = []
    patron_codigo_coi = re.compile(r'(\d{4}[-.]\d{3}[-.]\d{3})')

    for idx, row in df_elisa.iterrows():
        cta_elisa = str(row['Cuenta']).strip()
        desc_elisa = str(row['Descripcion']).strip()
        saldo_elisa = safe_num(row['Saldo'])
        
        if not cta_elisa and not desc_elisa: continue

        target_coi = ""
        es_padre = False
        es_detalle_encontrado = False

        if cta_elisa in CUENTAS_IGNORADAS:
            pass 
        elif cta_elisa in EXCEPCIONES:
            target_coi = EXCEPCIONES[cta_elisa]
            es_padre = True
        elif cta_elisa in MAPA_ELISA_COI:
            target_coi = MAPA_ELISA_COI[cta_elisa]
            es_padre = True
        
        if not target_coi and cta_elisa not in CUENTAS_IGNORADAS:
            match = patron_codigo_coi.search(desc_elisa)
            if match:
                target_coi = match.group(1).replace('.', '-')
                es_detalle_encontrado = True

        if target_coi:
            lista_destinos.append(target_coi)

        datos_procesados.append({
            'Elisa_Cuenta': cta_elisa,
            'Elisa_Desc': desc_elisa,
            'Elisa_Saldo': saldo_elisa,
            'Target_COI': target_coi,
            'Es_Padre': es_padre,
            'Es_Detalle': es_detalle_encontrado
        })

    frecuencia_destinos = Counter(lista_destinos)
    destinos_multiples = {k for k, v in frecuencia_destinos.items() if v > 1 and k != ""}
    
    saldos_agrupados_elisa = {}
    for d in destinos_multiples:
        saldos_agrupados_elisa[d] = sum(
            item['Elisa_Saldo'] for item in datos_procesados if item['Target_COI'] == d
        )

    filas_resultado = []
    headers_inyectados = set()

    for item in datos_procesados:
        t_coi = item['Target_COI']
        
        if t_coi in destinos_multiples:
            if t_coi not in headers_inyectados:
                desc_coi_temp = diccionario_coi[t_coi]['Descripcion'] if t_coi in diccionario_coi else "CÓDIGO NO ENCONTRADO"
                saldo_coi_temp = diccionario_coi[t_coi]['Saldo'] if t_coi in diccionario_coi else 0.0
                
                fila_header = {
                    'Elisa_Cuenta': '[AGRUPADO]',
                    'Elisa_Desc': f'[SUMA ELISA] {desc_coi_temp}',
                    'Elisa_Saldo': saldos_agrupados_elisa[t_coi],
                    'COI_Cuenta': t_coi,
                    'COI_Desc': desc_coi_temp,
                    'COI_Saldo': saldo_coi_temp,
                    'Es_Agrupacion_Elisa': True, 
                    'Es_Padre': False,
                    'Es_Detalle': False,
                    'Hacer_Resta': True
                }
                filas_resultado.append(fila_header)
                headers_inyectados.add(t_coi)
            
            desc_coi_temp = diccionario_coi[t_coi]['Descripcion'] if t_coi in diccionario_coi else ""
            fila_detalle = {
                'Elisa_Cuenta': item['Elisa_Cuenta'],
                'Elisa_Desc': item['Elisa_Desc'],
                'Elisa_Saldo': item['Elisa_Saldo'],
                'COI_Cuenta': t_coi,
                'COI_Desc': desc_coi_temp,
                'COI_Saldo': None,
                'Es_Agrupacion_Elisa': False,
                'Es_Padre': False,
                'Es_Detalle': True,
                'Hacer_Resta': False
            }
            filas_resultado.append(fila_detalle)
            
        else:
            hacer_resta = False
            if item['Es_Padre']:
                hacer_resta = True
            elif item['Es_Detalle'] and frecuencia_destinos.get(t_coi, 0) == 1:
                hacer_resta = True

            desc_coi_temp = diccionario_coi[t_coi]['Descripcion'] if t_coi in diccionario_coi else ("" if not t_coi else "CÓDIGO NO ENCONTRADO")
            saldo_coi_temp = diccionario_coi[t_coi]['Saldo'] if t_coi in diccionario_coi else 0.0
            
            fila = {
                'Elisa_Cuenta': item['Elisa_Cuenta'],
                'Elisa_Desc': item['Elisa_Desc'],
                'Elisa_Saldo': item['Elisa_Saldo'],
                'COI_Cuenta': t_coi if t_coi else '',
                'COI_Desc': desc_coi_temp,
                'COI_Saldo': saldo_coi_temp if hacer_resta else None,
                'Es_Agrupacion_Elisa': False,
                'Es_Padre': item['Es_Padre'],
                'Es_Detalle': item['Es_Detalle'],
                'Hacer_Resta': hacer_resta
            }
            filas_resultado.append(fila)

    # --- EXPORTACIÓN ---
    df_res = pd.DataFrame(filas_resultado)
    writer = pd.ExcelWriter(FILE_OUTPUT, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}})
    wb = writer.book
    ws = wb.add_worksheet('Conciliacion')

    # --- ESTILOS VISUALES (Negativos neutros) ---
    f_hdr = wb.add_format({'bg_color': '#203764', 'font_color': 'white', 'bold': True, 'border': 1})
    
    f_padre = wb.add_format({'bg_color': '#FFF2CC', 'bold': True, 'border': 1, 'font_size': 10})
    f_padre_num = wb.add_format({'bg_color': '#FFF2CC', 'bold': True, 'border': 1, 'font_size': 10, 'num_format': '$ #,##0.00'})
    
    f_det = wb.add_format({'border': 1, 'font_size': 10})
    f_det_num = wb.add_format({'border': 1, 'font_size': 10, 'num_format': '$ #,##0.00'})
    
    f_gray = wb.add_format({'font_color': '#7F7F7F', 'italic': True, 'font_size': 10})
    f_gray_num = wb.add_format({'font_color': '#7F7F7F', 'italic': True, 'font_size': 10, 'num_format': '$ #,##0.00'})
    
    f_diff = wb.add_format({'num_format': '$ #,##0.00', 'border': 1, 'font_size': 10, 'bg_color': '#FCE4D6'}) 

    # --- ESTILOS DE FORMATO CONDICIONAL ---
    fmt_ok_cond = wb.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'bold': True})
    fmt_bad_cond = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'bold': True})

    headers = ['Elisa Cta', 'Elisa Desc', 'Elisa Saldo', 'COI Cta', 'COI Desc', 'COI Saldo', 'Diferencia', 'Estatus']
    for col_num, header_title in enumerate(headers): ws.write(0, col_num, header_title, f_hdr)

    for i, row in enumerate(df_res.to_dict('records')):
        xl_row = i + 1 
        
        if row['Es_Padre'] or row.get('Es_Agrupacion_Elisa'):
            fmt_texto, fmt_numero = f_padre, f_padre_num
        elif row['Es_Detalle']:
            fmt_texto, fmt_numero = f_det, f_det_num
        else:
            fmt_texto, fmt_numero = f_gray, f_gray_num

        ws.write(xl_row, 0, row['Elisa_Cuenta'], fmt_texto)
        ws.write(xl_row, 1, row['Elisa_Desc'], fmt_texto)
        ws.write_number(xl_row, 2, row['Elisa_Saldo'], fmt_numero)
        
        ws.write(xl_row, 3, row['COI_Cuenta'], fmt_texto)
        ws.write(xl_row, 4, row['COI_Desc'], fmt_texto)

        if row['Hacer_Resta']:
            ws.write_number(xl_row, 5, row['COI_Saldo'] if row['COI_Saldo'] is not None else 0, fmt_numero)
            ws.write_formula(xl_row, 6, f"=ABS(C{xl_row+1})-ABS(F{xl_row+1})", f_diff)
            ws.write_formula(xl_row, 7, f'=IF(ABS(G{xl_row+1})<0.1,"OK","REVISAR")', fmt_texto)
        else:
            ws.write(xl_row, 5, "", fmt_texto)
            ws.write(xl_row, 6, "", fmt_texto)
            ws.write(xl_row, 7, "", fmt_texto)

    # --- APLICAR FORMATO CONDICIONAL A LA COLUMNA ESTATUS ---
    # Aplica a toda la columna H (índice 7), desde la fila 1 hasta el final de los datos
    total_filas = len(df_res)
    if total_filas > 0:
        ws.conditional_format(1, 7, total_filas, 7, {
            'type': 'text', 'criteria': 'containing', 'value': 'OK', 'format': fmt_ok_cond
        })
        ws.conditional_format(1, 7, total_filas, 7, {
            'type': 'text', 'criteria': 'containing', 'value': 'REVISAR', 'format': fmt_bad_cond
        })

    ws.set_column('B:B', 45); ws.set_column('E:E', 45); ws.set_column('A:A', 15)
    ws.set_column('C:D', 15); ws.set_column('F:H', 15)
    ws.freeze_panes(1, 0)
    writer.close()
    print(f"¡Cruce completado con éxito! Archivo generado: {FILE_OUTPUT}")

if __name__ == "__main__":
    cruzar_archivos()