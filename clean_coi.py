import pandas as pd
import numpy as np
import re

# --- CONFIGURACIÓN ---
FILE_PATH = 'aux_coi_diciembre.xlsx' 

def obtener_nombre_rubro(cuenta_str):
    c = str(cuenta_str).strip()
    
    # REGLAS DE NOMBRES DE GRUPOS
    if c.startswith("111"): return "Caja"
    if c.startswith("112"): return "Bancos"          
    if c.startswith("114"): return "Otros Activos"
    
    # --- AGRUPACIÓN DE CLIENTES ---
    # 1. Extranjeros aparte (para que no se mezclen en la suma nacional)
    if c.startswith("1150-002") or c.startswith("1150-003"): 
        return "Clientes Extranjeros"
    
    # 2. Clientes Nacionales (Incluye 1150-001 y 1150-004)
    # Al ponerles el mismo nombre, el Excel los agrupará y sumará juntos en el encabezado amarillo.
    if c.startswith("115"): 
        return "Clientes Nacionales"
    
    if c.startswith("117"): return "Deudores Diversos"
    if c.startswith("118"): return "Impuestos Acreditables"
    if c.startswith("119"): return "Impuestos por Acreditar"
    if c.startswith("120"): return "Anticipo a Proveedores"
    
    if c.startswith("201") or c.startswith("210"): return "Proveedores"
    if c.startswith("211"): return "Impuestos por Pagar"
    if c.startswith("212"): return "Acreedores Diversos"
    if c.startswith("219"): return "Anticipo de Clientes"
    
    if c.startswith("51"):  return "Ventas"
    if c.startswith("71"):  return "Gastos de Operación"
    if c.startswith("72"):  return "Gastos Financieros"
    
    return f"Rubro {c[:4]}"

def limpiar_saldo(valor):
    if pd.isna(valor): return None
    if isinstance(valor, (int, float)): return valor
    clean_val = str(valor).replace(',', '').replace(' ', '')
    try: return float(clean_val)
    except: return None

def limpiar_descripcion(texto):
    if pd.isna(texto): return ""
    txt = str(texto).strip()
    if txt.lower() == 'nan': return ""
    return txt

def procesar_coi_final():
    print(f"--- Procesando COI: Suma Nacionales (001 + 004) ---")
    
    try:
        df = pd.read_excel(FILE_PATH, header=None, engine='openpyxl')
    except Exception as e:
        print(f"Error crítico: {e}")
        return

    # 1. ENCONTRAR COLUMNA SALDO
    saldo_col_idx = -1
    for i in range(min(20, len(df))):
        row_vals = [str(x) for x in df.iloc[i].tolist()]
        candidates = [idx for idx, val in enumerate(row_vals) if "Saldo" in val and "inicial" not in val]
        if candidates:
            saldo_col_idx = candidates[-1]
            break
    if saldo_col_idx == -1: saldo_col_idx = df.shape[1] - 1

    # 2. EXTRAER CUENTAS
    raw_cuentas = []
    cuenta_actual = None
    desc_actual = None
    saldo_actual = 0.0
    en_bloque = False
    patron = re.compile(r"Cuenta\s*:\s*([\d-]+)\s+(.*)")

    for index, row in df.iterrows():
        fila_txt = " ".join([str(x) for x in row.iloc[:3] if pd.notna(x)])
        match = patron.search(fila_txt)
        
        if match:
            # Guardar anterior
            if cuenta_actual:
                raw_cuentas.append({
                    'Cuenta': cuenta_actual,
                    'Descripcion': desc_actual.strip(),
                    'Saldo': saldo_actual
                })
            
            # Nueva
            cuenta_actual = match.group(1).strip()
            desc_actual = match.group(2).strip()
            saldo_actual = 0.0
            en_bloque = True
            
            # Saldo en línea de título (Madres)
            if saldo_col_idx < len(row):
                val = row.iloc[saldo_col_idx]
                num = limpiar_saldo(val)
                if num is not None: saldo_actual = num
            continue
        
        # Saldo en movimientos (Hijas)
        if en_bloque and saldo_col_idx < len(row):
            val = row.iloc[saldo_col_idx]
            num = limpiar_saldo(val)
            if num is not None:
                if isinstance(val, str) and ("Saldo" in val or "Haber" in val): continue
                saldo_actual = num

    if cuenta_actual:
        raw_cuentas.append({
            'Cuenta': cuenta_actual,
            'Descripcion': desc_actual.strip(),
            'Saldo': saldo_actual
        })

    df_clean = pd.DataFrame(raw_cuentas)
    if df_clean.empty: 
        print("Error: No se encontraron cuentas.")
        return

    # --- 3. JERARQUÍA ESTRICTA (Detectar Hojas vs Padres) ---
    def get_clean_base(cta):
        base = cta
        while base.endswith('-000') or base.endswith('000'):
            if base.endswith('-000'): base = base[:-4]
            elif base.endswith('000'): base = base[:-3]
        return base

    df_clean['Codigo_Base'] = df_clean['Cuenta'].apply(get_clean_base)
    todas_cuentas_full = set(df_clean['Cuenta'].unique())

    def determinar_rol(row):
        mi_base = row['Codigo_Base']
        mi_cuenta = row['Cuenta']
        prefix = mi_base + "-"
        for c in todas_cuentas_full:
            if c != mi_cuenta and c.startswith(prefix):
                return True # Es Padre
        return False # Es Hoja (Movimiento final)

    df_clean['Es_Padre'] = df_clean.apply(determinar_rol, axis=1)
    df_clean['Es_Madre_Suprema'] = df_clean['Cuenta'].str.endswith('000-000')

    # --- 4. CHECK DE VALIDACIÓN ---
    hojas_df = df_clean[~df_clean['Es_Padre']].copy()
    
    def calcular_check(row):
        if not row['Es_Padre']: return ""
        
        mi_base = row['Codigo_Base']
        prefix = mi_base + "-"
        
        # Check: Compara el saldo de esta cuenta PADRE contra la suma de sus HIJAS
        mask_desc = hojas_df['Cuenta'].str.startswith(prefix)
        suma_hijos = hojas_df[mask_desc]['Saldo'].sum()
        
        diff = row['Saldo'] - suma_hijos
        
        if abs(diff) < 0.1: return "OK (0.00)"
        return f"DIF: {diff:,.2f}"

    df_clean['Check'] = df_clean.apply(calcular_check, axis=1)

    # --- 5. EXPORTAR EXCEL ---
    df_clean['Grupo'] = df_clean['Cuenta'].apply(obtener_nombre_rubro)
    
    # Ordenamos por cuenta para mantener el orden lógico (001 antes que 004)
    df_clean.sort_values('Cuenta', inplace=True)
    
    filas_excel = []
    
    # El orden de los grupos será según aparezcan en la lista ordenada
    # Como 1150-001 (Nac) aparece antes que 1150-002 (Ext), 
    # el grupo "Clientes Nacionales" se creará primero y absorberá también a 1150-004 cuando llegue.
    
    grupos_procesados = []
    # Iteramos sobre el dataframe ordenado
    # Usamos un set para no repetir grupos ya procesados
    
    # Obtenemos lista única preservando orden
    grupos_orden = df_clean['Grupo'].unique()

    for grp in grupos_orden:
        datos = df_clean[df_clean['Grupo'] == grp]
        
        # TOTAL AMARILLO:
        # Sumamos SOLO las hojas (nietos/hijos finales).
        # Esto sumará las hojas de 1150-001 y las hojas de 1150-004.
        # Resultado esperado: $833k
        total_grp = datos[~datos['Es_Padre']]['Saldo'].sum()
        
        filas_excel.append({
            'Cuenta': '', 'Descripcion': grp, 'Saldo': total_grp, 
            'Check': '', 'Nivel': 1, 'Es_Padre': False
        })
        
        for _, row in datos.iterrows():
            filas_excel.append({
                'Cuenta': row['Cuenta'],
                'Descripcion': row['Descripcion'],
                'Saldo': row['Saldo'],
                'Check': row['Check'],
                'Nivel': 2 if row['Es_Padre'] else 3,
                'Es_Padre': row['Es_Padre']
            })
            
        filas_excel.append({'Cuenta': '', 'Descripcion': '', 'Saldo': np.nan, 'Check': '', 'Nivel': '', 'Es_Padre': False})

    df_export = pd.DataFrame(filas_excel)

    nombre_archivo = 'COI_Final_SumaCorrecta.xlsx'
    writer = pd.ExcelWriter(nombre_archivo, engine='xlsxwriter')
    
    df_export[['Cuenta', 'Descripcion', 'Saldo', 'Check']].to_excel(writer, index=False, sheet_name='Reporte')
    
    wb = writer.book
    ws = writer.sheets['Reporte']
    
    # Formatos
    fmt_curr = wb.add_format({'num_format': '$ #,##0.00;[Red]-$ #,##0.00'})
    fmt_group = wb.add_format({'bold': True, 'bg_color': '#FFFF00', 'border': 1, 'num_format': '$ #,##0.00'})
    fmt_padre = wb.add_format({'bold': True, 'bg_color': '#DDEBF7', 'num_format': '$ #,##0.00'})
    fmt_ok = wb.add_format({'font_color': 'green', 'bold': True, 'align': 'center'})
    fmt_bad = wb.add_format({'bg_color': 'red', 'font_color': 'white', 'bold': True, 'align': 'center'})

    for i, row in enumerate(df_export.to_dict('records')):
        ridx = i + 1
        
        if pd.notna(row['Saldo']): ws.write(ridx, 2, row['Saldo'], fmt_curr)
        
        val = str(row['Check'])
        if "OK" in val: ws.write(ridx, 3, val, fmt_ok)
        elif "DIF" in val: ws.write(ridx, 3, val, fmt_bad)
        
        if row['Nivel'] == 1:
            ws.set_row(ridx, None, fmt_group)
            ws.write(ridx, 1, row['Descripcion'], fmt_group)
            if pd.notna(row['Saldo']): ws.write(ridx, 2, row['Saldo'], fmt_group)
            
        elif row['Es_Padre']:
            ws.write(ridx, 0, row['Cuenta'], fmt_padre)
            ws.write(ridx, 1, row['Descripcion'], fmt_padre)
            if pd.notna(row['Saldo']): ws.write(ridx, 2, row['Saldo'], fmt_padre)

    ws.set_column('A:A', 20)
    ws.set_column('B:B', 60)
    ws.set_column('C:C', 18)
    ws.set_column('D:D', 15)
    
    writer.close()
    print(f"¡Listo! Archivo con suma unificada generado: {nombre_archivo}")

if __name__ == "__main__":
    procesar_coi_final()