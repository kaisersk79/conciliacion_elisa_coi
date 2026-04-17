import pandas as pd
import re

# --- CONFIGURACIÓN ---
FILE_PATH = 'aux_coi.xlsx' # Asegúrate de que este sea tu archivo original
FILE_OUTPUT = 'COI_Limpio_Base.xlsx'

def limpiar_saldo(valor):
    if pd.isna(valor): return None
    if isinstance(valor, (int, float)): return float(valor)
    clean_val = str(valor).replace('$', '').replace(',', '').replace(' ', '').strip()
    try: return float(clean_val)
    except: return None

def determinar_nivel(cuenta):
    """Clasifica la cuenta según su terminación en COI."""
    if cuenta.endswith('-000-000'):
        return '1. MAYOR'
    elif cuenta.endswith('-000'):
        return '2. MENOR'
    else:
        return '3. DETALLE'

def limpiar_coi():
    print("--- Extrayendo Base Limpia de COI ---")
    
    try:
        # Usamos el motor genérico por si es CSV o Excel
        if FILE_PATH.endswith('.csv'):
            df = pd.read_csv(FILE_PATH, header=None, encoding='latin1')
        else:
            df = pd.read_excel(FILE_PATH, header=None)
    except Exception as e:
        print(f"Error al leer el archivo original: {e}")
        return

    # 1. Encontrar la columna que tiene los saldos finales
    saldo_col_idx = -1
    for i in range(min(30, len(df))):
        row_vals = [str(x).lower() for x in df.iloc[i].tolist()]
        # Buscamos la columna de "saldo actual" o equivalente
        candidatos = [idx for idx, val in enumerate(row_vals) if "saldo" in val and "inicial" not in val]
        if candidatos:
            saldo_col_idx = candidatos[-1]
            break
            
    if saldo_col_idx == -1: 
        saldo_col_idx = df.shape[1] - 1 # Asumimos la última si no hay cabecera clara

    # 2. Extracción plana
    raw_cuentas = []
    cuenta_actual, desc_actual, saldo_actual = None, None, 0.0
    patron_cuenta = re.compile(r"Cuenta\s*:\s*([\d-]+)\s+(.*)")

    for index, row in df.iterrows():
        fila_txt = " ".join([str(x) for x in row.iloc[:3] if pd.notna(x)])
        match = patron_cuenta.search(fila_txt)
        
        if match:
            # Guardamos la cuenta anterior antes de empezar la nueva
            if cuenta_actual:
                raw_cuentas.append({
                    'Cuenta': cuenta_actual,
                    'Descripcion': desc_actual,
                    'Saldo_Final': saldo_actual,
                    'Nivel': determinar_nivel(cuenta_actual)
                })
            
            # Inicializamos la nueva cuenta
            cuenta_actual = match.group(1).strip()
            desc_actual = match.group(2).strip()
            saldo_actual = 0.0
            
            # A veces el saldo de las cuentas Mayor viene en la misma fila del título
            if saldo_col_idx < len(row):
                num = limpiar_saldo(row.iloc[saldo_col_idx])
                if num is not None: saldo_actual = num
            continue
        
        # Si estamos dentro de una cuenta, buscamos su saldo final en los movimientos
        if cuenta_actual and saldo_col_idx < len(row):
            val = row.iloc[saldo_col_idx]
            num = limpiar_saldo(val)
            if num is not None:
                # Evitar capturar cabeceras de texto que se hayan filtrado
                if isinstance(val, str) and ("saldo" in val.lower() or "haber" in val.lower()): 
                    continue
                saldo_actual = num

    # Guardar la última cuenta procesada
    if cuenta_actual:
        raw_cuentas.append({
            'Cuenta': cuenta_actual, 'Descripcion': desc_actual, 
            'Saldo_Final': saldo_actual, 'Nivel': determinar_nivel(cuenta_actual)
        })

    df_clean = pd.DataFrame(raw_cuentas)
    
    # Filtramos cuentas con saldo 0 para limpiar basura (Opcional, si quieres ver todo coméntalo)
    # df_clean = df_clean[df_clean['Saldo_Final'] != 0.0]

    # 3. Exportar base de datos plana
    writer = pd.ExcelWriter(FILE_OUTPUT, engine='xlsxwriter')
    df_clean.to_excel(writer, index=False, sheet_name='Base_COI')
    
    wb = writer.book
    ws = writer.sheets['Base_COI']
    
    # Formato visual simple
    f_hdr = wb.add_format({'bg_color': '#4F81BD', 'font_color': 'white', 'bold': True})
    f_num = wb.add_format({'num_format': '$ #,##0.00'})
    
    for col_num, value in enumerate(df_clean.columns.values):
        ws.write(0, col_num, value, f_hdr)
        
    ws.set_column('A:A', 18)
    ws.set_column('B:B', 50)
    ws.set_column('C:C', 18, f_num)
    ws.set_column('D:D', 15)
    
    writer.close()
    print(f"¡Listo! Base extraída sin alteraciones en: {FILE_OUTPUT}")

if __name__ == "__main__":
    limpiar_coi()