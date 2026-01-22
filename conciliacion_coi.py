import pandas as pd
import numpy as np
import re
import math

# ==========================================
# --- CONFIGURACIÓN DE ARCHIVOS ---
# ==========================================
# Asegúrate de que estos nombres sean idénticos a tus archivos
FILE_ODOO = 'Reporte_Contable_Final.xlsx'
FILE_COI = 'COI_Final_SumaCorrecta.xlsx'
FILE_OUTPUT = 'Analisis_Comparativo_Noviembre.xlsx'

# ==========================================
# --- FUNCIONES DE LIMPIEZA Y UTILIDAD ---
# ==========================================

def extract_key(desc):
    """
    Extrae la cuenta COI del paréntesis en la descripción de Odoo.
    Ejemplo: de '(1110-001-000 Caja) Efectivo' extrae '1110-001-000'
    """
    if pd.isna(desc): return None
    # Busca texto entre paréntesis que tenga dígitos
    match = re.search(r'\(([\d-]+)', str(desc))
    if match:
        return match.group(1).strip()
    return None

def clean_money(val):
    """Convierte texto con formato de moneda ($ 1,234.56) a número flotante."""
    if pd.isna(val): return 0.0
    # Quitamos símbolos de moneda, comas y espacios
    s = str(val).replace('$','').replace(',','').replace(' ','')
    try: return float(s)
    except: return 0.0

def safe_write(worksheet, row, col, val, fmt=None):
    """
    Escribe en Excel de forma segura. 
    Si el valor es NaN (vacío) o Infinito, escribe una cadena vacía "" 
    para evitar que el script falle (crash).
    """
    try:
        # Verificamos si es número inválido (NaN o Inf)
        if pd.isna(val) or (isinstance(val, (int, float)) and (math.isnan(val) or math.isinf(val))):
            worksheet.write(row, col, "", fmt)
        else:
            worksheet.write(row, col, val, fmt)
    except:
        # Si falla por cualquier otra razón (ej. tipo de dato raro), escribe vacío
        worksheet.write(row, col, "", fmt)

# ==========================================
# --- LÓGICA PRINCIPAL ---
# ==========================================

def generar_analisis():
    print(f"--- Iniciando Generación de {FILE_OUTPUT} ---")
    
    # 1. CARGAR DATOS
    try:
        # Intentamos leer Excel con el motor 'openpyxl'
        df_odoo = pd.read_excel(FILE_ODOO, engine='openpyxl')
        df_coi = pd.read_excel(FILE_COI, engine='openpyxl')
    except Exception as e:
        print(f"Error cargando archivos: {e}")
        # Intento de fallback si el usuario tiene CSVs renombrados
        try:
            df_odoo = pd.read_csv(FILE_ODOO.replace('.xlsx','.csv'))
            df_coi = pd.read_csv(FILE_COI.replace('.xlsx','.csv'))
        except:
            print("No se pudieron cargar los archivos.")
            return

    # 2. PROCESAR ODOO (Lado Izquierdo)
    # Contamos puntos para determinar el nivel jerárquico (para colores)
    df_odoo['Dots'] = df_odoo['Cuenta'].astype(str).apply(lambda x: x.count('.'))
    
    # Extraemos la llave de cruce y limpiamos saldos
    df_odoo['Key_COI'] = df_odoo['Descripcion'].apply(extract_key)
    df_odoo['Saldo'] = df_odoo['Saldo'].apply(clean_money)
    
    # --- CÁLCULO DE SUMAS AGRUPADAS ---
    # Esto es vital: Si Odoo tiene 3 renglones que apuntan a la misma cuenta COI 
    # (ej. Efectivo, Caja Chica, Sucursal), sumamos sus montos para compararlos 
    # contra el saldo único que trae el auxiliar de COI.
    odoo_details = df_odoo[df_odoo['Key_COI'].notna()].copy()
    odoo_sums = odoo_details.groupby('Key_COI')['Saldo'].sum().to_dict()

    # 3. PROCESAR COI (Lado Derecho)
    df_coi['Saldo'] = df_coi['Saldo'].apply(clean_money)
    
    # Creamos diccionarios para búsqueda rápida por cuenta
    coi_lookup = df_coi.groupby('Cuenta')['Saldo'].sum().to_dict()
    coi_desc_lookup = df_coi.set_index('Cuenta')['Descripcion'].to_dict()

    # 4. CONSTRUIR EL REPORTE (Fila por Fila)
    results = []
    
    for idx, row in df_odoo.iterrows():
        key = row['Key_COI']
        
        # Estructura de la fila final
        item = {
            'Odoo_Cuenta': row['Cuenta'],
            'Odoo_Descripcion': row['Descripcion'],
            'Elisa_Nov_25': row['Saldo'],
            'Dots': row['Dots'], # Nivel jerárquico para formato
            
            # Datos de COI (Vacíos por defecto)
            'COI_Cuenta': None,
            'COI_Descripcion': None,
            'COI_Nov_25': None,
            
            # Análisis
            'Diferencia': None,
            'Estatus': None
        }
        
        if key:
            # Si la fila tiene llave de cruce
            if key in coi_lookup:
                # ¡Match encontrado!
                item['COI_Cuenta'] = key
                item['COI_Descripcion'] = coi_desc_lookup.get(key, "")
                item['COI_Nov_25'] = coi_lookup[key]
                
                # Comparación: Suma Total Odoo vs Saldo COI
                sum_odoo = odoo_sums.get(key, 0.0)
                coi_bal = coi_lookup[key]
                diff = sum_odoo - coi_bal
                
                item['Diferencia'] = diff
                
                # Semáforo con tolerancia de 10 centavos
                if abs(diff) < 0.1:
                    item['Estatus'] = "OK"
                else:
                    item['Estatus'] = "DIFERENCIA"
            else:
                # Tiene llave pero no existe en COI
                item['COI_Cuenta'] = key
                item['Estatus'] = "NO EN COI"
                item['Diferencia'] = item['Elisa_Nov_25'] # Asumimos todo como diferencia
        
        results.append(item)

    df_final = pd.DataFrame(results)

    # 5. EXPORTAR A EXCEL CON FORMATO
    writer = pd.ExcelWriter(FILE_OUTPUT, engine='xlsxwriter')
    
    # Definir el orden exacto de columnas
    cols_order = [
        'Odoo_Cuenta', 'Odoo_Descripcion', 'Elisa_Nov_25', 
        'COI_Cuenta', 'COI_Descripcion', 'COI_Nov_25', 
        'Diferencia', 'Estatus'
    ]
    
    # Escribir solo los encabezados primero
    df_final.to_excel(writer, index=False, columns=[c for c in cols_order if c in df_final.columns], sheet_name='Analisis 2025')
    
    workbook = writer.book
    worksheet = writer.sheets['Analisis 2025']
    
    # --- Estilos ---
    fmt_currency = workbook.add_format({'num_format': '$ #,##0.00;[Red]-$ #,##0.00'})
    
    # Semáforos
    fmt_ok = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'bold': True, 'align': 'center'})
    fmt_dif = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'bold': True, 'align': 'center'})
    fmt_miss = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700', 'bold': True, 'align': 'center'})
    
    # Jerarquía (Amarillo para Nivel 1, Gris para Nivel 2)
    fmt_lvl1 = workbook.add_format({'bg_color': '#FFFF00', 'bold': True})
    fmt_lvl2 = workbook.add_format({'bg_color': '#D9D9D9', 'bold': True})
    
    # Escribir datos celda por celda aplicando formatos
    for i, row in df_final.iterrows():
        r = i + 1 # +1 por el encabezado
        dots = row['Dots']
        
        # Determinar estilo de fila Odoo
        fmt_row = None
        if dots == 0: fmt_row = fmt_lvl1
        elif dots == 1: fmt_row = fmt_lvl2
        
        # 1. Datos Odoo (Izquierda)
        safe_write(worksheet, r, 0, row['Odoo_Cuenta'], fmt_row)
        safe_write(worksheet, r, 1, row['Odoo_Descripcion'], fmt_row)
        safe_write(worksheet, r, 2, row['Elisa_Nov_25'], fmt_currency)
        
        # 2. Datos COI (Derecha)
        safe_write(worksheet, r, 3, row['COI_Cuenta'])
        safe_write(worksheet, r, 4, row['COI_Descripcion'])
        safe_write(worksheet, r, 5, row['COI_Nov_25'], fmt_currency)
        
        # 3. Diferencia y Estatus (Extremo Derecho)
        safe_write(worksheet, r, 6, row['Diferencia'], fmt_currency)
        
        st = row['Estatus']
        fmt_st = None
        if st == "OK": fmt_st = fmt_ok
        elif st == "DIFERENCIA": fmt_st = fmt_dif
        elif st == "NO EN COI": fmt_st = fmt_miss
        
        safe_write(worksheet, r, 7, st, fmt_st)

    # Ajustar ancho de columnas para legibilidad
    worksheet.set_column('A:A', 15) # Cta Odoo
    worksheet.set_column('B:B', 50) # Desc Odoo
    worksheet.set_column('C:C', 18) # Saldo Odoo
    worksheet.set_column('D:D', 18) # Cta COI
    worksheet.set_column('E:E', 40) # Desc COI
    worksheet.set_column('F:H', 18) # Saldos y Status

    writer.close()
    print(f"¡Listo! Reporte generado exitosamente: {FILE_OUTPUT}")

if __name__ == "__main__":
    generar_analisis()