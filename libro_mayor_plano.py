import pandas as pd
import numpy as np

# --- CONFIGURACIÓN ---
FILE_PATH = 'libro_mayor_diciembre.xlsx'
HEADER_ROW = 2 

MAJOR_NAME_MAP = {
    # --- ACTIVOS ---
    "101": "Caja",
    "102": "Banco",
    "102.01": "Bancos nacionales",
    "102.02": "Bancos extranjeros",
    "104": "Otros instrumentos financieros",
    "105": "Clientes",
    "105.02": "Clientes extranjeros",
    "107": "Deudores diversos",
    "107.02": "Socios y accionistas",
    "107.05": "Otros deudores diversos",
    "109": "Pagos anticipados",
    "113": "Impuestos a favor",
    "114": "Pagos provisionales",
    "115": "Inventario",
    "118": "Impuestos acreditables pagados",
    "119": "Impuestos acreditables por pagar",
    "120": "Anticipo a proveedores",
    
    # --- ACTIVOS FIJOS ---
    "153": "Maquinaria y equipo",
    "154": "Automóviles, autobuses y camiones",
    "155": "Mobiliario y equipo de oficina",
    "156": "Equipo de cómputo",
    "160": "Otros activos fijos",
    "171": "Depreciación acumulada",
    
    # --- PASIVOS ---
    "201": "Proveedores",
    "205": "Acreedores diversos",
    "206": "Anticipo de clientes",
    "210": "Provisión de sueldos",
    "211": "Impuestos por pagar (IMSS/Infonavit)",
    "213": "IVA por pagar",
    "216": "Retenciones de impuestos",
    
    # --- CAPITAL ---
    "301": "Capital social",
    
    # --- INGRESOS Y COSTOS ---
    "401": "Ingresos",
    "501": "Costos",
    
    # --- RESULTADOS (GASTOS Y PRODUCTOS) ---
    "601": "Gastos Generales",
    "602": "Costo de venta",
    "602.72": "Fletes y acarreos",
    "602.34": "Honorarios a personas físicas residentes nacionales",
    "602.61": "Propaganda y publicidad",
    
    "603": "Gastos de administración",
    "603.01": "Sueldos y salarios",
    "603.31": "Asimilados a salarios",
    "603.03": "Tiempos extras",
    "603.22": "Estímulo al personal",
    "603.12": "Aguinaldo",
    "603.06": "Vacaciones",
    "603.07": "Prima vacacional",
    "603.15": "Despensa",
    "603.26": "Cuotas al IMSS",
    "603.28": "Aportaciones al SAR",
    "603.27": "Aportaciones al infonavit",
    "603.29": "Impuesto estatal sobre nóminas",
    "603.54": "Limpieza",
    "603.56": "Mantenimiento y conservación",
    "603.57": "Seguros y fianzas",
    "603.55": "Papelería y artículos de oficina",
    "603.34": "Honorarios a personas físicas residentes nacionales",
    "603.58": "Otros impuestos y derechos",
    "603.50": "Teléfono, internet",
    "603.48": "Combustibles y lubricantes",
    "603.16": "Transporte",
    "603.25": "Otras prestaciones al personal",
    "603.49": "Viáticos y gastos de viaje",
    "603.81": "Gastos no deducibles (sin requisitos fiscales)",
    "603.82": "Otros gastos de administración",
    
    "604": "Gastos de fabricación",
    "604.59": "Recargos fiscales",
    
    "701": "Gastos financieros",
    "702": "Productos financieros",
    "702.01": "Utilidad cambiaria",
    "703": "Otros gastos",
    "704": "Otros productos"
}

def procesar_contabilidad():
    print(f"--- Procesando {FILE_PATH} (Saldo tomado directamente del renglón de la cuenta) ---")
    
    try:
        df = pd.read_excel(FILE_PATH, header=HEADER_ROW, engine='openpyxl')
    except Exception as e:
        print(f"Error: {e}")
        return

    # 1. Preparar Datos Base
    # Filtramos solo las filas que tienen código de cuenta
    cuentas_df = df[df['Código'].notna()].copy()
    
    # --- CAMBIO AQUÍ ---
    # Tomamos el saldo directamente de la columna 'Balance' de esa misma fila
    cuentas_df['Saldo_Final'] = pd.to_numeric(cuentas_df['Balance'], errors='coerce').fillna(0)
    
    df_final = cuentas_df[['Código', 'Nombre de la cuenta', 'Saldo_Final']].rename(
        columns={'Código': 'Cuenta', 'Nombre de la cuenta': 'Descripcion_Cuenta'}
    )
    
    # (Se eliminó la lógica de búsqueda de 'Balance inicial' en filas inferiores)

    # --- REGLA DE RECLASIFICACIÓN: SAMUEL VILLA (205 -> 107.05) ---
    mask_samuel = (
        df_final['Cuenta'].astype(str).str.startswith('205') & 
        df_final['Descripcion_Cuenta'].str.contains('Samuel|Villa Rodríguez', case=False, na=False)
    )
    
    # Agrupación y Reclasificación
    df_final['Grupo_N1'] = df_final['Cuenta'].astype(str).str[:3]
    df_final['Grupo_N2'] = df_final['Cuenta'].astype(str).str[:6]
    
    df_final.loc[mask_samuel, 'Grupo_N1'] = '107'
    df_final.loc[mask_samuel, 'Grupo_N2'] = '107.05'
    df_final.loc[mask_samuel, 'Descripcion_Cuenta'] = df_final.loc[mask_samuel, 'Descripcion_Cuenta'] + " (Reclasificado)"

    # Ordenar
    df_final.sort_values(['Grupo_N1', 'Grupo_N2', 'Cuenta'], inplace=True)

    # 4. Construir Reporte
    filas_reporte = []
    
    grupos_n1 = df_final.groupby('Grupo_N1')
    claves_n1_ordenadas = sorted(grupos_n1.groups.keys())

    for codigo_n1 in claves_n1_ordenadas:
        datos_n1 = grupos_n1.get_group(codigo_n1)
        if not codigo_n1[0].isdigit(): continue 

        # --- NIVEL 1 ---
        saldo_n1 = datos_n1['Saldo_Final'].sum()
        nombre_n1 = MAJOR_NAME_MAP.get(codigo_n1, f"Rubro {codigo_n1}")
        filas_reporte.append({'Cuenta': codigo_n1, 'Descripcion': nombre_n1, 'Saldo': saldo_n1, 'Nivel': 1})
        
        # --- NIVEL 2 ---
        grupos_n2 = datos_n1.groupby('Grupo_N2')
        claves_n2_ordenadas = sorted(grupos_n2.groups.keys())
        
        for codigo_n2 in claves_n2_ordenadas:
            datos_n2 = grupos_n2.get_group(codigo_n2)
            
            saldo_n2 = datos_n2['Saldo_Final'].sum()
            nombre_n2 = MAJOR_NAME_MAP.get(codigo_n2, f"Suma {codigo_n2}")
            filas_reporte.append({'Cuenta': codigo_n2, 'Descripcion': nombre_n2, 'Saldo': saldo_n2, 'Nivel': 2})
            
            # --- NIVEL 3 ---
            for _, row in datos_n2.iterrows():
                filas_reporte.append({
                    'Cuenta': row['Cuenta'], 
                    'Descripcion': row['Descripcion_Cuenta'], 
                    'Saldo': row['Saldo_Final'], 
                    'Nivel': 3
                })

        # --- LÓGICA ESPECIAL PARA EL 107 ---
        if codigo_n1 == "107":
            total_107 = saldo_n1
            # Buscar saldo mercancía (107.05.01) dentro del grupo
            row_mercancia = datos_n1[datos_n1['Cuenta'] == '107.05.01']
            val_mercancia = row_mercancia['Saldo_Final'].sum() if not row_mercancia.empty else 0.0
            
            val_calculado = total_107 - val_mercancia
            
            filas_reporte.append({
                'Cuenta': '', 
                'Descripcion': '107 Deudores diversos MENOS 107.05.01 Mercancías Enviadas - No Facturas', 
                'Saldo': val_calculado, 
                'Nivel': 'CALCULO'
            })

        # Espacio separador
        filas_reporte.append({'Cuenta': '', 'Descripcion': '', 'Saldo': np.nan, 'Nivel': ''})

    reporte = pd.DataFrame(filas_reporte)
    
    # Check
    reporte['Es_Cero'] = reporte['Saldo'].apply(
        lambda x: "SI" if pd.notna(x) and abs(x) < 0.01 else ("NO" if pd.notna(x) else "")
    )

    # 5. Exportar a Excel
    nombre_archivo = 'Reporte_Contable_Final.xlsx'
    print(f"Generando Excel: {nombre_archivo}...")
    
    writer = pd.ExcelWriter(nombre_archivo, engine='xlsxwriter')
    
    # EXPORTACIÓN
    columnas_visibles = ['Cuenta', 'Descripcion', 'Saldo', 'Es_Cero']
    reporte[columnas_visibles].to_excel(writer, index=False, sheet_name='Reporte')
    
    workbook  = writer.book
    worksheet = writer.sheets['Reporte']
    
    # --- FORMATOS ---
    currency_fmt = workbook.add_format({'num_format': '$ #,##0.00;[Red]-$ #,##0.00'})
    
    level1_fmt = workbook.add_format({
        'bold': True, 'bg_color': '#FFFF00', 'border': 1,
        'num_format': '$ #,##0.00;[Red]-$ #,##0.00'
    })
    
    level2_fmt = workbook.add_format({
        'bold': True, 'bg_color': '#F2F2F2',
        'num_format': '$ #,##0.00;[Red]-$ #,##0.00'
    })

    calculo_fmt = workbook.add_format({
        'bold': True, 'bg_color': '#DDEBF7', 'border': 1,
        'num_format': '$ #,##0.00;[Red]-$ #,##0.00'
    })

    for row_num, row_data in enumerate(reporte.to_dict('records')):
        excel_row = row_num + 1 
        nivel = row_data['Nivel']
        saldo = row_data['Saldo']
        
        # Reescribimos el Saldo
        if pd.notna(saldo):
            worksheet.write(excel_row, 2, saldo, currency_fmt)
        else:
            worksheet.write(excel_row, 2, "", currency_fmt)
        
        # Aplicar Estilos
        if nivel == 1:
            worksheet.set_row(excel_row, None, level1_fmt)
            worksheet.write(excel_row, 0, row_data['Cuenta'], level1_fmt)
            worksheet.write(excel_row, 1, row_data['Descripcion'], level1_fmt)
            if pd.notna(saldo): worksheet.write(excel_row, 2, saldo, level1_fmt)
        
        elif nivel == 2:
            worksheet.set_row(excel_row, None, level2_fmt)
            worksheet.write(excel_row, 0, row_data['Cuenta'], level2_fmt)
            worksheet.write(excel_row, 1, row_data['Descripcion'], level2_fmt)
            if pd.notna(saldo): worksheet.write(excel_row, 2, saldo, level2_fmt)

        elif nivel == 'CALCULO':
            worksheet.set_row(excel_row, None, calculo_fmt)
            worksheet.write(excel_row, 0, row_data['Cuenta'], calculo_fmt)
            worksheet.write(excel_row, 1, row_data['Descripcion'], calculo_fmt)
            if pd.notna(saldo): worksheet.write(excel_row, 2, saldo, calculo_fmt)

    worksheet.set_column('A:A', 15)
    worksheet.set_column('B:B', 60)
    worksheet.set_column('C:C', 18)
    worksheet.set_column('D:D', 10)

    writer.close()
    print(f"¡Listo! Archivo completado exitosamente: {nombre_archivo}")

if __name__ == "__main__":
    procesar_contabilidad()