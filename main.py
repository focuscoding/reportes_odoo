# ✅ Instalamos librerías (solo la primera vez en Replit o Colab)
# !pip install pandas openpyxl

import xmlrpc.client
import pandas as pd

# Datos de conexión
url = 'https://drogueriablv.odoo.com'
db = 'odoo-tecnored-drogueriablv-main-7701393'
username = 'ahidalgo.blv@gmail.com'
password = 'ahidalgo123'

# Conexión
common = xmlrpc.client.ServerProxy(f'{url}/xmlrpc/2/common')
uid = common.authenticate(db, username, password, {})
if not uid:
    print("❌ No se pudo autenticar.")
    exit()

models = xmlrpc.client.ServerProxy(f'{url}/xmlrpc/2/object')
print(f"✅ Conectado a Odoo con UID: {uid}")

# ⚙️ Rango de fechas
fecha_inicio = '2025-06-01'
fecha_fin = '2025-06-06'

# 🔍 Buscar facturas que contengan 'Farmago' en el nombre del cliente
print(f"🔍 Buscando facturas de Farmago entre {fecha_inicio} y {fecha_fin}...")
facturas = models.execute_kw(db, uid, password,
    'account.move', 'search_read',
    [[
        ['move_type', '=', 'out_invoice'],
        ['invoice_date', '>=', fecha_inicio],
        ['invoice_date', '<=', fecha_fin],
        ['invoice_partner_display_name', 'ilike', 'Farmago']
    ]],
    {'fields': ['id', 'name', 'invoice_number_next', 'invoice_date', 'invoice_number_next',
                 'invoice_number_control', 'invoice_partner_display_name',
                 'os_currency_rate'], 'limit': 5000})

print(f"📊 Se encontraron {len(facturas)} facturas.")

if not facturas:
    print("❌ No se encontraron facturas de Farmago en el rango de fechas especificado.")
    print(f"Rango buscado: {fecha_inicio} a {fecha_fin}")
    
    # Buscar si existen facturas en general en ese rango
    print("🔍 Verificando si existen facturas en general en ese rango...")
    test_facturas = models.execute_kw(db, uid, password,
        'account.move', 'search_read',
        [[
            ['move_type', '=', 'out_invoice'],
            ['invoice_date', '>=', fecha_inicio],
            ['invoice_date', '<=', fecha_fin]
        ]],
        {'fields': ['id'], 'limit': 10})
    
    print(f"📊 Se encontraron {len(test_facturas)} facturas en total en ese rango.")
    print("Script terminado.")
    exit(0)

# 🔎 Obtener IDs de facturas
factura_ids = [f['id'] for f in facturas]
print(f"🔎 Procesando {len(factura_ids)} facturas...")

# 🔍 Buscar líneas de facturas relacionadas
print("📋 Obteniendo líneas de facturas...")
lineas = models.execute_kw(db, uid, password,
    'account.move.line', 'search_read',
    [[['move_id', 'in', factura_ids]]],
    {'fields': ['move_id', 'product_id', 'name', 'quantity',            'price_unit','price_subtotal', 'price_unit_rate',                  'price_subtotal_rate','discount'],
     'limit': 50000})

print(f"📊 Se encontraron {len(lineas)} líneas de factura.")

# 🔎 Obtener IDs de productos para buscar laboratorios
product_ids = list(set([l['product_id'][0] for l in lineas if l['product_id']]))
print(f"🔎 Procesando {len(product_ids)} productos únicos...")

# Procesar productos en lotes para evitar timeouts
batch_size = 500
product_templates = []

for i in range(0, len(product_ids), batch_size):
    batch = product_ids[i:i+batch_size]
    print(f"📦 Procesando lote {i//batch_size + 1}/{(len(product_ids)-1)//batch_size + 1} ({len(batch)} productos)...")
    
    batch_templates = models.execute_kw(db, uid, password,
        'product.product', 'read',
        [batch], {'fields': ['id', 'product_tmpl_id']})
    
    product_templates.extend(batch_templates)

template_ids = list(set([pt['product_tmpl_id'][0] for pt in product_templates]))
print(f"🔎 Procesando {len(template_ids)} plantillas únicas...")

# Procesar plantillas en lotes
templates_data = []

for i in range(0, len(template_ids), batch_size):
    batch = template_ids[i:i+batch_size]
    print(f"📋 Procesando lote {i//batch_size + 1}/{(len(template_ids)-1)//batch_size + 1} ({len(batch)} plantillas)...")
    
    batch_data = models.execute_kw(db, uid, password,
        'product.template', 'read',
        [batch], {'fields': ['id', 'laboratory_name', 'default_code', 'supplier_code']})
    
    templates_data.extend(batch_data)

# 🔄 Crear diccionario template_id → laboratorio_name
template_dict = {
    pt['id']: {
        'laboratorio': pt.get('laboratory_name', ''),
        'codigo': pt.get('default_code', ''),
        'codigo_proveedor': pt.get('supplier_code', ''),
    }
    for pt in templates_data
}

# 🔄 Crear diccionario product_id → template_id
prod_template_dict = {pt['id']: pt['product_tmpl_id'][0] for pt in product_templates}

# 🔄 Preparar data final
print("🔄 Procesando datos finales...")
factura_dict = {f['id']: f for f in facturas}
lineas_filtradas = []
lab_claves = ['santé', 'leti', 'calox', 'oftalmi', 'valmor', 'megalabs']

print(f"🔍 Filtrando líneas por laboratorios: {', '.join(lab_claves)}")
for i, linea in enumerate(lineas):
    if i % 1000 == 0 and i > 0:
        print(f"📊 Procesadas {i}/{len(lineas)} líneas...")
    if not linea['product_id']:
        continue
    product_id = linea['product_id'][0]
    template_id = prod_template_dict.get(product_id)
    lab_info = template_dict.get(template_id, {'laboratorio': '',              'codigo': '', 'codigo_proveedor':''})
    laboratorio = lab_info['laboratorio'][1].lower() if                        isinstance(lab_info['laboratorio'], list) else                             lab_info['laboratorio'].lower()
    codigo = lab_info['codigo']

    if any(clave in laboratorio for clave in lab_claves):
        factura = factura_dict[linea['move_id'][0]]
        lineas_filtradas.append({
            'Fecha Factura': factura['invoice_date'],
            'Cliente': factura['invoice_partner_display_name'],
            # 'Factura': factura['name'],
            'Nro. Factura': factura['invoice_number_next'],
            # 'Nro. Control': factura['invoice_number_control'],
            'Código de Barras': codigo,
            # 'Tasa Cambio': factura.get('os_currency_rate', 1),
            'Producto': linea['name'],
            'Laboratorio': laboratorio.upper(),
            'Código Laboratorio': lab_info['codigo_proveedor'],
            'Cantidad': linea['quantity'],
            'Precio Unitario': linea['price_unit'],
            'Descuento': linea['discount'],
            'Subtotal': linea['price_subtotal'],
            })

print(f"✅ Filtrado completado. {len(lineas_filtradas)} líneas coinciden con laboratorios objetivo.")

# Proveedores seleccionados manualmente
proveedores_seleccionados = ['Oftalmi', 'Leti', 'Calox']
proveedores_seleccionados = [p.strip().lower() for p in proveedores_seleccionados]

# Crear DataFrame original
print("📊 Creando DataFrame...")
df = pd.DataFrame(lineas_filtradas)

if df.empty:
    print("❌ No hay datos para procesar.")
    exit()

print(f"📋 DataFrame creado con {len(df)} registros.")
print(f"🏭 Laboratorios encontrados: {df['Laboratorio'].unique()}")

df['Laboratorio_lower'] = df['Laboratorio'].str.lower().str.strip()

# Use contains logic instead of exact match
mask = df['Laboratorio_lower'].str.contains('|'.join(proveedores_seleccionados), na=False)
df_filtrado = df[mask].drop(columns='Laboratorio_lower')

print(f"🔍 Después del filtro por proveedores seleccionados: {len(df_filtrado)} registros")
print(f"🏭 Proveedores a procesar: {df_filtrado['Laboratorio'].unique()}")

if df_filtrado.empty:
    print("❌ No hay datos después del filtro por proveedores.")
    print(f"Proveedores buscados: {proveedores_seleccionados}")
    exit()

# Exportar 1 archivo por proveedor
print("📁 Iniciando generación de archivos Excel...")
for proveedor in df_filtrado['Laboratorio'].unique():
    print(f"📋 Procesando proveedor: {proveedor}")
    df_proveedor = df_filtrado[df_filtrado['Laboratorio'] == proveedor].copy()
    print(f"📊 Registros para {proveedor}: {len(df_proveedor)}")

    archivo_excel = f'facturas_{proveedor}.xlsx'
    print(f"📁 Generando archivo: {archivo_excel}")
    with pd.ExcelWriter(archivo_excel, engine='xlsxwriter') as writer:
        df_proveedor.to_excel(writer, index=False, sheet_name='Facturas Filtradas')

        workbook = writer.book
        worksheet = writer.sheets['Facturas Filtradas']

        # Formatos
        money_format = workbook.add_format({
            'num_format': '_-* "Bs.S "* #,##0.00_-;-_* "Bs.S "* -#,##0.00_-;_-* "Bs.S "* "-"??_-;_-@_-',
            'align': 'right'
        })
        header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'top'
        })
        total_format = workbook.add_format({'bold': True})
        total_money_format = workbook.add_format({
            'bold': True,
            'num_format': '_-* "Bs.S "* #,##0.00_-;-_* "Bs.S "* -#,##0.00_-;_-* "Bs.S "* "-"??_-;_-@_-',
            'align': 'right'
        })

        # Posiciones
        col_i = df_proveedor.columns.get_loc('Precio Unitario')
        col_k = df_proveedor.columns.get_loc('Subtotal')
        col_l = len(df_proveedor.columns)  # Nueva columna "Monto NC"

        # Insertar encabezado "Monto NC"
        worksheet.write(0, col_l, 'Monto NC', header_format)

        # Cálculo fórmula y formato Bs.S
        for row in range(1, len(df_proveedor) + 1):
            worksheet.write_formula(row, col_l, f'=J{row + 1} * K{row + 1} / 100', money_format)

        for row in range(1, len(df_proveedor) + 1):
            valor_i = df_proveedor.iloc[row - 1, col_i]
            if pd.notnull(valor_i):
                worksheet.write_number(row, col_i, valor_i, money_format)
            valor_k = df_proveedor.iloc[row - 1, col_k]
            if pd.notnull(valor_k):
                worksheet.write_number(row, col_k, valor_k, money_format)

        # Formatear todos los encabezados
        for col_idx, col_name in enumerate(df_proveedor.columns):
            worksheet.write(0, col_idx, col_name, header_format)

        # Totales
        total_row = len(df_proveedor) + 1
        worksheet.write(total_row, 6, 'Total Unidades', total_format)
        worksheet.write_formula(total_row, 6, f'=SUM(G2:G{total_row})', total_format)
        worksheet.write_formula(total_row, 7, f'=SUM(H2:H{total_row})', total_format)  # sin formato Bs
        worksheet.write(total_row, 10, 'Total NC', total_format)
        worksheet.write_formula(total_row, 11, f'=SUM(L2:L{total_row})', total_money_format)

        # Ajuste de ancho
        columnas = list(df_proveedor.columns) + ['Monto NC']
        for idx, col in enumerate(columnas):
            if col in ['Precio Unitario', 'Subtotal', 'Monto NC']:
                if col == 'Monto NC':
                    max_val = max(df_proveedor['Precio Unitario'] * df_proveedor['Subtotal'] / 100)
                    max_len_str = f"Bs.S {max_val:,.2f}"
                else:
                    max_val = df_proveedor[col].max()
                    max_len_str = f"Bs.S {max_val:,.2f}"
                max_len = max(len(col), len(max_len_str)) + 4
            elif col == 'Cantidad':
                max_val = df_proveedor[col].sum()
                max_len_str = f"{max_val:,.2f}"
                max_len = max(len(col), len(max_len_str)) + 2
            else:
                max_len = max(df_proveedor[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(idx, idx, max_len)

    print(f"✅ Archivo generado: {archivo_excel}")

print("🎉 Proceso completado exitosamente!")
