import streamlit as st
import pandas as pd
import xmlrpc.client
from datetime import date
import uuid
import os

st.title("📊 Reportes Farmago")
st.header("📁 Reportes a proveedores")

# Elegir proveedores
st.subheader("Elegir Proveedores")
opciones_proveedores = ['Todos', 'Santé', 'Leti', 'Calox', 'Oftalmi', 'Valmor', 'Megalabs','Siegfried']
seleccionados = st.multiselect("Selecciona uno o varios laboratorios", opciones_proveedores, default=['Leti', 'Calox', 'Megalabs', 'Siegfried'])

# Procesar selección
if 'Todos' in seleccionados:
    proveedores_seleccionados = ['santé', 'leti', 'calox', 'oftalmi', 'valmor', 'megalabs']
else:
    proveedores_seleccionados = [p.strip().lower() for p in seleccionados]

# Selección de fechas
st.subheader("Seleccionar rango de fechas")
fecha_inicio = st.date_input("Fecha inicio", value=date(2025, 6, 1))
fecha_fin = st.date_input("Fecha fin", value=date(2025, 6, 6))

# Convertir fechas a string
fecha_inicio_str = fecha_inicio.strftime('%Y-%m-%d')
fecha_fin_str = fecha_fin.strftime('%Y-%m-%d')

if st.button("Generar Reporte"):
    archivos_generados = []
    st.session_state.archivos_generados = []

    # Conexión Odoo
    url = st.secrets["odoo"]["url"]
    db = st.secrets["odoo"]["db"]
    username = st.secrets["odoo"]["username"]
    password = st.secrets["odoo"]["password"]

    common = xmlrpc.client.ServerProxy(f'{url}/xmlrpc/2/common')
    uid = common.authenticate(db, username, password, {})
    if not uid:
        st.error("❌ No se pudo autenticar en Odoo.")
        st.stop()

    models = xmlrpc.client.ServerProxy(f'{url}/xmlrpc/2/object')

    facturas = models.execute_kw(db, uid, password,
        'account.move', 'search_read',
        [[
            ['move_type', '=', 'out_invoice'],
            ['invoice_date', '>=', fecha_inicio_str],
            ['invoice_date', '<=', fecha_fin_str],
            ['invoice_partner_display_name', 'ilike', 'Farmago']
        ]],
        {'fields': ['id', 'name', 'invoice_number_next', 'invoice_date', 'invoice_number_next',
                     'invoice_number_control', 'invoice_partner_display_name',
                     'os_currency_rate'], 'limit': 5000})

    if not facturas:
        st.warning("❌ No se encontraron facturas de Farmago en el rango de fechas especificado.")
        st.stop()

    factura_ids = [f['id'] for f in facturas]
    lineas = models.execute_kw(db, uid, password,
        'account.move.line', 'search_read',
        [[['move_id', 'in', factura_ids]]],
        {'fields': ['move_id', 'product_id', 'name', 'quantity','price_unit','price_subtotal','price_unit_rate','price_subtotal_rate','discount'],
         'limit': 50000})

    product_ids = list(set([l['product_id'][0] for l in lineas if l['product_id']]))
    batch_size = 500
    product_templates = []
    for i in range(0, len(product_ids), batch_size):
        batch = product_ids[i:i+batch_size]
        batch_templates = models.execute_kw(db, uid, password,
            'product.product', 'read',
            [batch], {'fields': ['id', 'product_tmpl_id']})
        product_templates.extend(batch_templates)

    template_ids = list(set([pt['product_tmpl_id'][0] for pt in product_templates]))
    templates_data = []
    for i in range(0, len(template_ids), batch_size):
        batch = template_ids[i:i+batch_size]
        batch_data = models.execute_kw(db, uid, password,
            'product.template', 'read',
            [batch], {'fields': ['id', 'laboratory_name', 'default_code', 'supplier_code']})
        templates_data.extend(batch_data)

    template_dict = {
        pt['id']: {
            'laboratorio': pt.get('laboratory_name', ''),
            'codigo': pt.get('default_code', ''),
            'codigo_proveedor': pt.get('supplier_code', ''),
        }
        for pt in templates_data
    }

    prod_template_dict = {pt['id']: pt['product_tmpl_id'][0] for pt in product_templates}
    factura_dict = {f['id']: f for f in facturas}
    lineas_filtradas = []

    for linea in lineas:
        if not linea['product_id']:
            continue
        product_id = linea['product_id'][0]
        template_id = prod_template_dict.get(product_id)
        lab_info = template_dict.get(template_id, {'laboratorio': '', 'codigo': '', 'codigo_proveedor': ''})
        laboratorio = lab_info['laboratorio'][1].lower() if isinstance(lab_info['laboratorio'], list) else lab_info['laboratorio'].lower()
        if any(clave in laboratorio for clave in proveedores_seleccionados):
            factura = factura_dict[linea['move_id'][0]]
            lineas_filtradas.append({
                'Fecha Factura': factura['invoice_date'],
                'Cliente': factura['invoice_partner_display_name'],
                'Nro. Factura': factura['invoice_number_next'],
                'Código de Barras': lab_info['codigo'],
                'Producto': linea['name'],
                'Laboratorio': laboratorio.upper(),
                'Código Laboratorio': lab_info['codigo_proveedor'],
                'Cantidad': linea['quantity'],
                'Precio Unitario': linea['price_unit'],
                'Descuento': linea['discount'],
                'Subtotal': linea['price_subtotal'],
            })

    df = pd.DataFrame(lineas_filtradas)
    if df.empty:
        st.warning("❌ No hay datos después del filtro por proveedores.")
        st.stop()

    def convertir_a_float(valor):
        if isinstance(valor, str):
            return float(valor.replace(",", ""))
        elif pd.notnull(valor):
            return float(valor)
        return 0.0

    for proveedor in df['Laboratorio'].unique():
        df_proveedor = df[df['Laboratorio'] == proveedor].copy()
        archivo_excel = f'facturas_{proveedor}.xlsx'
        with pd.ExcelWriter(archivo_excel, engine='xlsxwriter') as writer:
            df_proveedor.to_excel(writer, sheet_name='Facturas Filtradas', index=False)
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

            worksheet.write(0, col_l, 'Monto NC', header_format)

            # Fórmulas y formatos
            for row in range(1, len(df_proveedor) + 1):
                worksheet.write_formula(row, col_l, f'=J{row + 1} * K{row + 1} / 100', money_format)

                valor_i = df_proveedor.iloc[row - 1, col_i]
                if pd.notnull(valor_i):
                    worksheet.write_number(row, col_i, convertir_a_float(valor_i), money_format)

                valor_k = df_proveedor.iloc[row - 1, col_k]
                if pd.notnull(valor_k):
                    worksheet.write_number(row, col_k, convertir_a_float(valor_k), money_format)

            # Encabezados
            for col_idx, col_name in enumerate(df_proveedor.columns):
                worksheet.write(0, col_idx, col_name, header_format)

            # Totales
            total_row = len(df_proveedor) + 1
            worksheet.write(total_row, 6, 'Total Unidades', total_format)
            worksheet.write_formula(total_row, 7, f'=SUM(H2:H{total_row})', total_format)
            worksheet.write(total_row, 10, 'Total NC', total_format)
            worksheet.write_formula(total_row, 11, f'=SUM(L2:L{total_row})', total_money_format)

            # Ajuste de ancho de columnas
            columnas = list(df_proveedor.columns) + ['Monto NC']
            for idx, col in enumerate(columnas):
                if col in ['Precio Unitario', 'Subtotal', 'Monto NC']:
                    if col == 'Monto NC':
                        try:
                            max_val = max(df_proveedor['Precio Unitario'].apply(convertir_a_float) * df_proveedor['Subtotal'].apply(convertir_a_float) / 100)
                            max_len_str = f"Bs.S {max_val:,.2f}"
                        except:
                            max_len_str = "Bs.S 0.00"
                    else:
                        max_val = df_proveedor[col].apply(convertir_a_float).max()
                        max_len_str = f"Bs.S {max_val:,.2f}"
                    max_len = max(len(col), len(max_len_str)) + 4
                elif col == 'Cantidad':
                    max_val = df_proveedor[col].sum()
                    max_len_str = f"{max_val:,.2f}"
                    max_len = max(len(col), len(max_len_str)) + 2
                else:
                    max_len = max(df_proveedor[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(idx, idx, max_len)

        archivos_generados.append(archivo_excel)
      
      
    
    st.success('✅ Archivos generados correctamente')
    for fn in archivos_generados:
        with open(fn, "rb") as f:
            st.download_button(
                label=f"📥 Descargar {fn}",
                data=f,
                file_name=fn,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=str(uuid.uuid4())  # ▷ cada llamada obtiene un identificador único
            )

            