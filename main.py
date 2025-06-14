import streamlit as st
import xmlrpc.client
import pandas as pd
from datetime import date

# Título de la app
st.title("🧾 Generador de Reportes por Proveedor")

# Selección de fechas
definir_fecha_inicio = st.date_input("📅 Fecha de inicio", value=date.today())
definir_fecha_fin = st.date_input("📅 Fecha de fin", value=date.today())

if definir_fecha_inicio > definir_fecha_fin:
    st.warning("⚠️ La fecha de inicio no puede ser posterior a la fecha de fin.")
    st.stop()

# Convirtiendo fechas a string formato Odoo
date_start = definir_fecha_inicio.strftime('%Y-%m-%d')
date_end = definir_fecha_fin.strftime('%Y-%m-%d')

# Conexión Odoo
url = st.secrets["odoo"]["url"]
db = st.secrets["odoo"]["db"]
username = st.secrets["odoo"]["username"]
password = st.secrets["odoo"]["password"]


common = xmlrpc.client.ServerProxy(f'{url}/xmlrpc/2/common')
uid = common.authenticate(db, username, password, {})

if not uid:
    st.error("❌ No se pudo autenticar con Odoo.")
    st.stop()

models = xmlrpc.client.ServerProxy(f'{url}/xmlrpc/2/object')
st.success(f"✅ Conectado a Odoo con UID: {uid}")

# Buscar facturas de Farmago en rango de fechas
st.write(f"🔍 Buscando facturas de Farmago entre {date_start} y {date_end}...")
facturas = models.execute_kw(db, uid, password,
    'account.move', 'search_read',
    [[
        ['move_type', '=', 'out_invoice'],
        ['invoice_date', '>=', date_start],
        ['invoice_date', '<=', date_end],
        ['invoice_partner_display_name', 'ilike', 'Farmago']
    ]],
    {'fields': ['id', 'name', 'invoice_date', 'invoice_number_next',
                'invoice_partner_display_name', 'os_currency_rate'], 'limit': 5000})

if not facturas:
    st.warning("❌ No se encontraron facturas de Farmago.")
    st.stop()

factura_ids = [f['id'] for f in facturas]
lineas = models.execute_kw(db, uid, password,
    'account.move.line', 'search_read',
    [[['move_id', 'in', factura_ids]]],
    {'fields': ['move_id', 'product_id', 'name', 'quantity',
                'price_unit','price_subtotal', 'price_unit_rate',
                'price_subtotal_rate','discount'], 'limit': 50000})

product_ids = list(set([l['product_id'][0] for l in lineas if l['product_id']]))
batch_size = 500
product_templates = []

for i in range(0, len(product_ids), batch_size):
    batch = product_ids[i:i+batch_size]
    batch_templates = models.execute_kw(db, uid, password,
        'product.product', 'read', [batch],
        {'fields': ['id', 'product_tmpl_id']})
    product_templates.extend(batch_templates)

template_ids = list(set([pt['product_tmpl_id'][0] for pt in product_templates]))
templates_data = []
for i in range(0, len(template_ids), batch_size):
    batch = template_ids[i:i+batch_size]
    batch_data = models.execute_kw(db, uid, password,
        'product.template', 'read', [batch],
        {'fields': ['id', 'laboratory_name', 'default_code', 'supplier_code']})
    templates_data.extend(batch_data)

template_dict = {
    pt['id']: {
        'laboratorio': pt.get('laboratory_name', ''),
        'codigo': pt.get('default_code', ''),
        'codigo_proveedor': pt.get('supplier_code', ''),
    } for pt in templates_data
}

prod_template_dict = {pt['id']: pt['product_tmpl_id'][0] for pt in product_templates}

factura_dict = {f['id']: f for f in facturas}
lineas_filtradas = []
lab_claves = ['santé', 'leti', 'calox', 'oftalmi', 'valmor', 'megalabs']

for linea in lineas:
    if not linea['product_id']:
        continue
    product_id = linea['product_id'][0]
    template_id = prod_template_dict.get(product_id)
    lab_info = template_dict.get(template_id, {'laboratorio': '', 'codigo': '', 'codigo_proveedor': ''})
    laboratorio = lab_info['laboratorio'][1].lower() if isinstance(lab_info['laboratorio'], list) else lab_info['laboratorio'].lower()
    codigo = lab_info['codigo']
    if any(clave in laboratorio for clave in lab_claves):
        factura = factura_dict[linea['move_id'][0]]
        lineas_filtradas.append({
            'Fecha Factura': factura['invoice_date'],
            'Cliente': factura['invoice_partner_display_name'],
            'Nro. Factura': factura['invoice_number_next'],
            'Código de Barras': codigo,
            'Producto': linea['name'],
            'Laboratorio': laboratorio.upper(),
            'Código Laboratorio': lab_info['codigo_proveedor'],
            'Cantidad': linea['quantity'],
            'Precio Unitario': linea['price_unit'],
            'Descuento': linea['discount'],
            'Subtotal': linea['price_subtotal'],
        })

if not lineas_filtradas:
    st.warning("❌ No hay líneas de factura que coincidan con los laboratorios seleccionados.")
    st.stop()

df = pd.DataFrame(lineas_filtradas)
df['Laboratorio_lower'] = df['Laboratorio'].str.lower().str.strip()
proveedores_seleccionados = ['Oftalmi', 'Leti', 'Calox']
proveedores_seleccionados = [p.lower() for p in proveedores_seleccionados]

mask = df['Laboratorio_lower'].str.contains('|'.join(proveedores_seleccionados), na=False)
df_filtrado = df[mask].drop(columns='Laboratorio_lower')

if df_filtrado.empty:
    st.warning("❌ No hay datos para los proveedores seleccionados.")
    st.stop()

st.write("📥 Datos filtrados por proveedor:")
st.dataframe(df_filtrado)

for proveedor in df_filtrado['Laboratorio'].unique():
    df_proveedor = df_filtrado[df_filtrado['Laboratorio'] == proveedor]
    archivo_excel = f"facturas_{proveedor}.xlsx"
    df_proveedor.to_excel(archivo_excel, index=False)
    with open(archivo_excel, "rb") as file:
        st.download_button(
            label=f"⬇️ Descargar Excel: {proveedor}",
            data=file,
            file_name=archivo_excel,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
