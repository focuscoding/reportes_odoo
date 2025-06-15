import streamlit as st
import pandas as pd
import base64
from datetime import datetime, timedelta
from io import BytesIO
import xmlrpc.client


# Conexión Odoo
url = st.secrets["odoo"]["url"]
db = st.secrets["odoo"]["db"]
username = st.secrets["odoo"]["username"]
password = st.secrets["odoo"]["password"]

# Autenticación con Odoo
common = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common")
uid = common.authenticate(db, username, password, {})
models = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object")

# Palabras laboratorio filtro
laboratorios_permitidos = ["Santé", "Leti", "Calox", "Oftalmi", "Valmor", "Megalabs"]

def get_facturas_filtradas(fecha_inicio, fecha_fin):
    # Buscar facturas entre fechas, solo partner_name que contenga "Farmago"
    domain = [
        ("invoice_date", ">=", fecha_inicio),
        ("invoice_date", "<=", fecha_fin),
        ("move_type", "=", "out_invoice"),
        ("state", "=", "posted"),
        # Para filtrar partner_id por nombre con "Farmago" usamos un filtro después, ya que Odoo no filtra texto parcialmente directo
    ]

    fields = [
        "name", "invoice_date", "partner_id", "amount_total", "invoice_line_ids"
    ]

    facturas = models.execute_kw(
        db, uid, password,
        "account.move", "search_read",
        [domain], {"fields": fields, "limit": 1000}
    )

    # Filtrar facturas que contengan "farmago" en partner_name
    facturas_filtradas = []
    for f in facturas:
        partner_name = f["partner_id"][1] if f["partner_id"] else ""
        if "farmago" in partner_name.lower():
            facturas_filtradas.append(f)

    if not facturas_filtradas:
        return pd.DataFrame()  # vacío si nada

    # Obtener IDs de las facturas filtradas
    factura_ids = [f["id"] for f in facturas_filtradas if "id" in f]

    # Para obtener datos de las líneas de factura, hacemos otro llamado para traer fields extendidos
    # O directamente obtenemos las líneas y sus campos laboratorio

    # Primero extraemos todas las líneas con sus campos:
    # Para ello, obtener todas las líneas de esas facturas (invoice_line_ids)

    # Obtener todas las líneas con sus datos relevantes:
    lineas = models.execute_kw(
        db, uid, password,
        "account.move.line", "search_read",
        [[("move_id", "in", factura_ids)]],
        {"fields": ["move_id", "name", "quantity", "price_unit", "price_total", "laboratory_name"]}
    )

    # Filtrar líneas que tengan laboratory_name con alguna palabra de laboratorios_permitidos
    lineas_filtradas = []
    for linea in lineas:
        lab_name = linea.get("laboratory_name") or ""
        lab_name_lower = lab_name.lower()
        if any(lab.lower() in lab_name_lower for lab in laboratorios_permitidos):
            lineas_filtradas.append(linea)

    if not lineas_filtradas:
        return pd.DataFrame()  # No hay líneas que cumplan filtro

    # Armar DataFrame con líneas filtradas y datos de factura
    # Necesitamos info de factura para cada línea:
    # Creamos un dict para acceder rápido al invoice_date y partner_name por move_id
    factura_info = {}
    for f in facturas_filtradas:
        factura_info[f["id"]] = {
            "invoice_date": f["invoice_date"],
            "partner_name": f["partner_id"][1] if f["partner_id"] else "",
            "invoice_name": f["name"],
            "amount_total": f["amount_total"]
        }

    datos = []
    for linea in lineas_filtradas:
        move_id = linea["move_id"][0] if linea["move_id"] else None
        info_factura = factura_info.get(move_id, {})
        datos.append({
            "Factura": info_factura.get("invoice_name", ""),
            "Fecha": info_factura.get("invoice_date", ""),
            "Cliente": info_factura.get("partner_name", ""),
            "Producto": linea.get("name", ""),
            "Cantidad": linea.get("quantity", 0),
            "Precio Unitario": linea.get("price_unit", 0.0),
            "Total Línea": linea.get("price_total", 0.0),
            "Laboratorio": linea.get("laboratory_name", "")
        })

    df = pd.DataFrame(datos)
    return df

# Streamlit app
st.set_page_config(layout="wide")
st.title("Reporte de facturas Farmago con filtro por laboratorios")

col1, col2 = st.columns(2)
with col1:
    fecha_inicio = st.date_input("Fecha inicio", datetime.now() - timedelta(days=30))
with col2:
    fecha_fin = st.date_input("Fecha fin", datetime.now())

if st.button("📥 Obtener facturas"):
    with st.spinner("Consultando facturas en Odoo..."):
        df = get_facturas_filtradas(fecha_inicio.strftime("%Y-%m-%d"), fecha_fin.strftime("%Y-%m-%d"))

    if df.empty:
        st.error("❌ No se encontraron facturas ni líneas con los filtros indicados.")
    else:
        st.success(f"✅ Se encontraron {len(df)} líneas de factura que cumplen los filtros.")
        st.dataframe(df)

        # Función para descarga Excel
        def generar_enlace_descarga(df, nombre_archivo):
            output = BytesIO()
            df.to_excel(output, index=False, engine='openpyxl')
            output.seek(0)
            b64 = base64.b64encode(output.read()).decode()
            href = f'<a href="data:application/octet-stream;base64,{b64}" download="{nombre_archivo}">📤 Descargar Excel</a>'
            return href

        nombre_excel = f"facturas_farmago_filtradas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.markdown(generar_enlace_descarga(df, nombre_excel), unsafe_allow_html=True)
