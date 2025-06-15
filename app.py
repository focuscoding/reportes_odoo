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

# Función para obtener facturas
def get_facturas(fecha_inicio, fecha_fin):
    domain = [
        ("invoice_date", ">=", fecha_inicio),
        ("invoice_date", "<=", fecha_fin),
        ("move_type", "=", "out_invoice"),
        ("state", "=", "posted")
    ]

    fields = [
        "name", "invoice_date", "partner_id", "amount_total", "invoice_line_ids"
    ]

    facturas = models.execute_kw(
        db, uid, password,
        "account.move", "search_read",
        [domain], {"fields": fields, "limit": 1000}
    )

    # Obtener nombres de los clientes
    for f in facturas:
        f["partner_name"] = f["partner_id"][1] if f["partner_id"] else ""

    return pd.DataFrame(facturas)

# Streamlit app
st.set_page_config(layout="wide")
st.title("Reporte de Facturas por Laboratorio")

# Selección de fechas
col1, col2 = st.columns(2)
with col1:
    fecha_inicio = st.date_input("Fecha de inicio", datetime.now() - timedelta(days=30))
with col2:
    fecha_fin = st.date_input("Fecha de fin", datetime.now())

# Botón para cargar facturas
if st.button("📥 Obtener facturas"):
    with st.spinner("Obteniendo facturas desde Odoo..."):
        df = get_facturas(fecha_inicio.strftime("%Y-%m-%d"), fecha_fin.strftime("%Y-%m-%d"))

    if df.empty:
        st.error("❌ No se encontraron facturas en el rango de fechas indicado.")
    else:
        st.success(f"✅ {len(df)} facturas obtenidas.")

        # Entrada de texto para filtrar proveedores por coincidencia parcial
        filtro_texto = st.text_input("🔍 Filtrar proveedores por nombre (parcial):")

        if filtro_texto:
            df_filtrado = df[df['partner_name'].str.contains(filtro_texto, case=False, na=False)]
        else:
            df_filtrado = df.copy()

        proveedores_unicos = sorted(df_filtrado['partner_name'].dropna().unique())
        seleccionados = st.multiselect("Selecciona uno o más proveedores:", proveedores_unicos)

        if seleccionados:
            df_filtrado = df_filtrado[df_filtrado['partner_name'].isin(seleccionados)]

            if df_filtrado.empty:
                st.error("❌ No se encontraron facturas para los proveedores seleccionados.")
            else:
                st.success(f"✅ {len(df_filtrado)} facturas encontradas para los proveedores seleccionados.")
                st.dataframe(df_filtrado)

                # Función para generar el enlace de descarga
                def generar_enlace_descarga(df, nombre_archivo):
                    output = BytesIO()
                    df.to_excel(output, index=False, engine='openpyxl')
                    output.seek(0)
                    b64 = base64.b64encode(output.read()).decode()
                    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{nombre_archivo}">📤 Descargar archivo Excel</a>'
                    return href

                nombre_excel = f"facturas_filtradas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                st.markdown(generar_enlace_descarga(df_filtrado, nombre_excel), unsafe_allow_html=True)
        else:
            st.info("Selecciona al menos un proveedor para aplicar el filtro.")

