import streamlit as st
import pandas as pd
import xmlrpc.client
from datetime import date, timedelta, datetime
import uuid
import os
import requests
from io import BytesIO
import urllib.parse
import json
import pytz

# Función para enviar datos a Google Sheets
def enviar_a_google_sheets(data):
    script_url = st.secrets["google"]["script_url"]
    try:
        response = requests.post(
            script_url,
            data=json.dumps(data),
            headers={'Content-Type': 'application/json'}
        )
        return response.status_code == 200
    except Exception as e:
        st.error(f"Error al enviar a Google Sheets: {e}")
        return False


# URL de descarga directa del archivo de OneDrive
onedrive_url = st.secrets["odoo"]["onedrive"]

# Descargar el archivo
response = requests.get(onedrive_url)
panel_df = pd.read_excel(BytesIO(response.content), sheet_name=0)

# Asegúrate de que las columnas necesarias estén presentes
# Supongamos que las columnas son: B = 'Código de Barras', K = 'Descuento'
panel_df = panel_df.rename(columns={
    panel_df.columns[1]: 'Código de Barras',
    panel_df.columns[10]: 'Descuento Panel'
})

st.image("./logo/shakira_pc.jpeg", use_container_width=True )

st.title("📊 Reportes Farmago")
st.header("📁 Reportes a proveedores")
if "archivos_generados" not in st.session_state:
    st.session_state.archivos_generados = []

 # 🔄 Diccionario de mapeo local solo para esta sección
mapa_laboratorios = {
    'LABORATORIOS LETI, S.A.V.': 'Leti',
    'CALOX INTERNATIONAL, C.A.': 'Calox',
    'MEGALABS VZL, C.A.': 'Megalabs',
    'LABORATORIOS SIEGFRIED S.A.': 'Siegfried',
    'LABORATORIOS VALMOR, C.A.': 'Valmor',
    'LABORATORIOS L.O. OFTALMI, C.A': 'Oftalmi',
    'LABORATORIOS LA SANTÉ C,A.': 'Sante',
}


# Diccionario de destinatarios por laboratorio
correos_cc_global = ['staddeo@drogueriablv.com', 'kmontero@drogueriablv.com']
emails_por_laboratorio = {
    'sante': ['maria.herrera@pharmetiquelabs.com.ve'],
    'leti': ['asdrubal.mosqueda@grupoleti.com','miller.guerra@grupoleti.com',],
    'calox': ['dmartinez@calox.com'],
    'oftalmi': ['gerente.mercadeo@oftalmi.com','gerente.caracas@oftalmi.com','lourdes.defreitas@oftalmi.com'],
    'valmor': ['dvaleri@valmorca.com.ve'],
    'megalabs': ['ytorres@megalabs.com.ve','Vcampos@megalabs.com.ve','Gmolinaro@megalabs.com.ve',
                 'CPerez@megalabs.com.ve','Atencionalcliente@megalabs.com.ve','SRosario@megalabs.com.ve'],
    'siegfried': ['drosales@siegfried.com.ve','psojo@siegfried.com.ve','asuarez@siegfried.com.ve'],
}


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
hoy = date.today()
fecha_inicio = st.date_input("Fecha inicio", value=hoy - timedelta(days=7))
fecha_fin = st.date_input("Fecha fin", value=hoy - timedelta(days=1))

fecha_inicio_legible = fecha_inicio.strftime('%d-%m-%Y')
fecha_fin_legible = fecha_fin.strftime('%d-%m-%Y')

# Convertir fechas a string
fecha_inicio_str = fecha_inicio.strftime('%Y-%m-%d')
fecha_fin_str = fecha_fin.strftime('%Y-%m-%d')
# debug = st.checkbox("🛠️ Activar Debug")
usd_report = st.checkbox("💵 USD")
if st.button("Generar Reportes"):
    archivos_generados = []
    

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
            fila = {
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
                # 'Subtotal': linea['price_subtotal'], 
            }
            if usd_report:
                fila['Tasa Día'] = factura.get('os_currency_rate', None)
            
            lineas_filtradas.append(fila)
  
    
    df = pd.DataFrame(lineas_filtradas)
    st.session_state.df = df
    if df.empty:
        st.warning("❌ No hay datos después del filtro por proveedores.")
        st.stop()

    #Conversiones
    df['Código de Barras'] = df['Código de Barras'].astype(str)
    panel_df['Código de Barras'] = panel_df['Código de Barras'].astype(str)

    df = df.drop(columns=['Descuento'], errors='ignore')  # Eliminar si ya existe
    df = df.merge(panel_df[['Código de Barras', 'Descuento Panel']], on='Código de Barras', how='left')
    df.rename(columns={'Descuento Panel': 'Descuento'}, inplace=True)
   
    def convertir_a_float(valor):
        if isinstance(valor, str):
            return float(valor.replace(",", ""))
        elif pd.notnull(valor):
            return float(valor)
        return 0.0

    for proveedor in df['Laboratorio'].unique():
        nombre_normalizado = mapa_laboratorios.get(proveedor, proveedor)
        df_proveedor = df[df['Laboratorio'] == proveedor].copy()
        
        df_proveedor = df[df['Laboratorio'] == proveedor].copy()
        # Calcular y agregar columna Subtotal Monto NC al DataFrame
        df_proveedor['Subtotal'] = None

        df_proveedor['Monto NC'] = None
        
        if usd_report:
            df_proveedor['Monto USD NC'] = None

        orden_columnas = [
            'Fecha Factura',
            'Cliente',
            'Nro. Factura',
            'Código de Barras',
            'Producto',
            'Laboratorio',
            'Código Laboratorio',
            'Cantidad',
            'Precio Unitario',
            'Descuento',
            'Subtotal',
            'Monto NC',
            'Tasa Día',
            'Monto USD NC'
        ]

        # if debug:
        #     st.write("🧾 Reviso Monto NC")
        #     st.dataframe(df.head(1))

        #     st.write("🧬 Columnas")
        #     st.write(df.columns.tolist())

        # import openpyxl
        #     try:
        #         wb = openpyxl.load_workbook(archivo_excel)
        #         ws = wb.active
        #         columnas_excel = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
        #         st.write("🔍 Columnas reales en el Excel generado:", columnas_excel)
        #     except Exception as e:
        #         st.error(f"❌ Error al leer columnas del Excel: {e}")



        df_proveedor = df_proveedor[[col for col in orden_columnas if col in df_proveedor.columns]]
        archivo_excel = f'{nombre_normalizado} - Ventas a Farmago del {fecha_inicio_legible} al {fecha_fin_legible}.xlsx'
        with pd.ExcelWriter(archivo_excel, engine='xlsxwriter') as writer:
            df_proveedor.to_excel(writer, sheet_name='Facturas Filtradas', index=False)
            workbook = writer.book
            worksheet = writer.sheets['Facturas Filtradas']

            # Formatos

            percent_format = workbook.add_format({
                'num_format': '0%',  # o simplemente '0%' si no quieres decimales
                'align': 'right'
            })
            money_format = workbook.add_format({
                'num_format': '_-* "Bs.S "* #,##0.00_-;-_* "Bs.S "* -#,##0.00_-;_-* "Bs.S "* "-"??_-;_-@_-',
                'align': 'right'
            })
            dollar_format = workbook.add_format({
                'num_format': '_-* "$"* #,##0.00_-;-_* "$"* -#,##0.00_-;_-* "$"* "-"??_-;_-@_-',
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
            total_usd_format = workbook.add_format({
                'bold': True,
                'num_format': '_-* "$"* #,##0.00_-;-_* "$"* -#,##0.00_-;_-* "$"* "-"??_-;_-@_-',
                'align': 'right'
            })

            # Posiciones
            col_j = df_proveedor.columns.get_loc('Descuento')
            col_i = df_proveedor.columns.get_loc('Precio Unitario')
            col_k = df_proveedor.columns.get_loc('Subtotal')
            col_l = df_proveedor.columns.get_loc('Monto NC')
            if 'Monto USD NC' in df_proveedor.columns:
                col_n = df_proveedor.columns.get_loc('Monto USD NC')
            # col_l = len(df_proveedor.columns)  # Nueva columna "Monto NC"

            # worksheet.write(0, col_l, 'Monto NC', header_format)

            # Fórmulas y formatos
            for row in range(1, len(df_proveedor) + 1):
                worksheet.write_formula(row, col_l, f'=J{row + 1} * K{row + 1}', money_format)

                valor_i = df_proveedor.iloc[row - 1, col_i]
                if pd.notnull(valor_i):
                    worksheet.write_number(row, col_i, convertir_a_float(valor_i), money_format)
                
                worksheet.write_formula(row, col_k, f'=H{row + 1} * I{row + 1}', money_format)
                if 'Monto USD NC' in df_proveedor.columns:
                    worksheet.write_formula(row, col_n, f'=L{row + 1} / M{row + 1}', dollar_format)


                if usd_report and 'Tasa Día' in df_proveedor.columns:
                    col_tasa = df_proveedor.columns.get_loc('Tasa Día')
                    for row in range(1, len(df_proveedor) + 1):
                        tasa = df_proveedor.iloc[row - 1, col_tasa]
                        if pd.notnull(tasa):
                            worksheet.write_number(row, col_tasa, convertir_a_float(tasa), money_format)


            # Encabezados
            for col_idx, col_name in enumerate(df_proveedor.columns):
                worksheet.write(0, col_idx, col_name, header_format)

            

            # Totales
            total_row = len(df_proveedor) + 1
            worksheet.write(total_row, 6, 'Total Unidades', total_format)
            worksheet.write_formula(total_row, 7, f'=SUM(H2:H{total_row})', total_format)          
            if usd_report:
                worksheet.write(total_row, 12, 'Total NC USD', total_format)
                worksheet.write_formula(total_row, 13, f'=SUM(N2:N{total_row})', total_usd_format)
            else:
                worksheet.write(total_row, 10, 'Total NC', total_format)
                worksheet.write_formula(total_row, 11, f'=SUM(L2:L{total_row})', total_money_format)
            # Ajuste de ancho de columnas
            columnas = list(df_proveedor.columns)
            for idx, col in enumerate(columnas):
                if col in ['Precio Unitario', 'Subtotal', 'Monto NC', 'Tasa Día', 'Monto USD NC']:
                    if col == 'Monto NC':
                        try:
                            total_monto_nc = (df_proveedor['Precio Unitario'].apply(convertir_a_float) * df_proveedor['Cantidad'].apply(convertir_a_float) * df_proveedor['Descuento'].apply(convertir_a_float)).sum()
                            max_len_str = f"Bs.S {total_monto_nc:,.2f}"
                        except:
                            max_len_str = "Bs.S 0.00"
                    elif col == 'Subtotal':
                        try:
                            max_subtotal = max(df_proveedor['Cantidad'].apply(convertir_a_float) * df_proveedor['Precio Unitario'].apply(convertir_a_float))
                            # st.write(max_subtotal)
                            max_len_str = f"Bs.S {max_subtotal:,.2f}"
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
            worksheet.set_column(9, 9, 12, percent_format)

        # =============================================
        # NUEVO CÓDIGO PARA ENVIAR A GOOGLE SHEETS
        # =============================================
        # Calcular montos para Google Sheets
        df_excel = pd.read_excel(archivo_excel, sheet_name='Facturas Filtradas')

        if usd_report:
            # Extraer el TOTAL de la última fila (columna 'Monto USD NC')
            monto_usd = df_excel.iloc[-1]['Monto USD NC']  # Última fila, columna N
            monto_bs = None
            
        else:
            monto_bs = (df_proveedor['Precio Unitario'].apply(convertir_a_float) * 
                       df_proveedor['Cantidad'].apply(convertir_a_float) * 
                       df_proveedor['Descuento'].apply(convertir_a_float)).sum()
            monto_usd = None
        
        # Preparar datos para enviar
        data = {
            "laboratorio": nombre_normalizado,
            "mes": fecha_fin.strftime("%B %Y").capitalize(),  # "Julio 2025"
            "concepto": f"Ventas Farmago del {fecha_inicio_legible} al {fecha_fin_legible}",
            "monto_bs": monto_bs,
            "monto_usd": monto_usd
        }
        
        # Enviar datos a Google Sheets
        if enviar_a_google_sheets(data):
            st.success(f"Datos de {nombre_normalizado} enviados a Google Sheets")
        else:
            st.warning(f"No se pudo enviar datos de {nombre_normalizado} a Google Sheets")
        # =============================================
        # FIN DEL NUEVO CÓDIGO
        # =============================================
            
        st.session_state.archivos_generados.append(archivo_excel)
 
    
    
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

# st.write("📊 Debug - ¿Existe df en session_state?:", "df" in st.session_state)
# if "df" in st.session_state:
#     st.write("🔍 Columnas de df:", st.session_state.df.columns.tolist())
#     st.write("🔬 Valores únicos de Laboratorio:", st.session_state.df['Laboratorio'].unique().tolist())


st.subheader("✉️ Correos sugeridos por proveedor")

if "mostrar_correos" not in st.session_state:
    st.session_state.mostrar_correos = False

# Botón que activa la visualización de correos
if st.button("Armar correos"):
    st.session_state.mostrar_correos = True

# Mostrar correos solo si ya se generó el DataFrame
if "df" in st.session_state and not st.session_state.df.empty:
        df = st.session_state.df
     
        # 🔁 Generar correos usando nombres normalizados
        for lab_original in df['Laboratorio'].unique():
            nombre_normalizado = mapa_laboratorios.get(lab_original, lab_original)
            nombre_clave = nombre_normalizado.strip().lower()

            if nombre_clave in emails_por_laboratorio:
                emails = emails_por_laboratorio[nombre_clave]
                titulo = f"{nombre_normalizado} - Ventas a Farmago del {fecha_inicio_legible} al {fecha_fin_legible}"
                cuerpo = f"""Buen día, estimados. Espero se encuentren bien.

Envío el reporte de ventas a Farmago durante el período indicado.

¡Saludos!"""

        # Codificar texto para URL
                subject = urllib.parse.quote(titulo)
                body = urllib.parse.quote(cuerpo)
                to = ",".join(emails)
                cc = ",".join(correos_cc_global)  # NUEVO

                mailto_link = f"mailto:{to}?cc={cc}&subject={subject}&body={body}"

                st.markdown(f"📨 [Redactar correo {nombre_normalizado}]({mailto_link})", unsafe_allow_html=True)
            #     st.write("📨 Enlace generado:", mailto_link)

            #     st.markdown(f"### 📧 Correo para {nombre_normalizado}")
            #     st.write("**Destinatarios:**", ", ".join(emails))
            #     st.write("**Asunto:**", titulo)
            #     st.text_area("Cuerpo del correo", cuerpo, height=150, key=f"cuerpo_{nombre_clave}")
            else:
                st.warning(f"⚠️ No se encontró correo para el laboratorio: {lab_original}")
elif st.session_state.mostrar_correos:  # <- Este chequeo extra evita mensaje prematuro
    st.info("⚠️ Primero debes generar los reportes para armar los correos.")


if st.session_state.archivos_generados:
    for i, fn in enumerate(st.session_state.archivos_generados):
        with open(fn, 'rb') as f:
            st.download_button(
                f"📥 Descargar {fn}",
                f,
                fn,
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                key=f"dl_{i}"
            )
    if st.button("🧹 Limpiar archivos generados"):
        st.session_state.archivos_generados = []



# Hiperactualizado 06/07/2025

            
            
