import streamlit as st
import pandas as pd
import io
from datetime import date
import urllib.parse
from odoo_utils import OdooClient
import numpy as np
import requests

# --- FUNCIONES DE SOPORTE ---

def enviar_a_sheets(df_display, fecha_inicio, fecha_fin, apps_script_url, config_costos=None):
    config_costos = config_costos or {}
    CADENAS_PRECIO_FULL = ['farmago', 'farmatención']
    concepto = f"Sell-Out del {fecha_inicio.strftime('%d/%m/%Y')} al {fecha_fin.strftime('%d/%m/%Y')} (en Panel)"
    
    meses_es = {
        'January': 'Enero', 'February': 'Febrero', 'March': 'Marzo',
        'April': 'Abril', 'May': 'Mayo', 'June': 'Junio',
        'July': 'Julio', 'August': 'Agosto', 'September': 'Septiembre',
        'October': 'Octubre', 'November': 'Noviembre', 'December': 'Diciembre'
    }
    mes_en = fecha_inicio.strftime('%B')
    mes = f"{meses_es.get(mes_en, mes_en)} {fecha_inicio.strftime('%Y')}"

    df = df_display.copy()
    for col in ['quantity', 'price_unit', 'costo_laboratorio', 'descuento_valor']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    tipo_activo = st.session_state.get('tipo_reporte_activo', '')

    resultados = []
    errores = []

    labs = df['laboratory_name'].unique()
    for lab in labs:
        df_lab = df[df['laboratory_name'] == lab].copy()
        es_a_costo = config_costos.get(lab, False)

        # Misma lógica que motor_split_laboratorios
        if es_a_costo:
            df_lab['valor_calculado'] = df_lab['costo_laboratorio']
        else:
            if tipo_activo == 'SELL-OUT':
                def calcular_precio_fila(row):
                    cadena = str(row['cadena']).lower().strip()
                    es_cadena_full = any(c in cadena for c in CADENAS_PRECIO_FULL)
                    if es_cadena_full:
                        return row['price_unit']
                    descuento = row['descuento_valor']
                    if descuento >= 1 or descuento < 0:
                        return row['price_unit']
                    return row['price_unit'] / (1 - descuento)
                df_lab['valor_calculado'] = df_lab.apply(calcular_precio_fila, axis=1)
            else:
                df_lab['valor_calculado'] = df_lab['price_unit']

        df_lab['subtotal_bruto'] = df_lab['quantity'] * df_lab['valor_calculado']
        df_lab['total_descuento'] = df_lab['subtotal_bruto'] * df_lab['descuento_valor']

        moneda = str(df_lab['currency_id'].iloc[0]).lower().strip()
        total = df_lab['total_descuento'].sum()

        es_usd = any(m in moneda for m in ['usd', 'dolar', '$'])
        monto_bs  = 0 if es_usd else round(total, 2)
        monto_usd = round(total, 2) if es_usd else 0

        payload = {
            "action": "append_data",
            "data": {
                "laboratorio": lab,
                "mes": mes,
                "concepto": concepto,
                "monto_bs": monto_bs,
                "monto_usd": monto_usd
            }
        }

        try:
            response = requests.post(apps_script_url, json=payload, timeout=15)
            result = response.json()
            if result.get("success"):
                resultados.append(lab)
            else:
                errores.append(f"{lab}: {result.get('error', 'error desconocido')}")
        except Exception as ex:
            errores.append(f"{lab}: {str(ex)}")

    return resultados, errores


def estandarizar_barcodes(serie):
    return serie.astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

def limpiar_texto_odoo(val):
    return val[1] if isinstance(val, (list, tuple)) else val

def obtener_ofertas_sheets(url):
    try:
        df = pd.read_csv(url)
        # Mapeo según especificaciones: Col A(0), Col E(4), Col F(5), Col H(7), Col I(8)
        df = df.rename(columns={
            df.columns[0]: 'barcode_key',
            df.columns[4]: 'descuento_valor',
            df.columns[5]: 'nc_check',
            df.columns[7]: 'oferta_inicio',
            df.columns[8]: 'oferta_fin'
        })
        
        # APLICACIÓN FILTRO NC: Solo valores "NC" (case-insensitive)
        df['nc_check'] = df['nc_check'].astype(str).str.upper().str.strip()
        df = df[df['nc_check'] == 'NC'].copy()
        
        df['barcode_key'] = estandarizar_barcodes(df['barcode_key'])
        df['oferta_inicio'] = pd.to_datetime(df['oferta_inicio'], errors='coerce').dt.date
        df['oferta_fin'] = pd.to_datetime(df['oferta_fin'], errors='coerce').dt.date
        return df
    except Exception as e:
        st.error(f"Error en Google Sheets: {e}")
        return pd.DataFrame()

def procesar_excel_cadenas(file):
    try:
        df = pd.read_excel(file)
        df.rename(columns={df.columns[0]: 'barcode_key'}, inplace=True)
        df['barcode_key'] = estandarizar_barcodes(df['barcode_key'])
        return df
    except Exception as e:
        st.error(f"Error al procesar el Excel: {e}")
        return pd.DataFrame()

# MOTOR DE EXCEL DINÁMICO CON FÓRMULAS
def motor_split_laboratorios(df_final, config_costos=None):
    if df_final.empty:
        return {}
    
    config_costos = config_costos or {}
    laboratorios = df_final['laboratory_name'].unique()
    diccionario_excels = {}
    
    for lab in laboratorios:
        # Copia para procesar cálculos sin alterar el DataFrame de la pantalla
        df_lab = df_final[df_final['laboratory_name'] == lab].copy()
        es_a_costo = config_costos.get(lab, False)
        
        # LIMPIEZA DE SEGURIDAD NUMÉRICA
        for col in ['quantity', 'price_unit', 'costo_laboratorio', 'descuento_valor']:
            if col in df_lab.columns:
                df_lab[col] = df_lab[col].apply(lambda x: x[0] if isinstance(x, (list, tuple)) else x)
                df_lab[col] = pd.to_numeric(df_lab[col], errors='coerce').fillna(0)

        # Columna I: Determinar valor base según costo, cadena y tipo de reporte
        CADENAS_PRECIO_FULL = ['farmago', 'farmatención']

        if es_a_costo:
            # Costo siempre tiene prioridad
            df_lab['valor_calculado'] = df_lab['costo_laboratorio']
        else:
            # Solo SELL-OUT aplica la lógica de cadena
            tipo_activo = st.session_state.get('tipo_reporte_activo', '')
            
            if tipo_activo == 'SELL-OUT':
                def calcular_precio_fila(row):
                    cadena = str(row.get('cadena', '')).lower().strip()
                    es_cadena_full = any(c in cadena for c in CADENAS_PRECIO_FULL)
                    if es_cadena_full:
                        return row['price_unit']
                    else:
                        descuento = row['descuento_valor']
                        if descuento >= 1 or descuento < 0:
                            return row['price_unit']  # descuento inválido, no recalcula
                        return row['price_unit'] / (1 - descuento)
                
                df_lab['valor_calculado'] = df_lab.apply(calcular_precio_fila, axis=1)
            else:
                df_lab['valor_calculado'] = df_lab['price_unit']

        col_valor_base = 'valor_calculado'
        
        
        # Construcción de la estructura solicitada (Columnas A-J)
        reporte = pd.DataFrame({
            'invoice_date': df_lab['invoice_date'],
            'partner_id': df_lab['partner_id'],
            'invoice_number_next': df_lab['invoice_number_next'],
            'barcode': df_lab['barcode'],
            'name': df_lab['name'],
            'laboratory_name': df_lab['laboratory_name'],
            'supplier_code': df_lab['supplier_code'],
            'quantity': df_lab['quantity'],
            'valor_unitario': df_lab[col_valor_base],
            'descuento': df_lab['descuento_valor'],
            'Moneda': df_lab['currency_id']
        })
        
        # Fórmulas K y L
        reporte['subtotal_bruto'] = reporte['quantity'] * reporte['valor_unitario']
        reporte['total_descuento'] = reporte['subtotal_bruto'] * reporte['descuento']

        
        # Generación del binario para descarga
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            reporte_export = reporte.drop(columns=['Moneda'])
            reporte_export.to_excel(writer, index=False, sheet_name='Reporte', startrow=1, header=False)
            workbook  = writer.book
            worksheet = writer.sheets['Reporte']

            # --- FORMATO DE ENCABEZADOS ---
            header_format = workbook.add_format({'bold': True, 'border': 0})

            percent_format = workbook.add_format({'num_format': '0%'})

            dollar_format = workbook.add_format({"num_format": "$#,##0.00"})
            bs_format = workbook.add_format({"num_format": '"Bs." #,##0.00'})

            bold_format = workbook.add_format({'bold': True})

            worksheet.set_column(9, 9, None, percent_format)
            
            # Encabezados personalizados
            encabezados = [
                'Fecha Factura', 'Cliente', 'Nro. Factura', 'Código de Barras', 'Descripción', 
                'Laboratorio', 'Código Laboratorio', 'Cantidad', 'Precio', 'Descuento %', 
                'Total', 'Monto NC'
            ]
            
            for col_num, value in enumerate(encabezados):
                worksheet.write(0, col_num, value, header_format)

            # --- FORMATO DE COLUMNAS CON FÓRMULAS ---
            for row_num in range(len(reporte)):
                # Col I = Precio (índice 8), Col H = Cantidad (índice 7)
                worksheet.write_formula(row_num + 1, 10, f'=I{row_num + 2}*H{row_num + 2}')
                # Col J = Descuento % (índice 9), Col K = Total (índice 10)
                worksheet.write_formula(row_num + 1, 11, f'=J{row_num + 2}*K{row_num + 2}')

            # --- Totales

            last_row = len(reporte) + 1
            worksheet.write(last_row, 10, "Total NC", bold_format)
        
            worksheet.write_formula(
                last_row,
                11,
                f"=SUM(L2:L{last_row})",
                bold_format
            )
              
            # --- Formato moneda
            for row_num, moneda in enumerate(reporte["Moneda"], start=1):
                fmt = dollar_format if str(moneda).lower() in ["usd", "dolares", "$"] else bs_format
                #Aplica formato condicional a cada celda sin tocar su contenido
                worksheet.conditional_format(row_num, 8, row_num, 8, {'type': 'no_errors', 'format': fmt})  # Col I
                worksheet.conditional_format(row_num, 10, row_num, 10, {'type': 'no_errors', 'format': fmt}) # Col K
                worksheet.conditional_format(row_num, 11, row_num, 11, {'type': 'no_errors', 'format': fmt}) # Col L

            worksheet.conditional_format(last_row, 11, last_row, 11, {'type': 'no_errors', 'format': fmt})

            # -- Ajuste anchos
            for i, col in enumerate(reporte.columns):
                if i < 9:
                    column_data = reporte[col].astype(str).fillna("")
                    max_len = max(column_data.map(len).max(), len(col))
                    worksheet.set_column(i, i, max_len + 2)

            # --- ajuste anchos 
            for col in [10, 11]:
                column_data = reporte['subtotal_bruto' if col == 10 else 'total_descuento']

                column_data = column_data.fillna(0).astype(float)

                max_len = max(column_data.astype(str).map(len).max(), 12)

                worksheet.set_column(col, col, max_len + 6)



        output.seek(0)
        diccionario_excels[lab] = output.getvalue()
        
    return diccionario_excels

# --- FUNCIÓN PRINCIPAL ---
def render_reporte(fecha_inicio, fecha_fin):
    st.header("🎯 Panel de Inteligencia Comercial")
    
    

    def limpiar(val): 
        return val[1] if isinstance(val, (list, tuple)) else val
    
    def limpiar_barcode(val): 
        if not val: return ""
        return str(val).split('.')[0].strip()

    # --- INICIALIZACIÓN DE SESIÓN ---
    if 'df_resultado' not in st.session_state:
        st.session_state.df_resultado = None
    if 'archivos_binarios' not in st.session_state:
        st.session_state.archivos_binarios = {}
    if 'tipo_reporte_activo' not in st.session_state:
        st.session_state.tipo_reporte_activo = ""
    if 'config_costos' not in st.session_state:
        st.session_state.config_costos = {}

    if 'config_costos_aplicada' not in st.session_state:
        st.session_state.config_costos_aplicada = {}

    if st.session_state.get('_reset_checkboxes', False):
        keys_a_borrar = [k for k in st.session_state if k.startswith("chk_")]
        for k in keys_a_borrar:
            del st.session_state[k]
        st.session_state._reset_checkboxes = False


    # 1. SELECTOR DE MODO
    
    tipo_reporte = st.radio(
        "Seleccione el tipo de análisis:",
        ["Extracción General", "SELL-OUT", "Farmago", "Farmatención"],
        horizontal=True,
        key="selector_principal"
    )

    st.divider()

    # 2. CONFIGURACIÓN DE REFERENCIAS
    df_referencia = pd.DataFrame()
    if tipo_reporte == "SELL-OUT":
        url = st.text_input("Link de Google Sheets", "https://docs.google.com/spreadsheets/d/1c4Eil9IoOhUTNr3_jrZn5HI5GNZq9NTkgPH0CbjwYMA/export?format=csv")
        if url: 
            df_referencia = obtener_ofertas_sheets(url)
            if not df_referencia.empty:
                st.info("✅Se detectaron ofertas válidas.")
                
    elif tipo_reporte in ["Farmago", "Farmatención"]:
        uploaded_file = st.file_uploader(f"Subir Excel para {tipo_reporte}", type=["xlsx", "xls"])
        if uploaded_file: df_referencia = procesar_excel_cadenas(uploaded_file)

    # 3. EJECUCIÓN
    if st.button(f"🚀 Generar Reporte", type="primary"):
        try:
                        
            for k in list(st.session_state.keys()):
                if k.startswith("chk_"):
                    st.session_state[k] = False
            
            st.session_state.config_costos = {}
            st.session_state.config_costos_aplicada = {} 

            config = st.secrets["odoo_bd1"]
            client = OdooClient(config["url"], config["db"], config["username"], config["password"])

            domain = [
                ('date', '>=', str(fecha_inicio)), ('date', '<=', str(fecha_fin)),
                ('move_type', '=', 'out_invoice'), ('parent_state', '=', 'posted'),
                ('move_name', 'not ilike', 'ND%'), ('product_id', '!=', False), ('quantity', '>', 0)
            ]

            if tipo_reporte == "SELL-OUT" and not df_referencia.empty:
                barcodes = df_referencia['barcode_key'].unique().tolist()
                domain.append(('product_id.barcode', 'in', barcodes))
            elif tipo_reporte in ["Farmago", "Farmatención"]:
                domain.append(('partner_id.name', 'ilike', tipo_reporte))

            with st.spinner("Consultando Odoo..."):
                fields_lineas = ['move_id', 'product_id', 'name', 'quantity', 'price_unit']
                data_lineas = client.search_read('account.move.line', domain, fields_lineas)
                if not data_lineas:
                    st.warning("No hay datos para esta selección.")
                    return
                df_lineas = pd.DataFrame(data_lineas)

                move_ids = list(set([x[0] for x in df_lineas['move_id'] if isinstance(x, list)]))
                product_ids = list(set([x[0] for x in df_lineas['product_id'] if isinstance(x, list)]))

                data_moves = client.search_read('account.move', [('id', 'in', move_ids)], 
                                               ['invoice_date', 'partner_id', 'invoice_number_next', 'currency_id'])
                df_moves = pd.DataFrame(data_moves).rename(columns={'id': 'move_id_int'})

                data_prods = client.search_read('product.product', [('id', 'in', product_ids)], 
                                               ['laboratory_name', 'supplier_code', 'barcode'])
                df_prods = pd.DataFrame(data_prods).rename(columns={'id': 'product_id_int'})

                data_costs = client.search_read('product.supplierinfo', [('product_tmpl_id', 'in', product_ids)], ['product_tmpl_id', 'price'])
                df_costs = pd.DataFrame(data_costs)
                if not df_costs.empty:
                    df_costs['product_id_int'] = df_costs['product_tmpl_id'].apply(lambda x: x[0] if isinstance(x, (list, tuple)) else x)
                    df_costs = df_costs.rename(columns={'price': 'costo_proveedor'}).drop_duplicates('product_id_int')
                              
                else:
                    df_costs = pd.DataFrame(columns=['product_id_int', 'costo_proveedor'])

                partner_ids_raw = list(set([m['partner_id'][0] for m in data_moves if isinstance(m.get('partner_id'), (list, tuple))]))
                data_partners = client.search_read('res.partner', [('id', 'in', partner_ids_raw)], ['id', 'cadena'])
                df_partners = pd.DataFrame(data_partners).rename(columns={'id': 'partner_id_int', 'cadena': 'cadena_val'})

                df_lineas['move_id_int'] = df_lineas['move_id'].apply(lambda x: x[0] if isinstance(x, list) else x)
                df_lineas['product_id_int'] = df_lineas['product_id'].apply(lambda x: x[0] if isinstance(x, list) else x)


                df_final = df_lineas.merge(df_moves, on='move_id_int', how='left')
                df_final = df_final.merge(df_prods, on='product_id_int', how='left')
                df_final = df_final.merge(df_costs[['product_id_int', 'costo_proveedor']], on='product_id_int', how='left')
                df_final['partner_id_int'] = df_final['partner_id'].apply(lambda x: x[0] if isinstance(x, (list, tuple)) else None)
                df_final = df_final.merge(df_partners, on='partner_id_int', how='left')

                if tipo_reporte == "SELL-OUT" and not df_referencia.empty:
                    df_final['barcode_key_tmp'] = estandarizar_barcodes(df_final['barcode'])
                    
                                        
                    # --- MERGE ---
                    df_final = df_final.merge(
                        df_referencia,
                        left_on='barcode_key_tmp',
                        right_on='barcode_key',
                        how='inner'
                    )
                    
                    # --- CONVERTIR DESCUENTO TIPO "10%" A 0.10 ---
                    df_final['descuento_valor'] = df_final['descuento_valor'].astype(str)\
                                                .str.replace('%','')\
                                                .astype(float) / 100
                    
                    # --- FILTRO POR FECHAS ---
                    df_final['invoice_date_obj'] = pd.to_datetime(df_final['invoice_date']).dt.date
                    df_final = df_final[
                        (df_final['invoice_date_obj'] >= df_final['oferta_inicio']) &
                        (df_final['invoice_date_obj'] <= df_final['oferta_fin'])
                    ]

                res = pd.DataFrame({
                    'invoice_date': pd.to_datetime(df_final['invoice_date']).dt.strftime('%d/%m/%Y'),
                    'partner_id': df_final['partner_id'].apply(limpiar),
                    'cadena': df_final['cadena_val'].apply(lambda x: limpiar(x) if x else ""),
                    'invoice_number_next': df_final['invoice_number_next'],
                    'barcode': df_final['barcode'].apply(limpiar_barcode),
                    'name': df_final['name'],
                    'laboratory_name': df_final['laboratory_name'].apply(limpiar),
                    'supplier_code': df_final['supplier_code'],
                    'quantity': df_final['quantity'],
                    'price_unit': df_final['price_unit'],
                    'costo_laboratorio': df_final['costo_proveedor'].fillna(0),
                    'descuento_valor': df_final['descuento_valor'] if 'descuento_valor' in df_final else 0,
                    'currency_id': df_final['currency_id'].apply(limpiar)
                })
                
                               
                st.session_state.df_resultado = res
                st.session_state.tipo_reporte_activo = tipo_reporte
                # Los binarios se generan respetando el diccionario de costos actual
                st.session_state.archivos_binarios = motor_split_laboratorios(res, st.session_state.config_costos)
                st.rerun()

        except Exception as e:
            st.error(f"Error crítico: {e}")
            

    # --- 4. RENDERIZADO Y CONFIGURACIÓN DINÁMICA (SOLO SIDEBAR) ---
    if st.session_state.df_resultado is not None:
        df_display = st.session_state.df_resultado
        tipo_activo = st.session_state.tipo_reporte_activo 

        
        
        # SIDEBAR: Único lugar de configuración para laboratorios reales
        labs_encontrados = sorted(df_display['laboratory_name'].unique())
        
        with st.sidebar:
            st.header("⚙️ Configuración de Reporte")

            st.info("Seleccione los laboratorios que desea exportar a **COSTO**.")
            
            for lab in labs_encontrados:
                st.session_state.config_costos[lab] = st.checkbox(
                    f"{lab}",
                    key=f"chk_{lab}"
                )
            
            st.divider()
            if st.button("🔄 Aplicar y Regenerar Excels", use_container_width=True, type="primary"):
                st.session_state.archivos_binarios = motor_split_laboratorios(df_display, st.session_state.config_costos)
                st.session_state.config_costos_aplicada = st.session_state.config_costos.copy()
                st.toast("Archivos de Excel actualizados con éxito")
                st.rerun()
                

        # VISTA PRINCIPAL (Limpia de configuradores innecesarios)
        st.success(f"✅ Extracción finalizada ({tipo_activo}): {len(df_display)} registros.")
        
        col1, col2 = st.columns([1, 4])
        with col1:
            if st.button("🗑️ Limpiar Todo"):
                st.session_state.df_resultado = None
                st.session_state.archivos_binarios = {}
                st.rerun()

        st.dataframe(df_display, use_container_width=True)

        
        st.divider()
        st.subheader("📤 Enviar resumen a Google Sheets")

        

        apps_url = st.text_input(
            "URL del Apps Script",
            value=st.secrets["appscript"]["url"],
            key="input_apps_script_url"
        )

        if st.button("📨 Enviar resumen NC a Sheets", type="primary"):
            if not apps_url:
                st.error("Ingresa la URL del Apps Script.")
            else:
                with st.spinner("Enviando datos..."):
                    ok, errores = enviar_a_sheets(
                        df_display, fecha_inicio, fecha_fin, apps_url, config_costos=st.session_state.config_costos
                    )
                if ok:
                    st.success(f"✅ Enviados: {', '.join(ok)}")
                if errores:
                    for e in errores:
                        st.error(f"❌ {e}")


        # SECCIÓN DE DESCARGAS
        if st.session_state.archivos_binarios:
            st.write("### 📥 Descargar por Laboratorio")
            items = list(st.session_state.archivos_binarios.items())
            for i in range(0, len(items), 3):
                cols = st.columns(3)
                for j in range(3):
                    if i + j < len(items):
                        lab, excel_data = items[i + j]
                        es_costo = st.session_state.config_costos_aplicada.get(lab, False)
                        etiqueta = " (Costo)" if es_costo else ""
                        
                        safe_lab = (
                            lab.replace(" ", "_")
                            .replace("/", "")
                            .replace("\\", "")
                            .replace(":", "")
                            .replace("á", "a")
                            .replace("é", "e")
                            .replace("í", "i")
                            .replace("ó", "o")
                            .replace("ú", "u")
                            .replace("Á", "A")
                            .replace("É", "E")
                            .replace("Í", "I")
                            .replace("Ó", "O")
                            .replace("Ú", "U")
                            .replace("ü", "u")
                            .replace("Ü", "U")
                            .replace("ñ", "n")
                            .replace("Ñ", "N")
                        )

                        with cols[j]:
                            btn_key = f"dl_{lab}_{tipo_activo}_{i+j}".replace(" ", "_")
                            st.download_button(
                                label=f"📦 {lab}{etiqueta}",
                                data=excel_data,
                                file_name=f"{safe_lab}_{tipo_activo}_del_{fecha_inicio.strftime('%d-%m-%Y')}_al_{fecha_fin.strftime('%d-%m-%Y')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=btn_key
                            )
