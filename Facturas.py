# Invoice system developed for El Shaddai - Barcelona
# Author:  Dalton Hardt
# Created:  22-Aug-2024
# Last update:  4-Sep-2024

import googleapiclient
import streamlit as st
from streamlit_option_menu import option_menu
from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.http import MediaIoBaseUpload
import pandas as pd
from datetime import datetime, timedelta
import locale
import base64
import io
import json


# Function to READ/GET values from spreadsheet
def leitura_worksheet(worksheet):
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=worksheet).execute()
    values = result.get("values", [])
    df = pd.DataFrame(values)  # transform all values in DataFrame
    df.columns = df.iloc[0]  # set column names equal to values in row index position 0
    df = df[1:]  # remove first row from DataFrame
    return df


# Function to READ/GET client values from spreadsheet
def leitura_registro_cliente(client_id):
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range='clientes').execute()
    values = result.get("values", [])
    return values[client_id]  # return the record values in a list


# Function to READ/GET invoice values from spreadsheet
def leitura_registro_factura(invoice_id):
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range='facturas').execute()
    values = result.get("values", [])
    return values[invoice_id]  # return the record values in a list


# Function to FORMAT CURRENCY to European standard (Spain)
def format_currency(amount, currency_symbol="€", locale_name="es_ES.UTF-8", decimal_places=2):
    # Set the locale for formatting
    locale.setlocale(locale.LC_ALL, locale_name)
    # Format the amount as currency with the specified decimal places
    formatted_amount = locale.format_string(f"%.{decimal_places}f", amount, grouping=True)
    # Replace the default currency symbol with the desired symbol
    formatted_amount = formatted_amount.replace(locale.localeconv()['currency_symbol'], currency_symbol)
    return f'{currency_symbol} {formatted_amount}'


# Function to UPDATE STATUS of all the invoices checking DUE DATE x TODAY DATE
def update_status_facturas(df):
    for index_factura in range(len(df)):
        # print(index_factura+2, "df_facturas.iloc[index_factura]['plazo_pago']", df_facturas.iloc[index_factura]['plazo_pago'])
        data_factura = datetime.strptime(df.iloc[index_factura]['plazo_pago'], "%d/%m/%Y")
        status_factura = df.iloc[index_factura]['status']
        nro_factura = df.iloc[index_factura]['nro_factura']
        if status_factura == 'Recibir' and data_factura < TODAY:
            # print(f'{index_factura + 2}  factura {nro_factura} com plazo pago {data_factura} está ATRASADA!!!')
            change_status_factura = ['Atrasado']
            celula = "facturas!Q" + str(index_factura + 2)
            # print('.....atualizando celula...', celula)
            sheet.values().update(spreadsheetId=SPREADSHEET_ID,
                                  range=celula, valueInputOption="USER_ENTERED",
                                  body={"values": [change_status_factura]}).execute()
            # print(f".....Registro: {index_factura+2} - {nro_factura} na celula {celula} com Status atualizado para 'Atrasado'")


# Function to change the STATUS of the invoice
@st.experimental_dialog("A T E N C I Ó N")
def change_invoice_status(index_sequence, status, fecha):
    st.write(f"¿Confirma?")
    if st.button("OK"):
        for i in range(len(index_sequence)):
            celula1 = "facturas!F" + str(index_sequence[i])
            celula2 = "facturas!Q" + str(index_sequence[i])
            sheet.values().update(spreadsheetId=SPREADSHEET_ID,
                                  range=celula1, valueInputOption="USER_ENTERED",
                                  body={"values": [[fecha]]}).execute()
            sheet.values().update(spreadsheetId=SPREADSHEET_ID,
                                  range=celula2, valueInputOption="USER_ENTERED",
                                  body={"values": [[status]]}).execute()
            # print(f"Registro: {i} na celula {celula} foi atualizado")
        st.rerun()


# Function to replace the placeholders in the invoice_template in Google Docs
def substituir_placeholders(document_id, substitutions):
    requests = []
    # Iterate over all "substitutions" items and create an update request for Google Docs
    for placeholder, value_text in substitutions.items():
        requests.append({
            'replaceAllText': {
                'containsText': {
                    'text': '{{' + placeholder + '}}',
                    'matchCase': True
                },
                'replaceText': value_text
            }
        })
    # Executa as solicitações de atualização no documento
    result = doc.batchUpdate(documentId=document_id, body={'requests': requests}).execute()
    return result


# Read all GCP Credentials stored in secrets.toml file
gcp_type = st.secrets.gcp_service_account["type"]
gcp_project_id = st.secrets.gcp_service_account["project_id"]
gcp_private_key_id = st.secrets.gcp_service_account["private_key_id"]
gcp_private_key = st.secrets.gcp_service_account["private_key"].replace('\n', '\\n')
gcp_client_email = st.secrets.gcp_service_account["client_email"]
gcp_client_id = st.secrets.gcp_service_account["client_id"]
gcp_auth_uri = st.secrets.gcp_service_account["auth_uri"]
gcp_token_uri = st.secrets.gcp_service_account["token_uri"]
gcp_auth_provider_x509_cert_url = st.secrets.gcp_service_account["auth_provider_x509_cert_url"]
gcp_client_x509_cert_url = st.secrets.gcp_service_account["client_x509_cert_url"]
gcp_universe_domain = st.secrets.gcp_service_account["universe_domain"]
# Create a dictionary string
account_info_str = f'''
{{
  "type": "{gcp_type}",
  "project_id": "{gcp_project_id}",
  "private_key_id": "{gcp_private_key_id}",
  "private_key": "{gcp_private_key}",
  "client_email": "{gcp_client_email}",
  "client_id": "{gcp_client_id}",
  "auth_uri": "{gcp_auth_uri}",
  "token_uri": "{gcp_token_uri}",
  "auth_provider_x509_cert_url": "{gcp_auth_provider_x509_cert_url}",
  "client_x509_cert_url": "{gcp_client_x509_cert_url}",
  "universe_domain": "{gcp_universe_domain}"
}}
'''
# Convert to a JSON string
account_info = json.loads(account_info_str)

# ------- BEGIN Google Definitions -------
SCOPES = st.secrets.google_definition["SCOPES"]
SPREADSHEET_ID = st.secrets.google_definition["SPREADSHEET_ID"]
INVOICE_TEMPLATE_ID = st.secrets.google_definition["INVOICE_TEMPLATE_ID"]
PDF_FOLDER_ID = st.secrets.google_definition["PDF_FOLDER_ID"]
# for Google SHEETS
creds = service_account.Credentials.from_service_account_info(account_info, scopes=SCOPES)
# SPREADSHEET_ID = "1FD6oaUPwjyKJo1yLe3UWBBdN1o1CgcgOSaEqW8ZSk24"  # The ID of the spreadsheet
service = build("sheets", "v4", credentials=creds)
# Call the Sheets API
sheet = service.spreadsheets()
# for Google DOCS
# creds_docs = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE_DOCS, scopes=SCOPES_DOCS)
service_docs = build("docs", "v1", credentials=creds)
# Call the Sheets API
doc = service_docs.documents()
# for Google DRIVE
# creds_drive = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE_DOCS, scopes=SCOPES_DRIVE)
service_drive = build("drive", "v3", credentials=creds)
# ------ END Google Definitions -------

# --- Starting Streamlit
st.set_page_config(layout="wide")
st.header("Base de Datos Facturas 🧾")
st.sidebar.markdown("# Facturas 🧾")

# set the locale to Spanish (Spain)
locale.setlocale(locale.LC_NUMERIC, 'es_ES.UTF-8')

# get today date
TODAY = datetime.strptime(datetime.now().strftime("%d/%m/%Y"), "%d/%m/%Y")

pd.set_option('display.precision', 2)
# Apply format with two decimal numbers for display only
pd.options.display.float_format = '{:.2f}'.format
# Pandas visualization parameters
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

# Create DataFrame with ALL Client values from spreadsheet
df_clientes = leitura_worksheet('clientes')
# st.dataframe(df_clientes)

# Create DataFrame with ONLY ACTIVE client values from spreadsheet
df_clientes_activos = df_clientes[df_clientes['status_cliente'] == 'Activo'].reset_index(drop=True)
df_clientes_activos.index += 1  # making index start from 1 to stay equal with "df_clientes"
# st.dataframe(df_clientes_activos)

# Create DataFrame with ALL invoice values from spreadsheet
df_facturas_original = leitura_worksheet('facturas')

# Call function to update invoice status ('Atrasado') based on the current date today()
update_status_facturas(df_facturas_original)

# Read DataFrame again with ALL invoice values from spreadsheet after updating the status
df_facturas = leitura_worksheet('facturas')

# Add client name (nombre_cliente) column mapping by client unique code (cod_cliente)
df_facturas['nombre_cliente'] = df_facturas.cod_cliente.map(
    df_clientes.set_index('cod_cliente')['nombre_cliente'].to_dict())

# Adjusting the float numbers to match european standard
df_facturas['cantidad'] = df_facturas['cantidad'].apply(
    lambda x: float(str(x).replace('.', '').replace(',', '.')))
df_facturas['precio_unit'] = df_facturas['precio_unit'].apply(
    lambda x: float(str(x).replace('.', '').replace(',', '.')))
df_facturas['base_imponible'] = df_facturas['base_imponible'].apply(
    lambda x: float(str(x).replace('.', '').replace(',', '.')))
df_facturas['cuota_tributaria'] = df_facturas['cuota_tributaria'].apply(
    lambda x: float(str(x).replace('.', '').replace(',', '.')))
df_facturas['valor_retencion'] = df_facturas['valor_retencion'].apply(
    lambda x: float(str(x).replace('.', '').replace(',', '.')))
df_facturas['total'] = df_facturas['total'].apply(
    lambda x: float(str(x).replace('.', '').replace(',', '.')))

# Group all records with the same invoice number and get the sum
df_total_facturas = df_facturas.groupby(['nro_factura'], as_index=False).agg(
    {'total': 'sum', 'nombre_cliente': 'first', 'descripcion': 'first', 'fecha_emision': 'first', 'plazo_pago': 'first',
     'status': 'first'})
df_total_facturas['total'] = df_total_facturas['total'].apply(lambda x: f'{x:.2f}')
df_total_facturas.index += 1  # make index start at 1
# st.dataframe(df_total_facturas)

# Tabs
TAB_0 = 'Todas Facturas'
TAB_1 = 'Editar Factura'
TAB_2 = 'Nueva Factura'

tab = option_menu(
    menu_title='',
    options=['Todas Facturas', 'Editar Factura', 'Nueva Factura'],
    icons=['list-task', 'bi-pencil-square', 'bi-file-earmark-plus'],
    menu_icon='cast',
    orientation='horizontal',
    default_index=0
)

if tab == TAB_0:  # Show ALL invoices
    st.divider()
    # Show all invoices with due date = Today
    st.subheader('Factura con fecha de vencimiento hoy   🎉 ')
    # print('Hoje: ', datetime.now().strftime("%-d/%m/%Y"))  # format "%-d" para mostrar o dia sem o zero na frente
    df_facturas_hoy = df_total_facturas[
        df_total_facturas['plazo_pago'] == datetime.now().strftime("%-d/%m/%Y")].reset_index(drop=True)
    df_facturas_hoy.index += 1  # making index start from 1 to stay equal with "df_clientes"
    if len(df_facturas_hoy.index) > 0:
        st.dataframe(df_facturas_hoy)
        # st.dataframe(df_facturas_atrasadas.iloc[::-1])  # show dataframe in reverse order (from newest to oldest)
    else:
        st.success(f"¡No hay facturas para recibir pago hoy!", icon='⏳')

    st.divider()
    # Show all invoices unpaid
    st.subheader('Facturas atrasadas   😳 ')
    df_facturas_atrasadas = df_total_facturas[df_total_facturas['status'] == 'Atrasado'].reset_index(drop=True)
    df_facturas_atrasadas.index += 1  # making index start from 1 to stay equal with "df_clientes"
    if len(df_facturas_atrasadas.index) > 0:
        st.dataframe(df_facturas_atrasadas)
        # st.dataframe(df_facturas_atrasadas.iloc[::-1])  # show dataframe in reverse order (from newest to oldest)
    else:
        st.success(f"¡No hay facturas retrasadas!", icon='✅')

    st.divider()
    # Show all invoices to be paid
    st.subheader('Facturas por recibir  ⌛')
    df_facturas_recibir = df_total_facturas[df_total_facturas['status'] == 'Recibir'].reset_index(drop=True)
    df_facturas_recibir.index += 1  # making index start from 1 to stay equal with "df_clientes"
    st.dataframe(df_facturas_recibir)
    # st.dataframe(df_facturas_recibir.iloc[::-1])  # show dataframe in reverse order (from newest to oldest)

    st.divider()
    # Show ALL invoices
    st.subheader('Todas Facturas')
    st.dataframe(df_total_facturas)
    # st.dataframe(df_total_facturas.iloc[::-1])  # show dataframe in reverse order (from newest to oldest)

if tab == TAB_1:  # Change Invoice
    st.divider()
    # Option to change the status to PAID for the invoices Unpaid and to Receive
    st.subheader('Cambiar status de pago de Facturas atrasadas y a recibir')
    df_facturas_nopagas = df_total_facturas[df_total_facturas['status'].isin(['Atrasado', 'Recibir'])].reset_index(
        drop=True)
    df_facturas_nopagas.index += 1
    st.dataframe(df_facturas_nopagas)
    # select Client name from dataframe

    st.divider()
    # Let the user select unpaid invoice and change the status
    col1, col2, col3 = st.columns([0.3, 0.4, 0.3])
    with col1:
        invoice = st.selectbox('Factura Nro.', df_facturas_nopagas['nro_factura'].sort_values(),
                               index=None, placeholder='Seleccione...', key='invoice_key')
    if invoice is not None:
        for index_total in range(len(df_total_facturas)):
            # print('buscando total da fatura...', index_total)
            if df_total_facturas.iloc[index_total]['nro_factura'] == invoice:
                invoice_total = df_total_facturas.iloc[index_total]['total']
                # print('...invoice_total', index_total, ' = ', invoice_total)
                break
        #
        index_invoice_sequence = []  # initialize a list to recieve the line indexes corresponding to the same invoice
        for index_invoice in range(len(df_facturas)):
            # print('index da factura:', index_invoice)
            if df_facturas.iloc[index_invoice]['nro_factura'] == invoice:
                index_invoice += 1
                result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range='facturas').execute()
                values = result.get("values", [])
                registro = leitura_registro_factura(index_invoice)
                # print('registro da FACTURA:\n', registro)
                # for i in range(len(registro)):
                #     st.text(f'{i} - {registro[i]}')
                # st.text(f'index no df_clientes = {str(index_cliente)}')
                invoice_cod_client = registro[1]
                invoice_fecha_emision = registro[3]
                invoice_plazo_pago = registro[4]
                invoice_fecha_pago = registro[5]
                invoice_descripcion = registro[6]
                invoice_status = registro[16]
                index_invoice_sequence.append(index_invoice + 1)
        # print('index_invoice_sequence =', index_invoice_sequence)

        # Display some info from the selected invoice
        with col2:
            st.write()
            st.write(f'Fecha emisión factura: {invoice_fecha_emision}')
            st.write(f'Total factura: {invoice_total}')
            st.write(f'Status actual de la factura: {invoice_status}')

        # Display a button to change the status of the invoice
        with col3:
            new_invoice_status = st.selectbox('Nuevo Status', ['Pagado', 'Cancelado'], index=None,
                                              placeholder='Seleccione...')
            if new_invoice_status == 'Pagado':
                new_invoice_fecha_pago = st.date_input('Fecha de Pago:', format='DD/MM/YYYY')
                # print('new_invoice_fecha_pago =', new_invoice_fecha_pago)
                change_invoice_status(index_invoice_sequence, new_invoice_status, str(new_invoice_fecha_pago))
            elif new_invoice_status == 'Cancelado':
                change_invoice_status(index_invoice_sequence, new_invoice_status, '')

if tab == TAB_2:  # Create NEW Invoice
    with st.container():
        # print('df_facturas:\n', df_facturas)
        last_invoice_row = len(df_facturas)  # get the last written row from the dataframe
        last_invoice = df_facturas.loc[last_invoice_row, 'nro_factura']  # get invoice number from column 'nro_factura'
        # print('last_invoice_row = ', last_invoice_row, 'with number = ', last_invoice)
        invoice_num = int(last_invoice[-3:]) + 1  # add 1 to create the new invoice sequential number
        current_year = datetime.now().strftime('%y')  # get the current year with two-digits
        current_month = datetime.now().strftime('%m')  # get the current year with two-digits
        invoice_nr = current_year + current_month + str(int(last_invoice[-3:]) + 1).zfill(3)  # zfill=3 format YYMM999
        st.divider()
        st.subheader('Nro: ' + invoice_nr)

        col1, col2 = st.columns([0.7, 0.3])
        with col1:
            # select Client name from dataframe
            client = st.selectbox('Cliente *', df_clientes_activos['nombre_cliente'].sort_values(),
                                  index=None, placeholder='Seleccione...', key='client_key')

        if client is not None:
            # st.text(f'Cliente seleccionado: {client}')
            for index_cliente in range(len(df_clientes)):
                # print(index_cliente)
                if df_clientes.iloc[index_cliente]['nombre_cliente'] == client:
                    index_cliente += 1
                    registro = leitura_registro_cliente(index_cliente)
                    # for i in range(len(registro)):
                    #     st.text(f'{i} - {registro[i]}')
                    # st.text(f'index no df_clientes = {str(index_cliente)}')
                    client_cod = registro[0]
                    client_cif = registro[2]
                    client_prov = registro[3]
                    client_city = registro[4]
                    client_address = registro[5]
                    client_postal = registro[6]
                    client_contact = registro[7]
                    client_email = registro[8]
                    client_phone = registro[9]
                    client_obs = registro[10]
                    break

            invoice_date = datetime.now()  # current date
            invoice_due = invoice_date + timedelta(days=30)  # due date in 30 days
            # st.text(f'invoice date = {str(invoice_date.strftime("%d-%m-%Y"))}')
            # st.text(f'invoice due date = {str(invoice_due.strftime("%d-%m-%Y"))}')

            col2, col3, col4, col5, col6 = st.columns(5)
            with col2:
                # form_invoice_status = st.selectbox('Status', ['Recibir', 'Enviada', 'Pagado', 'Atrasado'], 0)
                form_invoice_pedido = st.text_input('Nro. pedido:', placeholder='Introduzca...')
            with col3:
                form_invoice_date = st.date_input('Fecha emissión:', value=invoice_date, format="DD/MM/YYYY")
            with col4:
                form_invoice_due = st.date_input('Plazo pago:', value=invoice_due, format="DD/MM/YYYY")
            with col5:
                form_invoice_iva = st.number_input('% IVA', min_value=0, max_value=21, step=21, value=21)
            with col6:
                form_invoice_desconto = st.number_input('% Retención', min_value=0, value=0)

            # Lines of the invoice
            with st.container():
                col, buff = st.columns([0.2, 0.8])
                num_rows = col.number_input('Nro. de lineas en la factura:', value=1, min_value=1, max_value=12, step=1)

                # columns to lay out the inputs
                grid = st.columns([0.35, 0.06, 0.10, 0.10, 0.10, 0.10, 0.10])
                total_invoice = 0.0
                base_imponible_sum = 0.0
                cuota_tributaria_sum = 0.0
                valor_retencion_sum = 0.0
                for row in range(num_rows):
                    with grid[0]:
                        line = st.text_input('Descripción *', value='', placeholder='', key=f'description{row}')
                    with grid[1]:
                        line_qty = st.number_input('Cant.', min_value=1, max_value=999, key=f'qty{row}')
                    with grid[2]:
                        line_value = st.number_input('Val.unit.', min_value=0.0, format="%0.2f", key=f'value{row}')
                    with grid[3]:
                        base_imponible = line_qty * line_value
                        line_base = st.number_input('Base imp.', value=base_imponible, format="%0.2f", disabled=True,
                                                    key=f'base{row}')
                    with grid[4]:
                        cuota_tributaria = base_imponible * form_invoice_iva / 100
                        line_cuotatrib = st.number_input('Cuota trib.', value=cuota_tributaria, format="%0.2f",
                                                         disabled=True, key=f'cuota{row}')
                    with grid[5]:
                        valor_retencion = base_imponible * form_invoice_desconto / 100
                        line_retencion = st.number_input('Val.ret.', value=valor_retencion, format="%0.2f",
                                                         disabled=True, key=f'retencion{row}')
                    with grid[6]:
                        total = base_imponible + cuota_tributaria - valor_retencion
                        line_total = st.number_input('Total', value=total, format="%0.2f", disabled=True,
                                                     key=f'total{row}')
                        total_invoice += total
                        base_imponible_sum += base_imponible
                        cuota_tributaria_sum += cuota_tributaria
                        valor_retencion_sum += valor_retencion

                # field to enter the invoice Note
                if form_invoice_iva == 0:  # if IVA=0% then this note is obligatiry in Spain
                    nota_iva0 = 'Operación de inversión del sujeto pasivo de acuerdo al artículo 84, apartado uno,' \
                                ' número 2o.f de la Ley 37/92 de IVA.'
                else:
                    nota_iva0 = ''

                form_invoice_note = st.text_area('Nota:', value=nota_iva0,
                                                 placeholder='Introduzca una nota para incluir en la factura...')

                # area to display the bank selection radiobutton and the total amount of the invoice
                col1, buff, col2 = st.columns([0.5, 0.15, 0.35])
                with col1:
                    st.divider()
                    banco = st.radio(
                        "Seleccione el banco donde pagar la factura",
                        ["CaixaBank", "Santander"],
                        captions=[
                            "IBAN: ES6221008444260200031531",
                            "IBAN: ES7500494700362217405170",
                        ],
                    )
                with col2:
                    st.divider()
                    # Formatting the value as string with point as thousand, comma as decimal and two decimal places
                    total_invoice_formatado = "€ {:,.2f}".format(total_invoice).replace(",", "X").replace(".",
                                                                                                          ",").replace(
                        "X", ".")
                    st.metric(label='Total factura', value=total_invoice_formatado)

            if total_invoice != 0:
                # print('Criando os registros para gravar a factura\n')
                # print('st.session_state;\n', st.session_state)
                if banco == 'CaixaBank':
                    form_banco = 'CaixaBank - IBAN: ES62-2100-8444-2602-0003-1531'
                else:  # banco = Santander
                    form_banco = 'Santander - IBAN: ES75-0049-4700-3622-1740-5170'

                for i in range(num_rows):
                    description = st.session_state[f'description{i}']
                    qty = st.session_state[f'qty{i}']
                    value = st.session_state[f'value{i}']
                    base = st.session_state[f'base{i}']
                    cuota = st.session_state[f'cuota{i}']
                    retencion = st.session_state[f'retencion{i}']
                    total = st.session_state[f'total{i}']
                    invoice_date = form_invoice_date.strftime("%d/%m/%Y")
                    invoice_due = form_invoice_due.strftime("%d/%m/%Y")
                    registro[i] = [
                        invoice_nr, client_cod, form_invoice_pedido, invoice_date, invoice_due, ' ',
                        description, qty, value, base, form_invoice_iva / 100, cuota, form_invoice_desconto / 100,
                        retencion, total, form_invoice_note, 'Recibir']
                    # print(f'Registro {i}:', registro[i])

                # Adjusting the values format to Euro
                base_imponible_sum_formatado = "€ {:,.2f}".format(base_imponible_sum).replace(",", "X").replace(".",
                                                                                                                ",").replace(
                    "X", ".")
                cuota_tributaria_sum_formatado = "€ {:,.2f}".format(cuota_tributaria_sum).replace(",", "X").replace(".",
                                                                                                                    ",").replace(
                    "X", ".")
                valor_retencion_sum_formatado = "€ {:,.2f}".format(valor_retencion_sum).replace(",", "X").replace(".",
                                                                                                                  ",").replace(
                    "X", ".")
                st.divider()

                # --- BEGIN of PDF document creation
                # New PDF document name
                new_document = f'Factura-{invoice_nr}'
                # Making a copy of the Google Doc template to create the new invoice in PDF
                copied_file = service_drive.files().copy(
                    fileId=INVOICE_TEMPLATE_ID,
                    body={'name': new_document}
                ).execute()

                # Get the document ID from the copied file
                new_document_id = copied_file.get('id')

                # Creating a dictionary for the substitutions (key: placeholder in template, value: to be inserted)
                substituicoes = {
                    'invoice_nr': str(invoice_nr),
                    'invoice_date': invoice_date,
                    'invoice_due': invoice_due,
                    'client': client,
                    'client_cif': client_cif,
                    'client_address': client_address,
                    'client_city': client_city,
                    'client_prov': client_prov,
                    'client_contact': client_contact,
                    'client_email': client_email,
                    'client_phone': str(client_phone),
                    'form_invoice_pedido': str(form_invoice_pedido),
                    'nota_iva0': nota_iva0,
                    'form_invoice_iva': str(form_invoice_iva),
                    'form_invoice_desconto': str(form_invoice_desconto),
                    'form_invoice_note': form_invoice_note,
                    'base_imponible_sum': str(base_imponible_sum_formatado),
                    'cuota_tributaria_sum': str(cuota_tributaria_sum_formatado),
                    'valor_retencion_sum': str(valor_retencion_sum_formatado),
                    'total_invoice': str(total_invoice_formatado),
                    'form_banco': form_banco
                }

                # Dinamically add the items (lines) of the invoice service descriptions in the dictionary
                # Maximum number of lines = 12 (set in the template file)
                for idx in range(num_rows):
                    description = st.session_state[f'description{idx}']
                    qty = st.session_state[f'qty{idx}']
                    value = st.session_state[f'value{idx}']
                    value_formatado = "{:,.2f}".format(value).replace(",", "X").replace(".", ",").replace(
                        "X", ".")
                    base = st.session_state[f'base{idx}']
                    base_formatado = "{:,.2f}".format(base).replace(",", "X").replace(".", ",").replace(
                        "X", ".")
                    substituicoes[f'description{idx}'] = description
                    substituicoes[f'qty{idx}'] = str(qty)
                    substituicoes[f'value{idx}'] = value_formatado
                    substituicoes[f'base{idx}'] = base_formatado

                for idx in range(12 - num_rows):  # Enter 'blank' in the rest of the lines with no content
                    substituicoes[f'description{idx + num_rows}'] = ''
                    substituicoes[f'qty{idx + num_rows}'] = ''
                    substituicoes[f'value{idx + num_rows}'] = ''
                    substituicoes[f'base{idx + num_rows}'] = ''
                # print('substituicoes >>>>>>\n', substituicoes)
                # Replace the placeholders in the document
                doc_factura = substituir_placeholders(new_document_id, substituicoes)
                # --- END of PDF document creation

                # Create columns and buttons to show PDF and SAVE the invoice in spreadsheet
                action0, action1, action2 = st.columns(3)
                with action0:  # DOWNLOAD PDF button
                    pdf_request = service_drive.files().export_media(fileId=new_document_id, mimeType='application/pdf')
                    # Usa io.BytesIO para manter o PDF em memória
                    pdf_data = pdf_request.execute()
                    st.download_button(
                        label="Ver archivo PDF",
                        data=pdf_data,
                        file_name=new_document,
                        mime="application/pdf"
                    )

                with action2:  # SAVE invoice button
                    add_invoice = st.button('Salvar Factura', type='primary', use_container_width=True)

                if add_invoice:  # Save the invoice
                    # Export document as PDF
                    pdf_request = service_drive.files().export_media(fileId=new_document_id, mimeType='application/pdf')
                    pdf_metadata = {
                        'name': new_document,
                        'parents': [PDF_FOLDER_ID]  # Coloca o arquivo dentro da pasta específica
                    }
                    media = googleapiclient.http.MediaIoBaseUpload(io.BytesIO(pdf_request.execute()),
                                                                   mimetype='application/pdf')
                    # print("salvando o PDF na pasta 'Facturas' no Google Drive...")
                    file = service_drive.files().create(body=pdf_metadata, media_body=media, fields='id').execute()
                    # print('Arquivo PDF salvo no Google Drive com ID:', file.get('id'))

                    # Create new record(s) with the invoice line(s) in the spreadsheet
                    for i in range(num_rows):
                        request = sheet.values().append(spreadsheetId=SPREADSHEET_ID,
                                                        range="facturas", valueInputOption="USER_ENTERED",
                                                        body={"values": [registro[i]]}).execute()
                    st.success(f'Factura {invoice_nr} creada con éxito.', icon='✅')
                    # Delete all session state keys
                    for key in st.session_state.keys():
                        del st.session_state[key]
                    st.session_state.client_key = ''

    st.divider()
