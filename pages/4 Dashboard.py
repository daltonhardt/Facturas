import streamlit as st
from streamlit_option_menu import option_menu
from googleapiclient.discovery import build
from google.oauth2 import service_account
import pandas as pd
import locale
from datetime import datetime, date
import json


# Function to get values from spreadsheet
def leitura_worksheet(worksheet):
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=worksheet).execute()
    values = result.get("values", [])
    df = pd.DataFrame(values)  # transform all values in DataFrame
    df.columns = df.iloc[0]  # set column names equal to values in row index position 0
    df = df[1:]  # remove first row from DataFrame
    return df


# Read credentials from secrets.toml file
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

# Converte a string JSON para um dicionário Python
account_info = json.loads(account_info_str)

SCOPES = st.secrets.google_definition["SCOPES"]
SPREADSHEET_ID = st.secrets.google_definition["SPREADSHEET_ID"]
INVOICE_TEMPLATE_ID = st.secrets.google_definition["INVOICE_TEMPLATE_ID"]
INVOICE_PDF_FOLDER_ID = st.secrets.google_definition["INVOICE_PDF_FOLDER_ID"]

# ------- BEGIN Google Definitions -------
# for Google SHEETS
creds = service_account.Credentials.from_service_account_info(account_info, scopes=SCOPES)
service = build("sheets", "v4", credentials=creds)
# Call the Sheets API
sheet = service.spreadsheets()
# ------ END Google Definitions -------

# --- Starting Streamlit
st.set_page_config(
    layout="wide",
    menu_items={
        'About': "# Facturas *EL SHADDAI*"
    }
)
version_number = '2512.01'
st.sidebar.text(f'[ver. {version_number}]')

# get today date
today = date.today()
YEAR = today.strftime("%Y")
year = today.strftime("%y")

selected = option_menu(menu_title=f'Dashboard {YEAR}', options=['Overview'], icons=['bar-chart-fill'], menu_icon='cast',
                       orientation='horizontal')

# st.sidebar.markdown("# Dashboard 📈")

# set the locale to Spanish (Spain)
locale.setlocale(locale.LC_NUMERIC, 'es_ES.UTF-8')

# pd.options.display.float_format = "{:,.2f}".format
pd.set_option('display.precision', 2)
# Aplicar o formato de duas casas decimais apenas na exibição
pd.options.display.float_format = '{:.2f}'.format
# Pandas visualization
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

# Create DataFrame with ALL Client values from spreadsheet
df_clientes = leitura_worksheet('clientes')

# Create DataFrame with ALL Invoices values from spreadsheet
df_facturas = leitura_worksheet('facturas')

# Cria no dataframe as colunas de Mês e Ano da 'fecha_emision'
df_facturas['year'] = pd.to_datetime(df_facturas['fecha_emision'], dayfirst=True).dt.year - 2000
df_facturas['month'] = pd.to_datetime(df_facturas['fecha_emision'], dayfirst=True).dt.month
# Extract data for current year
df_facturas = df_facturas[df_facturas['year'] == int(year)].reset_index()

# Adjusting the float numbers
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

# Add the client name 'nombre_cliente' from 'clientes' spreadsheet into df_facturas
df_facturas['nombre_cliente'] = df_facturas.cod_cliente.map(
    df_clientes.set_index('cod_cliente')['nombre_cliente'].to_dict())

# Group by 'nro_factura' and SUM each invoice
df_total_facturas = df_facturas.groupby(['nro_factura'], as_index=False).agg(
    {'total': 'sum', 'nombre_cliente': 'first', 'descripcion': 'first', 'fecha_emision': 'first', 'plazo_pago': 'first',
     'status': 'first', 'month': 'first', 'year': 'first'})
df_total_facturas.index += 1
# df_total_facturas['total'] = df_total_facturas['total'].apply(lambda x: f'{x:.2f}')

col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    # Total invoice sum
    total_facturas_sum = df_total_facturas['total'].sum()
    # Formatando o valor como string com separador de milhar e duas casas decimais
    total_facturas_sum_formatado = "€ {:,.2f}".format(total_facturas_sum).replace(",", "X").replace(".", ",").replace("X", ".")
    st.metric(label='Total Acumulado', value=total_facturas_sum_formatado)
with col2:
    # Total invoice sum Paid
    df_facturas_pagadas = df_total_facturas[df_total_facturas['status'].isin(['Pagado'])].reset_index(
            drop=True)
    total_facturas_pagadas_sum = df_facturas_pagadas['total'].sum()
    total_facturas_pagadas_sum_formatado = "€ {:,.2f}".format(total_facturas_pagadas_sum).replace(",", "X").replace(".", ",").replace("X", ".")
    st.metric(label='Total Pagado', value=total_facturas_pagadas_sum_formatado)
with col3:
    # Total invoice sum to receive
    df_facturas_recibir = df_total_facturas[df_total_facturas['status'].isin(['Recibir'])].reset_index(
            drop=True)
    total_facturas_recibir_sum = df_facturas_recibir['total'].sum()
    total_facturas_recibir_sum_formatado = "€ {:,.2f}".format(total_facturas_recibir_sum).replace(",", "X").replace(".", ",").replace("X", ".")
    st.metric(label='Total a Recibir', value=total_facturas_recibir_sum_formatado)
with col4:
    # Total invoice sum not paid
    df_facturas_nopagas = df_total_facturas[df_total_facturas['status'].isin(['Atrasado'])].reset_index(
            drop=True)
    total_facturas_nopagas_sum = df_facturas_nopagas['total'].sum()
    total_facturas_nopagas_sum_formatado = "€ {:,.2f}".format(total_facturas_nopagas_sum).replace(",", "X").replace(".", ",").replace("X", ".")
    st.metric(label='Total No Pagado', value=total_facturas_nopagas_sum_formatado)
with col5:
    # Total invoice sum cancelled
    df_facturas_canceladas = df_total_facturas[df_total_facturas['status'].isin(['Cancelado'])].reset_index(
            drop=True)
    total_facturas_canceladas_sum = df_facturas_canceladas['total'].sum()
    total_facturas_canceladas_sum_formatado = "€ {:,.2f}".format(total_facturas_canceladas_sum).replace(",", "X").replace(".", ",").replace("X", ".")
    st.metric(label='Total Cancelado', value=total_facturas_canceladas_sum_formatado)

st.bar_chart(df_total_facturas, x='month', y='total', x_label='Mes', y_label='Total', color='month')
columns_to_show = ['nro_factura', 'fecha_emision', 'total', 'nombre_cliente', 'descripcion', 'status']
st.dataframe(df_total_facturas, column_order=columns_to_show)
