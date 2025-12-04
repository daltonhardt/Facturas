import io
import time

import pandas as pd
import streamlit as st
from streamlit_option_menu import option_menu

import googleapiclient
from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.http import MediaIoBaseUpload

from datetime import datetime
import locale
import json


# --- Constantes para o Template ---
START_TAG = '{{description0}}'
ITEM_PLACEHOLDERS = ['description', 'qty', 'value', 'base']


# Function to READ/GET values from spreadsheet
def leitura_worksheet(worksheet):
    try:
        resultado = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=worksheet).execute()
        values = resultado.get("values", [])
        df = pd.DataFrame(values)  # transform all values in DataFrame
        df.columns = df.iloc[0]  # set column names equal to values in row index position 0
        df = df[1:]  # remove the first row from DataFrame (column names)
        return df
    except (RuntimeError, TypeError, NameError):
        pass

# Function to READ/GET client values from spreadsheet
def leitura_registro_cliente(client_id):
    try:
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range='clientes').execute()
        values = result.get("values", [])

        header = values[0]  # first line contains header
        # print("=== header:\n", header)
        data = values[1:]    # next lines contain data
        # print("=== data:\n", data)

        # Filter records (rows) with status_cliente = "activo"
        status_index = header.index("status_cliente")
        ativos = [reg for reg in data if len(reg) > status_index and reg[status_index].strip().lower() == "activo"]
        return ativos[client_id]
        # print("values =\n", values)
        # return values[client_id] # return the record values in a list
    except (RuntimeError, TypeError, NameError):
        pass


@st.cache_data
def process_dynamic_items_table(template_id, document_name, substitutions):
    # Localiza a tabela com {{description0}}, apaga as linhas vazias do modelo e substitui os placeholders
    # Making a copy of the Google Docs template to create the new budget
    copied_file = service_drive.files().copy(
        fileId=template_id,
        body={'name': document_name}
    ).execute()

    # Get the document ID from the copied file
    document_id = copied_file.get('id')

    document = service_docs.documents().get(documentId=document_id).execute()

    table_start_index = -1
    start_row_index = -1
    requests = []
    total_linhas = 0

    # === 1️⃣ Localiza a tabela com o marcador START_TAG ===
    for element in document.get('body', {}).get('content', []):
        if 'table' in element:
            table_start_index = element.get('startIndex')
            table = element['table']
            for r_idx, reg in enumerate(table.get('tableRows', [])):
                for cell in reg.get('tableCells', []):
                    for content in cell.get('content', []):
                        if 'paragraph' in content:
                            for el in content['paragraph'].get('elements', []):
                                text = el.get('textRun', {}).get('content', '').strip()
                                if text == START_TAG:  # se o marcador de inicio for encontrado
                                    start_row_index = r_idx
                                    total_linhas = len(table['tableRows'])
                                    break
                        if start_row_index != -1: break
                    if start_row_index != -1: break
                if start_row_index != -1: break
        if table_start_index != -1 and start_row_index != -1:
            break

    if table_start_index == -1 or start_row_index == -1:
        raise ValueError(f'❌ No se pudo encontrar la tabla o el marcador. {START_TAG}.')

    # print(f'... ... START_TAG           = {START_TAG}')
    # print(f'... ... table_start_index   = {table_start_index}')
    # print(f'... ... start_row_index     = {start_row_index}')
    # print(f'... ... total_linhas table  = {total_linhas}')

    # === 2️⃣ Busca quantos itens existem ===
    item_indices = []
    for chave in substitutions.keys():
        if chave.startswith("description") and chave[-1].isdigit():
            try:
                indice = int(chave.replace("description", ""))
                item_indices.append(indice)
            except ValueError:
                pass
    num_items = len(item_indices)
    # print(f'... ... ... num_items = {num_items}')
    # print(f'... ... item_indices = {item_indices}')

    # === 3️⃣ Apaga as linhas de itens com placeholders restantes que não foram usadas ===
    for row_index in reversed(range(num_items + 1, total_linhas)):
        # print(f'... apagando a linha {row_index}...')
        requests.append({
            'deleteTableRow': {
                'tableCellLocation': {
                    'tableStartLocation': {'index': table_start_index},
                    'rowIndex': row_index
                }
            }
        })

    # === 4️⃣ Usa replaceAllText pra preencher os placeholders ===
    for placeholder, value_text in substitutions.items():
        requests.append({
            'replaceAllText': {
                'containsText': {
                    'text': '{{' + placeholder + '}}',
                    'matchCase': True
                },
                'replaceText': str(value_text)
            }
        })

    # print('>>> DUMP JSON:\n')
    # print(json.dumps(requests, indent=2))

    # === 5️⃣ Executa as alterações ===
    result = service_docs.documents().batchUpdate(
        documentId=document_id,
        body={'requests': requests}
    ).execute()

    # print(f"✅ {num_items} linha(s) e placeholders substituídos com sucesso!")
    return document_id, result


# ------- READ secret definitions -------
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
PRES_TEMPLATE_ID = st.secrets.google_definition["PRES_TEMPLATE_ID"]
PRES_PDF_FOLDER_ID = st.secrets.google_definition["PRES_PDF_FOLDER_ID"]
# for Google SHEETS
creds = service_account.Credentials.from_service_account_info(account_info, scopes=SCOPES)
service = build("sheets", "v4", credentials=creds)
# Call the Sheets API
sheet = service.spreadsheets()
# for Google DOCS 'v1' é a versão da Docs API
service_docs = build("docs", "v1", credentials=creds)
# Call the Sheets API
doc = service_docs.documents()
# for Google DRIVE
service_drive = build("drive", "v3", credentials=creds)
# ------ END Google Definitions -------

# ---- STREAMLIT INÍCIO DA INTERFACE E DO FLUXO  ----
st.set_page_config(layout="wide")
state= st.session_state
version_number = '2512.01'
st.sidebar.text(f'[ver. {version_number}]')
st.header("🧾 Presupuestos")

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
df_clientes_activos = df_clientes[df_clientes['status_cliente'] == 'Activo']
# st.dataframe(df_clientes_activos)

# Create DataFrame with ALL invoice values from spreadsheet
df_presupuestos = leitura_worksheet('presupuestos')

# Add client name (nombre_cliente) column mapping by client unique code (cod_cliente)
df_presupuestos['nombre_cliente'] = df_presupuestos.cod_cliente.map(
    df_clientes_activos.set_index('cod_cliente')['nombre_cliente'].to_dict())

# Adjusting the float numbers to match european standard
df_presupuestos['cantidad'] = df_presupuestos['cantidad'].apply(
    lambda x: float(str(x).replace('.', '').replace(',', '.')))
df_presupuestos['precio_unit'] = df_presupuestos['precio_unit'].apply(
    lambda x: float(str(x).replace('.', '').replace(',', '.')))
df_presupuestos['base_imponible'] = df_presupuestos['base_imponible'].apply(
    lambda x: float(str(x).replace('.', '').replace(',', '.')))
df_presupuestos['cuota_tributaria'] = df_presupuestos['cuota_tributaria'].apply(
    lambda x: float(str(x).replace('.', '').replace(',', '.')))
df_presupuestos['valor_retencion'] = df_presupuestos['valor_retencion'].apply(
    lambda x: float(str(x).replace('.', '').replace(',', '.')))
df_presupuestos['total'] = df_presupuestos['total'].apply(
    lambda x: float(str(x).replace('.', '').replace(',', '.')))


# Group all records with the same invoice number and get the sum
df_total_presupuestos = df_presupuestos.groupby(['nro_pres'], as_index=False).agg(
    {'total': 'sum', 'nombre_cliente': 'first', 'descripcion': 'first', 'fecha_emision': 'first'})
df_total_presupuestos['total'] = df_total_presupuestos['total'].apply(lambda x: f'{x:.2f}')
df_total_presupuestos.index += 1  # make index start at 1

# Tabs
tab_options = ['Crear nuevo', 'Listar todo']
TAB_1 = tab_options[0]
TAB_2 = tab_options[1]

tab = option_menu(
    menu_title='',
    options=tab_options,
    icons=['bi-pencil-square', 'list-task'],
    menu_icon='cast',
    orientation='horizontal',
)

# Create NEW
if tab == TAB_1:
    st.divider()

    # last_budget_row = len(df_presupuestos)  # get the last written row from the dataframe
    # last_budget = df_presupuestos.loc[last_budget_row, 'nro_pres']  # get Budget number from column 'nro_pres'
    # budget_nr = str(int(last_budget) + 1).zfill(4)  # zfill=4 format 9999
    if not df_presupuestos.empty:
        last_budget = int(df_presupuestos['nro_pres'].astype(int).max())
    else:
        last_budget = 0
    budget_nr = str(last_budget + 1).zfill(4)  # zfill=4 format 9999

    current_year = datetime.now().strftime('%y')  # get the current year with two-digits
    current_month = datetime.now().strftime('%m')  # get the current year with two-digits

    st.subheader('Nuevo Presupuesto: ' + budget_nr)

    if 'client_key' not in state:
        state.client_key = None

    col_cliente, buff = st.columns([0.7, 0.3])
    with col_cliente:
        # select Client name from dataframe
        client = st.selectbox('Cliente *', df_clientes_activos['nombre_cliente'].sort_values(),
                              index=None, placeholder='Seleccione...', key='client_key')

    if client is not None:
        
        # ** IMPORTANTE **
        # os nomes das colunas importadas do DF_TO_SAVE: 'Descripción',        'Cantidad',   'Valor'
        # devem combinar com as colunas do FORM NOVO   : 'budget_description', 'budget_qty', 'budget_value'
        if 'df_to_save' in state:
            if st.button('Rellenar líneas con datos importados', type='primary'):
                items_df = state.df_to_save
                num_rows = len(state.df_to_save)
                state.num_rows_input = num_rows
                for i, row in items_df.iterrows():
                    state[f'budget_description{i}'] = row['Descripción']
                    state[f'budget_qty{i}'] = float(row['Cantidad'])
                    state[f'budget_value{i}'] = float(row['Valor'])
                st.rerun() # Re-executa para aplicar os novos valores do session_state nos inputs
                # st.write(f'Total de {num_rows} linha(s) importada(s).')

        if 'df_to_save' in state and state.get('num_rows_input') is not None:
            # Usa o valor definido no clique do botão
            initial_num_rows = state.num_rows_input
        else:
            # Valor inicial padrão
            initial_num_rows = 1
        
        col_lines, buff = st.columns([0.2, 0.8])
        with col_lines:
            num_rows = st.number_input('Nro. lineas:', value = initial_num_rows, min_value=1, max_value=99, step=1)

        with st.container(border=True):
        # with st.form('budget_form'):
            for index_cliente in range(len(df_clientes_activos)):
                if df_clientes_activos.iloc[index_cliente]['nombre_cliente'] == client:
                    reg_cliente = leitura_registro_cliente(index_cliente)
                    client_cod = reg_cliente[0]
                    client_cif = reg_cliente[2]
                    client_prov = reg_cliente[3]
                    client_city = reg_cliente[4]
                    client_address = reg_cliente[5]
                    client_postal = reg_cliente[6]
                    client_contact = reg_cliente[7]
                    client_email = reg_cliente[8]
                    client_phone = reg_cliente[9]
                    client_obs = reg_cliente[10]
                    break

            budget_date = datetime.now()  # current date

            col1, col2, col3 = st.columns(3)
            with col1:
                form_budget_date = st.date_input('Fecha emissión:', value=budget_date, format="DD/MM/YYYY")
            with col2:
                form_budget_iva = st.number_input('% IVA', min_value=0, max_value=21, step=21, value=21)
            with col3:
                form_budget_desconto = st.number_input('% Retención', min_value=0, value=0)

            # columns to lay out the inputs
            grid = st.columns([0.42, 0.08, 0.10, 0.10, 0.10, 0.10, 0.10])
            total_budget = 0.0
            base_imponible_sum = 0.0
            cuota_tributaria_sum = 0.0
            valor_retencion_sum = 0.0
            for row in range(num_rows):
                with grid[0]:
                    description_key = f'budget_description{row}'
                    if description_key not in state:
                        state[description_key] = ''
                    budget_line = st.text_input('Descripción *', placeholder='', key=description_key)
                with grid[1]:
                    qty_key = f'budget_qty{row}'
                    if qty_key not in state:
                        state[qty_key] = 1.0
                    budget_line_qty = st.number_input('Cant.', min_value=0.01, key=qty_key)
                with grid[2]:
                    value_key = f'budget_value{row}'
                    if value_key not in state:
                        state[value_key] = 0.0
                    budget_line_value = st.number_input('Val.unit.', format="%0.2f", key=value_key)
                with grid[3]:
                    budget_base_imponible = budget_line_qty * budget_line_value
                    base_key = f'budget_base{row}'
                    state[base_key] = budget_base_imponible
                    line_base = st.number_input('Base imp.', format="%0.2f", disabled=True, key=base_key)
                with grid[4]:
                    budget_cuota_tributaria = budget_base_imponible * form_budget_iva / 100
                    # cuota_key = f'budget_cuota{row}_{line_qty}_{line_value}_{form_budget_iva}'
                    cuota_key = f'budget_cuota{row}'
                    state[cuota_key] = budget_cuota_tributaria
                    line_cuotatrib = st.number_input('Cuota trib.', format="%0.2f",
                                                     disabled=True, key=cuota_key)
                with grid[5]:
                    budget_valor_retencion = budget_base_imponible * form_budget_desconto / 100
                    # retencion_key = f'budget_retencion{row}_{line_qty}_{line_value}_{valor_retencion}'
                    retencion_key = f'budget_retencion{row}'
                    state[retencion_key] = budget_valor_retencion
                    line_retencion = st.number_input('Val.ret.', format="%0.2f",
                                                     disabled=True, key=retencion_key)
                with grid[6]:
                    budget_total = budget_base_imponible + budget_cuota_tributaria - budget_valor_retencion
                    # total_key = f'budget_total{row}_{line_qty}_{line_value}_{total}'
                    total_key = f'budget_total{row}'
                    state[total_key] = budget_total
                    line_total = st.number_input('Total', format="%0.2f", disabled=True,
                                                 key=total_key)
                    total_budget += budget_total
                    base_imponible_sum += budget_base_imponible
                    cuota_tributaria_sum += budget_cuota_tributaria
                    valor_retencion_sum += budget_valor_retencion

            # field to enter the budget Note
            if form_budget_iva == 0:  # if IVA=0% then this note is obligatory in Spain
                budget_nota_iva0 = 'Operación de inversión del sujeto pasivo de acuerdo al artículo 84, apartado uno,' \
                                   ' número 2o.f de la Ley 37/92 de IVA.'
            else:
                budget_nota_iva0 = ''

            form_budget_note = st.text_area('Nota:', value=budget_nota_iva0,
                                            placeholder='Introduzca una nota para incluir en el presupuesto...')

            # area to display the total amount of the budget
            buff, col2 = st.columns([0.65, 0.35])
            with col2:
                st.divider()
                # Formatting the value as string with point as thousands, comma as decimal and two decimal places
                total_budget_formatado = "€ {:,.2f}".format(total_budget).replace(",", "X").replace(".",
                                                                                                    ",").replace(
                    "X", ".")
                st.metric(label='TOTAL', value=total_budget_formatado)

            if total_budget != 0:
                # st.write(f'num_rows = {num_rows}')
                registro = []
                for i in range(num_rows):
                    # st.write(f'valor de i = {i}')
                    description = state[f'budget_description{i}']
                    qty = state[f'budget_qty{i}']
                    value = state[f'budget_value{i}']
                    base = state[f'budget_base{i}']
                    cuota = state[f'budget_cuota{i}']
                    retencion = state[f'budget_retencion{i}']
                    total = state[f'budget_total{i}']
                    budget_date = form_budget_date.strftime("%-d/%m/%Y")
                    row = [
                        budget_nr,                      # numero del presupueto
                        client_cod,                     # codigo del cliente
                        budget_date,                    # fecha de emisión
                        description,                    # descripción
                        qty,                            # cantidad
                        value,                          # precio unitário
                        base,                           # base imponible
                        form_budget_iva / 100,          # percentual IVA
                        cuota,                          # cuota tributaria
                        form_budget_desconto / 100,     # percentual retención
                        retencion,                      # valor retención
                        total,                          # total del presupuesto
                        form_budget_note                # nota
                    ]
                    registro.append(row)
                    # st.write(f'Registro {i}:', row)

                # Adjusting the values format to Euro
                base_imponible_sum_formatado = "€ {:,.2f}".format(base_imponible_sum).replace(",", "X").replace(".",
                                                                                                                ",").replace(
                    "X", ".")
                cuota_tributaria_sum_formatado = "€ {:,.2f}".format(cuota_tributaria_sum).replace(",", "X").replace(
                    ".",
                    ",").replace(
                    "X", ".")
                valor_retencion_sum_formatado = "€ {:,.2f}".format(valor_retencion_sum).replace(",", "X").replace(
                    ".",
                    ",").replace(
                    "X", ".")
                st.divider()

                # Creating a dictionary for the substitutions (key: placeholder in template, value: to be inserted)
                substituicoes = {
                    'pres_nr': str(budget_nr),
                    'pres_date': budget_date,
                    'client': client,
                    'client_cif': client_cif,
                    'client_address': client_address,
                    'client_postal': client_postal,
                    'client_city': client_city,
                    'client_prov': client_prov,
                    'client_contact': client_contact,
                    'client_email': client_email,
                    'client_phone': str(client_phone),
                    'form_pres_iva': str(form_budget_iva),
                    'form_pres_desconto': str(form_budget_desconto),
                    'form_pres_note': form_budget_note,
                    'base_imponible_sum': str(base_imponible_sum_formatado),
                    'cuota_tributaria_sum': str(cuota_tributaria_sum_formatado),
                    'valor_retencion_sum': str(valor_retencion_sum_formatado),
                    'total_pres': str(total_budget_formatado)
                }

                # Dynamically add the items (lines) of the invoice service descriptions in the dictionary
                for idx in range(num_rows):
                    description = state[f'budget_description{idx}']
                    qty = state[f'budget_qty{idx}']
                    value = state[f'budget_value{idx}']
                    value_formatado = "{:,.2f}".format(value).replace(",", "X").replace(".", ",").replace(
                        "X", ".")
                    base = state[f'budget_base{idx}']
                    base_formatado = "{:,.2f}".format(base).replace(",", "X").replace(".", ",").replace(
                        "X", ".")
                    substituicoes[f'description{idx}'] = description
                    substituicoes[f'qty{idx}'] = str(qty)
                    substituicoes[f'value{idx}'] = value_formatado
                    substituicoes[f'base{idx}'] = base_formatado

                # print('Total substituições >>>>>>\n', substituicoes)

                # New document name
                new_document_name = f'Presupuesto-{budget_nr}'

                # Create columns and buttons to show PDF and SAVE the invoice in spreadsheet
                action1, action2 = st.columns(2)

                with action1:  # DOWNLOAD PDF button
                    # Replace the placeholders in the document
                    new_document_id, _ = process_dynamic_items_table(
                        template_id = PRES_TEMPLATE_ID,
                        document_name=new_document_name,
                        substitutions=substituicoes
                    )

                    pdf_request = service_drive.files().export_media(fileId=new_document_id, mimeType='application/pdf')
                    pdf_data = pdf_request.execute()
                    st.download_button(
                        label="Ver archivo PDF",
                        data=pdf_data,
                        file_name=new_document_name,
                        mime="application/pdf"
                    )
                    # print("Processamento do documento concluído com sucesso.")

                with action2:  # SAVE budget button
                    add_budget = st.button('Guardar Presupuesto en la base de datos', type='primary', use_container_width=True)

                if add_budget: # save the budget (presupuesto)
                    # Replace the placeholders in the document
                    new_document_id, _ = process_dynamic_items_table(
                        template_id=PRES_TEMPLATE_ID,
                        document_name=new_document_name,
                        substitutions=substituicoes
                    )

                    # export document as PDF
                    pdf_request = service_drive.files().export_media(fileId=new_document_id, mimeType='application/pdf')
                    pdf_metadata = {
                        'name': new_document_name,
                        'parents': [PRES_PDF_FOLDER_ID] # coloca o arquivo PDF dentro da pasta específica
                    }
                    media = googleapiclient.http.MediaIoBaseUpload(io.BytesIO(pdf_request.execute()),
                                                                   mimetype='application/pdf')
                    # print("salvando o PDF na pasta 'Presupuestos' no Google Drive...")
                    file = service_drive.files().create(body=pdf_metadata, media_body=media, fields='id').execute()
                    # print('Arquivo PDF salvo no Google Drive com ID:', file.get('id'))

                    # Create new record(s) with the invoice line(s) in the spreadsheet
                    for row in registro:
                        request = sheet.values().append(spreadsheetId=SPREADSHEET_ID,
                                                        range="presupuestos",
                                                        valueInputOption="USER_ENTERED",
                                                        body={"values": [row]}
                                                        ).execute()
                    st.success(f'Presupuesto {budget_nr} creado con éxito.', icon='✅')
                    # print(f" state antes:\n {state}")
                    time.sleep(3)  # wait 3 seconds to display the success to the user
                    # Delete all session state keys
                    for key in state.keys():
                        del state[key]
                    # print(f"\n state depois:\n {state}")
                    st.rerun()  # roda o ‘App’ para zerar o campo com nome do cliente

if tab == TAB_2:
    st.divider()
    # Show ALL presupuestos
    st.subheader('Todos Presupuestos')
    st.dataframe(df_total_presupuestos)
