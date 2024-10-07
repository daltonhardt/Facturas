import streamlit as st
from streamlit_option_menu import option_menu
from googleapiclient.discovery import build
from google.oauth2 import service_account
import pandas as pd
import json


# Function to GET values from province/cities and CACHE the DataFrame
@st.cache_data
def leitura_worksheet_cities():
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range='ciudades').execute()
    values = result.get("values", [])
    df = pd.DataFrame(values)  # transform all values in DataFrame
    df.columns = df.iloc[0]  # set column names equal to values in row index position 0
    df = df[1:]  # remove first row from DataFrame
    return df


# Function to get values from spreadsheet
def leitura_worksheet(worksheet):
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=worksheet).execute()
    values = result.get("values", [])
    df = pd.DataFrame(values)  # transform all values in DataFrame
    df.columns = df.iloc[0]  # set column names equal to values in row index position 0
    df = df[1:]  # remove first row from DataFrame
    return df


# function to DELETE a client with a Dialog popup window
@st.experimental_dialog("A T E N C I Ó N")
def delete_client(*args):
    st.write(f"¿Confirma?")
    if st.button("OK"):
        index_client = args[1] + 1
        change_status_cliente = ['Inactivo']
        celula = "clientes!L" + str(index_client)
        sheet.values().update(spreadsheetId=SPREADSHEET_ID,
                              range=celula, valueInputOption="USER_ENTERED",
                              body={"values": [change_status_cliente]}).execute()
        # print(f"Registro: {args[2]} - {args[3]} na celula {celula} foi ¡borrado!")
        st.success(f"¡El cliente  {args[3]}  ha sido borrado!", icon='✅')
        st.rerun()

    elif st.button("Cancel"):
        # print("Delete foi Cancelado")
        st.rerun()


# function to REACTIVATE a client with a Dialog popup window
@st.experimental_dialog("A T E N C I Ó N")
def reactivate_client(*args):
    st.write(f"¿Confirma?")
    if st.button("OK"):
        index_client = args[1] + 1
        change_status_cliente = ['Activo']
        celula = "clientes!L" + str(index_client)
        sheet.values().update(spreadsheetId=SPREADSHEET_ID,
                              range=celula, valueInputOption="USER_ENTERED",
                              body={"values": [change_status_cliente]}).execute()
        # print(f"Registro: {args[2]} - {args[3]} na celula {celula} foi ¡borrado!")
        st.success(f"¡El cliente  {args[3]}  ha sido reactivado!", icon='✅')
        st.rerun()

    elif st.button("Cancel"):
        # print("Delete foi Cancelado")
        st.rerun()


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
PDF_FOLDER_ID = st.secrets.google_definition["PDF_FOLDER_ID"]

# ------- BEGIN Google Definitions -------
# for Google SHEETS
creds = service_account.Credentials.from_service_account_info(account_info, scopes=SCOPES)
service = build("sheets", "v4", credentials=creds)
# Call the Sheets API
sheet = service.spreadsheets()
# ------ END Google Definitions -------

# --- Starting Streamlit
st.set_page_config(layout="wide")
st.header("Registro de clientes 🔤")
# st.sidebar.markdown("# Clientes 🔤")

# Create DataFrame with ALL client values from spreadsheet
df_clientes = leitura_worksheet('clientes')
# st.dataframe(df_clientes)

# Create DataFrame with ONLY ACTIVE client values from spreadsheet
df_clientes_activos = df_clientes[df_clientes['status_cliente'] == 'Activo'].reset_index(drop=True)
df_clientes_activos.index += 1  # making index start from 1 to stay equal with "df_clientes"
# st.dataframe(df_clientes_activos)

# Create DataFrame with ONLY INACTIVE client values from spreadsheet
df_clientes_inactivos = df_clientes[df_clientes['status_cliente'] == 'Inactivo'].reset_index(drop=True)
df_clientes_inactivos.index += 1  # making index start from 1 to stay equal with "df_clientes"
# st.dataframe(df_clientes_inactivos)

# Tabs
TAB_0 = 'Clientes Activos'
TAB_1 = 'Clientes Inactivos'
TAB_2 = 'Añadir nuevo Cliente'
tab = option_menu(
    menu_title='',
    options=['Clientes Activos', 'Clientes Inactivos', 'Añadir nuevo Cliente'],
    icons=['star', 'bi-emoji-frown', 'person-plus-fill'],
    menu_icon='cast',
    orientation='horizontal',
    default_index=0
)

if tab == TAB_0:  # Show ONLY ACTIVE clients
    df_clientes_activos = df_clientes_activos.sort_values(['nombre_cliente']).reset_index(drop=True)
    with st.container():
        EXPANDED = st.checkbox('Expandir/contraer todo', value=False)
        num_columns = 2
        columns = st.columns(num_columns, gap='small')
        for index, row in df_clientes_activos.iterrows():
            # print(index, '--', index % num_columns)
            with columns[index % num_columns]:
                with st.expander(row['nombre_cliente'], expanded=EXPANDED):
                    col1, col2 = st.columns([0.4, 0.6])
                    with col1:
                        st.markdown(f"**Codigo**  \n"
                                    f"**CIF/NIF**  \n"
                                    f"**Contacto**  \n"
                                    f"**Teléfono**  \n"
                                    f"**Prov/Ciudad**  \n"
                                    f"**Dirección**  \n"
                                    f"**Cod. Postal**  \n"
                                    f"**Email**  \n"
                                    f"**Obs.**  "
                                    )
                    with col2:
                        st.markdown(f"{row['cod_cliente']}  \n"
                                    f"{row['cif']}  \n"
                                    f"{row['contacto']}  \n"
                                    f"{row['telefono']}  \n"
                                    f"{row['provincia']}, {row['ciudad']}  \n"
                                    f"{row['direccion']}  \n"
                                    f"{row['postal']}  \n"
                                    f"{row['email']}  \n"
                                    f"{row['obs']}"
                                    )
                    if "borrar" not in st.session_state:
                        m = st.markdown("""
                                        <style>
                                        div.stButton > button:first-child {
                                            background-color: rgb(204, 49, 49);
                                        }
                                        </style>""", unsafe_allow_html=True)
                        if st.button("Borrar", key='delete_' + str(index)):
                            # print('index no df_clientes_activos = ', str(index))
                            for index_cliente in range(len(df_clientes)):
                                if df_clientes.iloc[index_cliente]['cod_cliente'] == row['cod_cliente']:
                                    # print('index no df_clientes = ', str(index_cliente + 1))
                                    break
                            index_cliente += 1
                            delete_client(df_clientes, index_cliente, row['cod_cliente'], row['nombre_cliente'])

if tab == TAB_1:  # Show ONLY INACTIVE clients
    df_clientes_inactivos = df_clientes_inactivos.sort_values(['nombre_cliente']).reset_index()
    with st.container():
        EXPANDED = st.checkbox('Expandir/contraer todo', value=False)
        num_columns = 2
        columns = st.columns(num_columns, gap='small')
        for index, row in df_clientes_inactivos.iterrows():
            # print(index, '--', index % num_columns)
            with columns[index % num_columns]:
                with st.expander(row['nombre_cliente'], expanded=EXPANDED):
                    col1, col2 = st.columns([0.4, 0.6])
                    with col1:
                        st.markdown(f"**Codigo**  \n"
                                    f"**CIF/NIF**  \n"
                                    f"**Contacto**  \n"
                                    f"**Teléfono**  \n"
                                    f"**Prov/Ciudad**  \n"
                                    f"**Dirección**  \n"
                                    f"**Cod. Postal**  \n"
                                    f"**Email**  \n"
                                    f"**Obs.**  "
                                    )
                    with col2:
                        st.markdown(f"{row['cod_cliente']}  \n"
                                    f"{row['cif']}  \n"
                                    f"{row['contacto']}  \n"
                                    f"{row['telefono']}  \n"
                                    f"{row['provincia']}, {row['ciudad']}  \n"
                                    f"{row['direccion']}  \n"
                                    f"{row['postal']}  \n"
                                    f"{row['email']}  \n"
                                    f"{row['obs']}"
                                    )
                    if "Reactivar" not in st.session_state:
                        m = st.markdown("""
                                        <style>
                                        div.stButton > button:first-child {
                                            background-color: rgb(0, 180, 0);
                                        }
                                        </style>""", unsafe_allow_html=True)
                        if st.button("Reactivar", key='reactivate_' + str(index)):
                            # print('index no df_clientes_activos = ', str(index))
                            for index_cliente in range(len(df_clientes)):
                                if df_clientes.iloc[index_cliente]['cod_cliente'] == row['cod_cliente']:
                                    # print('index no df_clientes = ', str(index_cliente + 1))
                                    break
                            index_cliente += 1
                            reactivate_client(df_clientes, index_cliente, row['cod_cliente'], row['nombre_cliente'])

if tab == TAB_2:  # Create NEW Client
    with st.container():
        last_client_code = len(df_clientes)
        new_client_code = 'C' + str(last_client_code + 1).zfill(3)  # zfill=3 p/formato 'C000'
        st.subheader('Codigo: ' + new_client_code)
        new_client_name = st.text_input('Nombre Cliente:')
        new_cif = st.text_input('CIF/NIF:')
        # getting unique province names from dataframe
        df_ciudades = leitura_worksheet_cities()  # run function to cache the list of provinces/cities
        new_province = st.selectbox('Provincia:', df_ciudades['provincia'].drop_duplicates().sort_values(),
                                    index=None, placeholder='Seleccione')
        # st.write("Usted ha selecionado:", new_province)
        # getting all cities from seleted province
        ciudades = df_ciudades[df_ciudades['provincia'] == new_province].reset_index(drop=True).drop(
            columns=['provincia'])
        new_city = st.selectbox('Ciudad:', options=ciudades,
                                index=None, placeholder='Selecione')
        # st.write("Usted ha selecionado:", new_city)
        new_address = st.text_input('Dirección:')
        new_postal = st.text_input('Codigo Postal:')
        new_contact = st.text_input('Persona contacto:')
        new_email = st.text_input('Correo:')
        new_phone = st.text_input('Teléfono:')
        new_obs = st.text_input('Obs:')
        new_status_cliente = 'Activo'
        m = st.markdown("""
                        <style>
                        div.stButton > button:first-child {
                            background-color: rgb(0, 150, 0);
                        }
                        </style>""", unsafe_allow_html=True)
        save = st.button('Salvar')
        if save:
            new_list = [new_client_code, new_client_name, new_cif, new_province, new_city, new_address,
                        new_postal, new_contact, new_email, new_phone, new_obs, new_status_cliente]
            request = sheet.values().append(spreadsheetId=SPREADSHEET_ID,
                                            range="clientes", valueInputOption="USER_ENTERED",
                                            body={"values": [new_list]}).execute()
            st.success(f"¡Nuevo cliente  {new_client_name}  creado!", icon='✅')


# show all Clients
st.divider()
st.subheader("Lista de TODOS los clientes")
st.dataframe(df_clientes)
