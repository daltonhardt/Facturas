import logging
# Ocultar INFO e DEBUG do Google API Client
logging.basicConfig(level=logging.CRITICAL)
logging.getLogger().setLevel(logging.CRITICAL)
logging.disable(logging.CRITICAL)
# logging.getLogger("google").setLevel(logging.ERROR)
# logging.getLogger("googleapiclient").setLevel(logging.ERROR)
# logging.getLogger("google.auth").setLevel(logging.ERROR)
# logging.getLogger("googleapiclient.discovery_cache").setLevel(logging.ERROR)
# logging.getLogger("googleapiclient.discovery_cache").propagate=False
# logging.getLogger("googleapiclient.discovery_cache").setLevel(logging.CRITICAL)
# logging.getLogger("googleapiclient.http").setLevel(logging.ERROR)
# logging.getLogger().setLevel(logging.WARNING)

import os
import re
import io
import json
import time
import numpy as np
import streamlit as st
from google import genai
import mimetypes
from google.genai.errors import APIError
import tempfile
import pandas as pd
from pandas import DataFrame
import unicodedata
from unstract.llmwhisperer import LLMWhispererClientV2


def fix_json_quotes(s: str) -> str:
    """
    Escapa aspas duplas internas que aparecem DENTRO de valores string JSON,
    sem tocar nas aspas delimitadoras reais.
    """
    out = []
    i = 0
    n = len(s)
    in_string = False
    while i < n:
        ch = s[i]

        # Handle backslash escapes: copy backslash and next char as-is (if exists)
        if ch == '\\':
            # copia a barra e o próximo caractere se houver
            if i + 1 < n:
                out.append(s[i])
                out.append(s[i+1])
                i += 2
                continue
            else:
                out.append(ch)
                i += 1
                continue

        if ch == '"':
            if not in_string:
                # abre string JSON
                in_string = True
                out.append(ch)
                i += 1
                continue
            else:
                # possível fechamento — olha o próximo caractere não-espaço
                j = i + 1
                while j < n and s[j].isspace():
                    j += 1
                next_char = s[j] if j < n else None

                # Se o próximo caractere não-espaço é um delimitador JSON válido,
                # então esta aspa é o fechamento da string.
                if next_char in (',', '}', ']', ':') or next_char is None:
                    in_string = False
                    out.append(ch)  # mantém como fechamento
                    i += 1
                    continue
                else:
                    # é uma aspa DENTRO do texto → precisa ser escapada
                    out.append('\\"')
                    i += 1
                    continue
        else:
            out.append(ch)
            i += 1

    return ''.join(out)


def format_hms(segundos):
    h = int(segundos // 3600)
    m = int((segundos % 3600) // 60)
    s = int(segundos % 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


# Define Gemini AI prompt instructions
prompt = ("You work for an industrial service company and your job is to create a budget offer "
          "('presupuesto', in spanish) based on a customer request."
          "I have used LLMWhisperer to get a text file of the request for you."
          "It is important to keep track of the line descriptions, unit quantity and the value of the items. "
          "Do your best to extract texts and numbers in this order."
          "In spanish the quantity might be called 'unid' or 'cant' and the item description as 'concepto'. "
          "Return the data in a JSON format with the key values: 'Descripción','Cantidad','Valor', "
          "which represents the item description, quantity and value."
          "Clean the JSON file to consider only valid scape characters."
          "The 'Descripción' field should contain a complete description of the item or service requested. "
          "The 'Cantidad' field should contain the quantity of the item or service in decimal point. "
          "The 'Valor' should contain the unitary value of the item or service requested with decimal point. "
          "Sometimes the description of the items or services is split in two or more lines, try your best to "
          "get all the lines of the description. "
          "Sometimes the quantity and value of the items or services are blank or dashed, do not invent numbers "
          "and in this case leave a blank field in the JSON file. "
          "Return ONLY a JSON array. No text before or after. If you cannot extract data, return []."
          "The input text may contain numbers separated by commas (European format), "
          "convert them to decimal point (e.g., '1,00' to '1.00')."
          "CRITICAL: Always enclose all string values in double quotes and ensure your output is strictly a "
          "valid JSON array. DO NOT include anything other than the JSON array."
          "Do the job with no verbosity, don't display your comments. ")


# --- GOOGLE API initialization
try:
    # Read Gemini API-Key credentials stored in secrets.toml file
    apikey = st.secrets.google_api["apikey"]
    if not apikey:
        raise ValueError("La variable 'GEMINI_API_KEY' no está configurada.")
    client = genai.Client(api_key=apikey)
except (ValueError, Exception) as e:
    st.error(f"Error de configuración: {e}")
    exit()

# LLMWhisperer Definition
base_url = st.secrets.llmwhisperer_api["base_url"]
api_key = st.secrets.llmwhisperer_api["api_key"]

# ---- STREAMLIT INÍCIO DA INTERFACE E DO FLUXO  ----
st.set_page_config(layout="wide")
state = st.session_state
version_number = '2512.01'
st.sidebar.text(f'[ver. {version_number}]')
st.header(f"✨ Lector de archivos con Google IA")

# Initialize session state variables
if 'file' not in state:
   state.file = False

# Definição do modelo de IA a ser utilizado
model_name = "gemini-2.5-flash"

# Upload do arquivo
uploaded_file = st.sidebar.file_uploader("Seleccione el archivo ", type=["xlsx", "pdf", "txt"])

if uploaded_file:
    if not state.file:
        client_v2 = LLMWhispererClientV2(
            base_url=base_url,
            api_key=api_key
        )

        # Get usage info
        usage_info = client_v2.get_usage_info()
        today_page_count = usage_info['today_page_count']
        daily_quota = usage_info['daily_quota']
        remain_pages = daily_quota - today_page_count
        if today_page_count < daily_quota:
            st.info(f'Quedan {remain_pages} páginas de un total de {daily_quota} páginas gratuitas por día.')
        else:
            st.error(f'Alcanzado el límite diario de {daily_quota} páginas para procesar. ¡Inténtalo de nuevo mañana!')
            st.stop()

        # Cria um arquivo temporário.
        temp_dir = tempfile.gettempdir()
        file_extension = os.path.splitext(uploaded_file.name)[1].lower()

        # Cria um nome de arquivo temporário completo com extensão
        temp_file_path = os.path.join(temp_dir, f"temp_upload_{int(time.time())}{file_extension}")

        try:
            # Escreve o conteúdo do UploadedFile no arquivo temporário.
            with open(temp_file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            # Passa o CAMINHO do arquivo temporário para a função client.whisper().
            result = client_v2.whisper(file_path=temp_file_path)
            # st.text_area('RESULT:', result)

            with st.spinner('Step 1: Extrayendo los datos, espere...'):
                # Loop de Status (Mantenha o loop de status inalterado)
                while True:
                    status = client_v2.whisper_status(whisper_hash=result['whisper_hash'])
                    if status['status'] == 'processed':
                        resultado = client_v2.whisper_retrieve(
                            whisper_hash=result['whisper_hash']
                        )
                        # st.text_area("RESULTADO COMPLETO DO LMMWHISPERER:", resultado)
                        # st.code(resultado)
                        break
                    # st.info(f"Processando... Status atual: {status['status']}")
                    time.sleep(5)

                # Exibe o resultado
                extracted_text = resultado['extraction']['result_text']
                st.success("¡Extracción concluida!")
                # st.code(extracted_text)

            try:
                with st.spinner(f'Step 2: Usando {model_name} para extraer los datos. Por favor espere...'):
                    start_time = time.perf_counter()
                    try:
                        # get result from AI model
                        ai_result = client.models.generate_content(
                            model=model_name,
                            contents=[prompt, extracted_text])

                        # Exibe o resultado no Streamlit
                        # state.result = ai_result.text if hasattr(ai_result, "text") else str(ai_result)
                        # Segurança extra
                        if not hasattr(ai_result, "text") or not ai_result.text:
                            st.error("⚠️ La IA no retornó ningún contenido.")
                            st.stop()
                        gemini_output_text = ai_result.text

                        cleaned = gemini_output_text
                        # Remove ```json ... ```
                        cleaned = re.sub(r'^```[a-zA-Z]*\s*|\s*```$', '', cleaned.strip(), flags=re.MULTILINE)
                        # Remove caracteres de controle
                        cleaned = re.sub(r'[\x00-\x1F\x7F]', '', cleaned)
                        # remove espaços em branco onde não deve existir
                        cleaned = cleaned.replace('[ { "', '[{"')
                        cleaned = cleaned.replace(' }, { "', '}, {"')
                        cleaned = cleaned.replace('} ]', '}]')

                        cleaned = fix_json_quotes(cleaned)

                        # Se não começa como JSON, já bloqueia
                        if not (cleaned.startswith("[") or cleaned.startswith("{")):
                            st.error("⚠️ La IA no retornó un archivo JSON válido.")
                            st.text(cleaned[:1000])  # Debug
                            st.stop()

                        try:
                            data_list = json.loads(cleaned)
                        except json.JSONDecodeError as e:
                            st.error(f"❌ Error al interpretar el archivo JSON: {e}")
                            st.text(cleaned)  # DEBUG TOTAL
                            st.stop()

                        df = pd.DataFrame(data_list)
                        df.index = range(1, len(df) + 1)
                        state.df = df
                        state.file = True

                        end_time = time.perf_counter()  # ⏱️ FIM DO TIMER
                        elapsed = end_time - start_time
                        st.metric(label="⏱️ Tiempo de conversión",
                                  value=format_hms(elapsed))

                        # df_to_save = st.data_editor(
                        #     df,
                        #     key="df_edited",
                        #     num_rows="dynamic",
                        #     hide_index=False,
                        # )

                    except Exception as e:
                        st.error(f"Error inesperado: {e}")
                        st.stop()

            except APIError as e:
                st.error(f"Error en la API de Gemini.: {e}")

        finally:
            # Exclui o arquivo temporário (limpeza obrigatória!)
            if os.path.exists(temp_file_path):
                os.remove(temp_file_path)

    # Exibindo o dataframe para edição
    if 'df' in state:
        df_to_save = st.data_editor(
            state.df,
            num_rows="dynamic",
            hide_index=True,
        )
        state.df_to_save = df_to_save.copy().reset_index(drop=True)
    else:
        st.warning('No se pudo procesar el documento. Inténtelo de nuevo.')

    st.divider()
    col_info, col_btn_save = st.columns([0.6, 0.4])

    with col_btn_save:
        if st.button('✅ Guardar los datos en la memoria.', type='primary', use_container_width=True):
            if isinstance(state.df, pd.DataFrame) and not state.df.empty:
                # Armazena o DataFrame na sessão para uso posterior
                state.df_to_save.reset_index(drop=True)
                num_lines = len(state.df_to_save)
                st.success(f'{num_lines} lineas de datos guardados, ya puedes crear un nuevo documento.')
                # Limpando a session.state no final de toda operação
                # for key in state.keys():
                #     del state[key]
            else:
                st.warning('No hay datos para guardar. Primero, importe y procese un archivo.')
        else:
            st.info("No hay datos almacenados en la memoria.")

    # Limpando a session.state no final de toda operação
    # for key in state.keys():
    #     del state[key]
