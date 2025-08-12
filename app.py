import streamlit as st
import pandas as pd
from io import BytesIO
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

# -------------------- CONFIG --------------------
USERNAME = st.secrets["sharepoint_user"]        # usuario@dominio.com
APP_PASSWORD = st.secrets["app_password"]       # contraseña de aplicación de 16 caracteres
SITE_URL = "https://caseonit.sharepoint.com/sites/Sutel"
FOLDER_URL = "/sites/Sutel/Activos del sitio"
TARGET_FILE = "Masterfile Sutel_28_7_2025.xlsx"

# -------------------- CONEXIÓN --------------------
try:
    ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, APP_PASSWORD))
    web = ctx.web.get().execute_query()
    st.write(f"Conectado a: {web.properties['Title']}")
except Exception as e:
    st.error(f"Error de conexión: {e}")

# -------------------- DESCARGA Y LECTURA --------------------
try:
    file_url = f"{FOLDER_URL}/{TARGET_FILE}"
    file = ctx.web.get_file_by_server_relative_url(file_url).download(BytesIO()).execute_query()
    file_object = file._io  # BytesIO con el archivo
    file_object.seek(0)

    # Leer en pandas
    df = pd.read_excel(file_object)

    st.success(f"Archivo '{TARGET_FILE}' descargado correctamente")
    st.dataframe(df)

except Exception as e:
    st.error(f"Error al descargar {TARGET_FILE}: {e}")

