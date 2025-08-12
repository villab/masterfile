import streamlit as st
import pandas as pd
from io import BytesIO
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

# ================== CONFIGURACIÓN ==================
USERNAME = st.secrets["sharepoint_user"]  # Tu usuario de Office 365
PASSWORD = st.secrets["sharepoint_pass"]  # Tu contraseña o App Password
SITE_URL = "https://caseonit.sharepoint.com/sites/Sutel"

# Ruta exacta del archivo en SharePoint (server-relative URL)
FILE_URL = "/sites/Sutel/Documentos compartidos/01. Documentos MedUX/Automatizacion/MasterfileSutel.xlsx"

# ================== CONEXIÓN A SHAREPOINT ==================
try:
    ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, PASSWORD))
    file = ctx.web.get_file_by_server_relative_url(FILE_URL)
    file_stream = BytesIO()
    file.download(file_stream).execute_query()
    file_stream.seek(0)

    # ================== LECTURA DEL EXCEL ==================
    df = pd.read_excel(file_stream)
    st.success("Archivo cargado correctamente desde SharePoint ✅")
    st.dataframe(df)

except Exception as e:
    st.error(f"Error al descargar el archivo: {e}")
