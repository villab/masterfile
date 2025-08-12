import streamlit as st
import pandas as pd
from io import BytesIO
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

# -------------------- CONFIG --------------------
USERNAME = st.secrets["sharepoint_user"]
APP_PASSWORD = st.secrets["app_password"]
SITE_URL = "https://caseonit.sharepoint.com/sites/Sutel"

# -------------------- CONEXIÓN --------------------
try:
    ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, APP_PASSWORD))
    web = ctx.web.get().execute_query()
    st.success(f"Conectado a: {web.properties['Title']}")
except Exception as e:
    st.error(f"Error de conexión: {e}")
    st.stop()

# -------------------- FUNCIONES --------------------
def descargar_excel_sharepoint(file_url):
    """
    Descarga un archivo Excel desde SharePoint y lo retorna como DataFrame de pandas.
    """
    try:
        file_obj = BytesIO()
        ctx.web.get_file_by_server_relative_url(file_url).download(file_obj).execute_query()
        file_obj.seek(0)
        return pd.read_excel(file_obj)
    except Exception as e:
        st.error(f"Error al descargar el archivo {file_url}: {e}")
        return None

# -------------------- LÓGICA PRINCIPAL --------------------
# Ejemplo: ruta relativa del archivo
file_url = "/sites/Sutel/Activos del sitio/Masterfile Sutel_28_7_2025.xlsx"

df = descargar_excel_sharepoint(file_url)

if df is not None:
    st.dataframe(df.head())
    # Aquí puedes seguir con toda la lógica de tu script grande
    # Ejemplo: filtrados, métricas, transformaciones, etc.
else:
    st.warning("No se pudo cargar el archivo desde SharePoint.")
