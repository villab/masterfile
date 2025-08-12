from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
import streamlit as st

USERNAME = st.secrets["sharepoint_user"]
APP_PASSWORD = st.secrets["app_password"]

SITE_URL = "https://caseonit.sharepoint.com/sites/Sutel"

try:
    ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, APP_PASSWORD))
    web = ctx.web.get().execute_query()
    st.write(f"Conectado a: {web.properties['Title']}")
except Exception as e:
    st.error(f"Error de conexión: {e}")
