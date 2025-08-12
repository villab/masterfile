import streamlit as st
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

USERNAME = st.secrets["sharepoint_user"]
APP_PASSWORD = st.secrets["app_password"]
SITE_URL = "https://caseonit.sharepoint.com/sites/Sutel"

try:
    ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, APP_PASSWORD))
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()
    st.success(f"✅ Conectado correctamente a: {web.properties['Title']}")
except Exception as e:
    st.error(f"❌ Error de conexión: {e}")
