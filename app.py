import streamlit as st
import msal
import requests

# Leer de secrets
CLIENT_ID = st.secrets["client_id"]
TENANT_ID = st.secrets["tenant_id"]
CLIENT_SECRET = st.secrets["client_secret"]

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

# Autenticación MSAL
app = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET
)

result = app.acquire_token_for_client(scopes=SCOPE)

if "access_token" in result:
    st.success("✅ Autenticación exitosa con Graph API")
    token = result["access_token"]

    headers = {"Authorization": f"Bearer {token}"}
    url = "https://graph.microsoft.com/v1.0/sites?search=caseonit.sharepoint.com"
    resp = requests.get(url, headers=headers)

    st.write("Respuesta Graph:", resp.status_code)
    st.json(resp.json())
else:
    st.error(f"❌ Error al autenticar: {result.get('error')} {result.get('error_description')}")
