import streamlit as st
import requests
import msal

# ================== CONFIG ==================
TENANT_ID = st.secrets["tenant_id"]
CLIENT_ID = st.secrets["client_id"]
CLIENT_SECRET = st.secrets["client_secret"]

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

# ================== TOKEN ==================
app = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET
)

result = app.acquire_token_silent(SCOPE, account=None)

if not result:
    result = app.acquire_token_for_client(scopes=SCOPE)

if "access_token" not in result:
    st.error("❌ Error al obtener token")
    st.json(result)
else:
    access_token = result["access_token"]
    headers = {"Authorization": f"Bearer {access_token}"}
    st.success("✅ Token obtenido correctamente")

    # ================== BUSCAR EL SITE ==================
    site_url = "https://graph.microsoft.com/v1.0/sites/caseonit.sharepoint.com:/sites/Sutel"
    resp_site = requests.get(site_url, headers=headers)

    if resp_site.status_code == 200:
        site_data = resp_site.json()
        st.success("✅ Site encontrado")
        st.json(site_data)
    else:
        st.error(f"❌ Error al buscar el site (status {resp_site.status_code})")
        st.json(resp_site.json())
