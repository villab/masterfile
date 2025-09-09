import streamlit as st
import requests
from msal import ConfidentialClientApplication

# ========= CONFIG (ajusta con tus datos) =========
TENANT_ID = st.secrets["tenant_id"]
CLIENT_ID = st.secrets["client_id"]
CLIENT_SECRET = st.secrets["client_secret"]

# URL Graph
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

# 🔑 Obtener token
def get_access_token():
    app = ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_silent(SCOPE, account=None)
    if not result:
        result = app.acquire_token_for_client(scopes=SCOPE)
    return result["access_token"]

st.title("🔍 Debug SharePoint con Graph API")

try:
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}

    # 1️⃣ Listar sites
    url_sites = "https://graph.microsoft.com/v1.0/sites?search=*"
    sites = requests.get(url_sites, headers=headers).json()
    st.subheader("📂 Sites disponibles")
    st.json(sites)

    # Si ya sabes el nombre de tu site, como "Sutel", lo buscamos
    SITE_NAME = "Sutel"
    site_match = None
    for s in sites.get("value", []):
        if SITE_NAME.lower() in s.get("name", "").lower():
            site_match = s
            break

    if site_match:
        st.success(f"✅ Encontrado site '{SITE_NAME}': {site_match['id']}")
        SITE_ID = site_match["id"]

        # 2️⃣ Listar drives de ese site
        url_drives = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives"
        drives = requests.get(url_drives, headers=headers).json()
        st.subheader("📂 Drives disponibles en el site")
        st.json(drives)

    else:
        st.error(f"No se encontró ningún site con nombre parecido a '{SITE_NAME}'")

except Exception as e:
    st.error(f"Error: {e}")
