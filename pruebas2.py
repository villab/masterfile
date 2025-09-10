import streamlit as st
import msal
import requests

# ================== CONFIG ==================
CLIENT_ID = st.secrets["client_id"]
CLIENT_SECRET = st.secrets["client_secret"]
TENANT_ID = st.secrets["tenant_id"]

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

SITE_HOSTNAME = "caseonit.sharepoint.com"  # tu dominio de SharePoint
SITE_NAME = "Sutel"  # nombre del sitio a buscar

# ================== AUTENTICACIÓN ==================
app = msal.ConfidentialClientApplication(
    CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
)
result = app.acquire_token_for_client(scopes=SCOPE)

if "access_token" in result:
    access_token = result["access_token"]

    # 1. Obtener el siteId de "Sutel"
    url_site = f"https://graph.microsoft.com/v1.0/sites/{SITE_HOSTNAME}:/sites/{SITE_NAME}"
    headers = {"Authorization": f"Bearer {access_token}"}
    resp_site = requests.get(url_site, headers=headers)

    if resp_site.status_code == 200:
        site_info = resp_site.json()
        st.write("✅ Site encontrado:", site_info)
        site_id = site_info["id"]

        # 2. Listar document libraries de ese site
        url_drives = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
        resp_drives = requests.get(url_drives, headers=headers)

        if resp_drives.status_code == 200:
            st.write("📂 Document Libraries disponibles:")
            st.json(resp_drives.json())
        else:
            st.error("❌ Error al listar document libraries")
            st.json(resp_drives.json())
    else:
        st.error("❌ Error al buscar el site")
        st.json(resp_site.json())

else:
    st.error("❌ Error en la autenticación")
    st.json(result)
