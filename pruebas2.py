import requests
import streamlit as st
import msal

# 🔑 Configuración desde secrets de Streamlit
TENANT_ID = st.secrets["tenant_id"]
CLIENT_ID = st.secrets["client_id"]
CLIENT_SECRET = st.secrets["client_secret"]

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

# =========================
# 1. Obtener token
# =========================
app = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET
)
result = app.acquire_token_for_client(scopes=SCOPE)

if "access_token" not in result:
    st.error("❌ No se pudo obtener token")
    st.stop()

headers = {"Authorization": f"Bearer {result['access_token']}"}

# =========================
# 2. Resolver el site "Sutel"
# =========================
url_site = "https://graph.microsoft.com/v1.0/sites/caseonit.sharepoint.com:/sites/Sutel"
resp_site = requests.get(url_site, headers=headers)

if resp_site.status_code != 200:
    st.error(f"❌ Error al buscar el site: {resp_site.status_code} {resp_site.text}")
    st.stop()

site_id = resp_site.json()["id"]
st.write("📌 Site ID:", site_id)

# =========================
# 3. Obtener document libraries (drives)
# =========================
url_drives = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
resp_drives = requests.get(url_drives, headers=headers)

if resp_drives.status_code != 200:
    st.error(f"❌ Error al listar drives: {resp_drives.status_code} {resp_drives.text}")
    st.stop()

drives = resp_drives.json().get("value", [])
drive_id = None
for d in drives:
    if "Documentos compartidos" in d["name"] or "Documents" in d["name"]:
        drive_id = d["id"]
        st.write("📂 Drive encontrado:", d["name"], "➡️ ID:", drive_id)

if not drive_id:
    st.error("❌ No se encontró la document library 'Documentos compartidos'")
    st.stop()

# =========================
# 4. Listar archivos de carpeta Masterfile
# =========================
url_masterfile = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/01. Documentos MedUX/Automatizacion/Masterfile:/children"
resp_files = requests.get(url_masterfile, headers=headers)

if resp_files.status_code != 200:
    st.error(f"❌ Error al listar archivos de Masterfile: {resp_files.status_code} {resp_files.text}")
    st.stop()

files = resp_files.json().get("value", [])
st.write("📑 Archivos en carpeta Masterfile:")
for f in files:
    st.write(f"- {f['name']} ({f['webUrl']})")
