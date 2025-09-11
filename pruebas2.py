import requests
import msal
import streamlit as st

# ==========================
# 🔑 Credenciales de Azure
# ==========================
TENANT_ID = st.secrets["tenant_id"]
CLIENT_ID = st.secrets["client_id"]
CLIENT_SECRET = st.secrets["client_secret"]

# ==========================
# 📌 Configuración SharePoint
# ==========================
SITE_HOST = "caseonit.sharepoint.com"
SITE_NAME = "Sutel"       # 👈 pon aquí el nombre exacto del site
LIBRARY = "Documentos"    # 👈 normalmente "Documentos" en español

# ==========================
# 🎟️ Obtener token
# ==========================
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

app = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET,
)

result = app.acquire_token_for_client(scopes=SCOPE)
if "access_token" not in result:
    raise Exception(f"❌ Error al obtener token: {result}")

token = result["access_token"]
headers = {"Authorization": f"Bearer {token}"}

print("✅ Token obtenido correctamente\n")

# ==========================
# 🔍 Diagnóstico: Site
# ==========================
site_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_HOST}:/sites/{SITE_NAME}"
resp = requests.get(site_url, headers=headers)
print("📌 Verificando acceso al site...")
print("STATUS:", resp.status_code)
print(resp.json(), "\n")

if resp.status_code != 200:
    raise SystemExit("⛔ No se pudo acceder al site, revisa permisos en Azure.")

site_id = resp.json().get("id")
print(f"✅ Site ID: {site_id}\n")

# ==========================
# 🔍 Diagnóstico: Drives
# ==========================
drive_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
drives_resp = requests.get(drive_url, headers=headers)
print("📌 Verificando drives...")
print("STATUS:", drives_resp.status_code)
print(drives_resp.json(), "\n")

if drives_resp.status_code != 200:
    raise SystemExit("⛔ No se pudo acceder a los drives.")

# Buscar la biblioteca configurada
drives = drives_resp.json().get("value", [])
drive_id = next((d["id"] for d in drives if d["name"] == LIBRARY), None)

if drive_id:
    print(f"✅ Drive encontrado: {LIBRARY} → {drive_id}")
else:
    print(f"❌ No se encontró la biblioteca '{LIBRARY}'. Disponibles:")
    for d in drives:
        print("-", d["name"])
