import requests
import msal
import streamlit as st
import json

# ==========================
# 🔑 Credenciales de Azure
# ==========================
TENANT_ID = st.secrets["tenant_id"]
CLIENT_ID = st.secrets["client_id"]
CLIENT_SECRET = st.secrets["client_secret"]

# ==========================
# 📌 Configuración SharePoint
# ==========================
SITE_NAME = "Sutel"       # 👈 nombre del site en SharePoint
LIBRARY = "Documentos"    # 👈 normalmente "Documentos" o "Shared Documents"

# ==========================
# 🎟️ Obtener token
# ==========================
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

print("🔄 Obteniendo token...")
app = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET,
)

result = app.acquire_token_for_client(scopes=SCOPE)
if "access_token" not in result:
    raise Exception(f"❌ Error al obtener token: {json.dumps(result, indent=2)}")

token = result["access_token"]
headers = {"Authorization": f"Bearer {token}"}

print("✅ Token obtenido correctamente\n")

# ==========================
# 🔍 Buscar el site por nombre
# ==========================
sites_url = f"https://graph.microsoft.com/v1.0/sites?search={SITE_NAME}"
print(f"📌 Llamando a: {sites_url}")
resp = requests.get(sites_url, headers=headers)
print("STATUS:", resp.status_code)
print("RESPUESTA:", json.dumps(resp.json(), indent=2), "\n")

if resp.status_code != 200 or "value" not in resp.json():
    raise SystemExit("⛔ No se pudo buscar sites. Revisa permisos en Azure.")

sites = resp.json()["value"]

if not sites:
    raise SystemExit(f"⛔ No se encontró ningún site con nombre '{SITE_NAME}'")

# Tomamos el primero que coincida
site = sites[0]
site_id = site["id"]
print(f"✅ Site encontrado: {site['name']} → {site_id}\n")

# ==========================
# 🔍 Listar drives del site
# ==========================
drive_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
print(f"📌 Llamando a: {drive_url}")
drives_resp = requests.get(drive_url, headers=headers)
print("STATUS:", drives_resp.status_code)
print("RESPUESTA:", json.dumps(drives_resp.json(), indent=2), "\n")

if drives_resp.status_code != 200:
    raise SystemExit("⛔ No se pudo acceder a los drives.")

drives = drives_resp.json().get("value", [])
drive_id = next((d["id"] for d in drives if d["name"] == LIBRARY), None)

if drive_id:
    print(f"✅ Drive encontrado: {LIBRARY} → {drive_id}")
else:
    print(f"❌ No se encontró la biblioteca '{LIBRARY}'. Disponibles:")
    for d in drives:
        print("-", d["name"])
