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

st.write("🔄 Obteniendo token...")
app = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET,
)

result = app.acquire_token_for_client(scopes=SCOPE)
if "access_token" not in result:
    st.error(f"❌ Error al obtener token: {json.dumps(result, indent=2)}")
    st.stop()

token = result["access_token"]
headers = {"Authorization": f"Bearer {token}"}
st.success("✅ Token obtenido correctamente")

# ==========================
# 🔍 Buscar el site por nombre
# ==========================
sites_url = f"https://graph.microsoft.com/v1.0/sites?search={SITE_NAME}"
st.write(f"📌 Llamando a: `{sites_url}`")
resp = requests.get(sites_url, headers=headers)

st.write("STATUS:", resp.status_code)
st.json(resp.json())

if resp.status_code != 200 or "value" not in resp.json():
    st.error("⛔ No se pudo buscar sites. Revisa permisos en Azure.")
    st.stop()

sites = resp.json()["value"]

if not sites:
    st.error(f"⛔ No se encontró ningún site con nombre '{SITE_NAME}'")
    st.stop()

# Tomamos el primero que coincida
site = sites[0]
site_id = site["id"]
st.success(f"✅ Site encontrado: {site['name']} → {site_id}")

# ==========================
# 🔍 Listar drives del site
# ==========================
drive_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
st.write(f"📌 Llamando a: `{drive_url}`")
drives_resp = requests.get(drive_url, headers=headers)

st.write("STATUS:", drives_resp.status_code)
st.json(drives_resp.json())

if drives_resp.status_code != 200:
    st.error("⛔ No se pudo acceder a los drives.")
    st.stop()

drives = drives_resp.json().get("value", [])
drive_id = next((d["id"] for d in drives if d["name"] == LIBRARY), None)

if drive_id:
    st.success(f"✅ Drive encontrado: {LIBRARY} → {drive_id}")
else:
    st.warning(f"❌ No se encontró la biblioteca '{LIBRARY}'. Disponibles:")
    for d in drives:
        st.write("-", d["name"])



# ==========================
# 📂 Listar archivos de la biblioteca Documentos
# ==========================
items_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children"
print(f"📌 Llamando a: {items_url}")
items_resp = requests.get(items_url, headers=headers)
print("STATUS:", items_resp.status_code)
print("RESPUESTA:", json.dumps(items_resp.json(), indent=2), "\n")

if items_resp.status_code == 200:
    items = items_resp.json().get("value", [])
    if not items:
        print("⚠️ La biblioteca está vacía.")
    else:
        print("✅ Archivos y carpetas encontrados:")
        for item in items:
            tipo = "📁 Carpeta" if "folder" in item else "📄 Archivo"
            print(f"{tipo}: {item['name']} → {item['id']}")
else:
    raise SystemExit("⛔ No se pudo acceder a los items de la biblioteca.")
