import msal
import requests
from io import BytesIO
import pandas as pd
import streamlit as st

# ============ CONFIG ============ #
CLIENT_ID = "04f0c124-f2bc-4f59-9a21-0803cd61d7e8"  # App pública de Microsoft
AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = ["Files.ReadWrite.All"]

SITE_URL = "https://caseonit.sharepoint.com/sites/Sutel"
SITE_PATH = "/Documentos compartidos/01. Documentos MedUX/Automatizacion/Masterfile/MasterfileSutel.xlsx"

# ============ AUTENTICACIÓN ============ #
app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)

# Intenta reusar sesión
accounts = app.get_accounts()
result = None
if accounts:
    result = app.acquire_token_silent(SCOPES, account=accounts[0])

if not result:
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        st.error("❌ Error al iniciar flujo de autenticación (device flow).")
        st.stop()
    else:
        st.write("🔑 Ve a [https://microsoft.com/devicelogin](https://microsoft.com/devicelogin) e ingresa este código:")
        st.code(flow["user_code"])
        result = app.acquire_token_by_device_flow(flow)

# Validar si realmente obtuvimos un token
if not result or "access_token" not in result:
    st.error(f"❌ No se pudo autenticar. Detalle: {result}")
    st.stop()

# ============ SI TENEMOS TOKEN, PROBAR DESCARGA ============ #
token = result["access_token"]
headers = {"Authorization": f"Bearer {token}"}

url = f"{SITE_URL}/_api/v2.0/drives/me/root:{SITE_PATH}:/content"
resp = requests.get(url, headers=headers)

if resp.status_code == 200:
    excel_bytes = BytesIO(resp.content)
    df = pd.read_excel(excel_bytes)
    st.success("✅ Archivo descargado con éxito")
    st.dataframe(df.head())
else:
    st.error(f"❌ Error al descargar archivo: {resp.status_code} {resp.text}")
