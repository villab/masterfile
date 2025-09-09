import msal
import requests
from io import BytesIO
import pandas as pd

# ============ CONFIG ============ #
CLIENT_ID = "04f0c124-f2bc-4f59-9a21-0803cd61d7e8"  # App pública de Microsoft (Office Desktop)
AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = ["Files.ReadWrite.All"]

SITE_URL = "https://caseonit.sharepoint.com/sites/Sutel"
SITE_PATH = "/Documentos compartidos/01. Documentos MedUX/Automatizacion/Masterfile/MasterfileSutel.xlsx"

# ============ AUTENTICACIÓN ============ #
app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)

# Intenta reusar sesión
accounts = app.get_accounts()
if accounts:
    result = app.acquire_token_silent(SCOPES, account=accounts[0])
else:
    result = None

# Si no hay sesión guardada, abre navegador
if not result:
    result = app.acquire_token_interactive(scopes=SCOPES)

if "access_token" not in result:
    raise Exception("❌ Error al obtener token:", result.get("error_description"))

token = result["access_token"]
headers = {"Authorization": f"Bearer {token}"}

# ============ DESCARGAR ARCHIVO ============ #
url = f"{SITE_URL}/_api/v2.0/drives/me/root:{SITE_PATH}:/content"
resp = requests.get(url, headers=headers)
resp.raise_for_status()

# Leer Excel en memoria
excel_bytes = BytesIO(resp.content)
df = pd.read_excel(excel_bytes)

print("✅ Archivo descargado con éxito")
print(df.head())
