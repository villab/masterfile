import msal
import requests
import jwt
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode, JsCode
from zoneinfo import ZoneInfo
import smtplib
from email.message import EmailMessage
import requests
import msal

# ⚙️ Configuración
TENANT_ID = st.secrets["tenant_id"]
CLIENT_ID = st.secrets["client_id"]
CLIENT_SECRET = st.secrets["client_secret"]

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]

# 🔑 Crear app confidencial
app = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET
)

# 📥 Obtener token
result = app.acquire_token_for_client(scopes=SCOPES)

if "access_token" in result:
    print("✅ Token obtenido")
    token = result["access_token"]

    # Decodificar cabecera del token para verificar roles
    decoded = jwt.decode(token, options={"verify_signature": False})
    print("Roles en el token:", decoded.get("roles", []))
else:
    print("❌ Error al obtener token:", result.get("error_description"))

