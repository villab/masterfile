import streamlit as st
import msal
import jwt
import requests

# ================== CONFIG ==================
TENANT_ID = st.secrets["azure_tenant_id"]
CLIENT_ID = st.secrets["azure_client_id"]
CLIENT_SECRET = st.secrets["azure_client_secret"]

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

# ================== TOKEN ==================
app = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET
)

result = app.acquire_token_for_client(scopes=SCOPE)

if "access_token" in result:
    st.success("✅ Token obtenido correctamente")

    # ================== DECODIFICAR CLAIMS ==================
    claims = jwt.decode(result["access_token"], options={"verify_signature": False})
    st.write("🔎 Claims del token:", claims)

    # Mostrar roles o scopes
    roles = claims.get("roles", [])
    scp = claims.get("scp", "")

    st.write("📌 Roles (Application permissions):", roles)
    st.write("📌 Scopes (Delegated permissions):", scp)

else:
    st.error("❌ Error al obtener token")
    st.json(result)
