import streamlit as st
import msal

st.title("🔑 Prueba de autenticación Azure AD + Graph")

CLIENT_ID = st.secrets["CLIENT_ID"]
CLIENT_SECRET = st.secrets["CLIENT_SECRET"]
TENANT_ID = st.secrets["TENANT_ID"]

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

app = msal.ConfidentialClientApplication(
    CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
)

result = app.acquire_token_for_client(scopes=SCOPE)

if "access_token" in result:
    st.success("✅ Autenticación exitosa, se obtuvo el token")
    st.json(result)  # 👈 esto sí se ve en pantalla
else:
    st.error("❌ Error en la autenticación")
    st.json(result)
