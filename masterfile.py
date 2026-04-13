# VERSION FINAL - MASTERFILE con MSAL y Microsoft Graph API
# ==============================================
# - Sin dependencias de AgGrid (Solución error LargeUtf8)
# - Uso de st.data_editor nativo
# - Backups separados para FIJO y MOVILIDAD
# - Envío de correo con contador persistente
# ==============================================

import streamlit as st
import pandas as pd
import numpy as np
import time
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo
import smtplib
from email.message import EmailMessage
import requests
import msal
from config import get_secret

# ------ Configuración de vista ----------
st.set_page_config(
    page_title="Masterfile Sutel",
    page_icon="📋",
    layout="wide"
)

st.markdown("""
<div style="
    background: linear-gradient(90deg,#0f2027,#203a43,#2c5364);
    padding: 18px 24px;
    border-radius: 12px;
    margin-bottom: 10px;
">
    <h2 style="margin:0; color:#E5E7EB;">📋 Masterfile Entorno de Medición</h2>
    <p style="margin:0;font-size:13px;opacity:0.8;">Gestión de archivos Fijo y Movilidad</p>
</div>
""", unsafe_allow_html=True)

# ================== CONFIGURACIÓN ==================
TENANT_ID = get_secret("tenant_id")
CLIENT_ID = get_secret("client_id")
CLIENT_SECRET = get_secret("client_secret")

SITE_HOST = "caseonit.sharepoint.com"
SITE_NAME = "Sutel"
FOLDER_PATH = "01. Documentos MedUX/Automatizacion/Masterfile"

ARCHIVOS = {
    "Fijo": "MasterfileSutel.xlsx",
    "Movilidad": "MasterfileSutel_Movilidad.xlsx"
}

SMTP_SERVER = get_secret("smtp_server")
SMTP_PORT = get_secret("smtp_port")
SMTP_USER = get_secret("smtp_user")
SMTP_PASS = get_secret("smtp_pass")
EMAIL_FROM = get_secret("email_from")
EMAIL_TO = get_secret("email_to")

ID_COL = "ID SONDA"
ROWKEY = "_row_id"

# ========= Autenticación con MSAL =========
@st.cache_data(ttl=3000)
def get_access_token_cached():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise Exception(f"No se pudo obtener token: {result}")
    return result["access_token"]

# ========= Funciones SharePoint con Graph =========
@st.cache_data(ttl=3600)
def get_site_drive_cached():
    token = get_access_token_cached()
    headers = {"Authorization": f"Bearer {token}"}
    search_url = f"https://graph.microsoft.com/v1.0/sites?search={SITE_NAME}"
    resp = requests.get(search_url, headers=headers)
    sites = resp.json().get("value", [])
    site = next((s for s in sites if SITE_HOST in s.get("webUrl", "")), sites[0])
    site_id = site["id"]
    drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    drives_resp = requests.get(drives_url, headers=headers)
    drive = drives_resp.json().get("value", [])[0]
    return site_id, drive["id"]

def get_file_from_sharepoint(path):
    token = get_access_token_cached()
    site_id, drive_id = get_site_drive_cached()
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{path}:/content"
    resp = requests.get(url, headers=headers)
    return BytesIO(resp.content)

def upload_file_to_sharepoint(path, file_bytes, max_retries=5):
    token = get_access_token_cached()
    site_id, drive_id = get_site_drive_cached()
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{path}:/content"
    for intento in range(max_retries):
        resp = requests.put(url, headers=headers, data=file_bytes.getvalue(), timeout=300)
        if resp.status_code in (200, 201): return
        time.sleep(2 ** intento)
    raise Exception(f"Error subiendo archivo {path}")

def ensure_folder(path):
    token = get_access_token_cached()
    site_id, drive_id = get_site_drive_cached()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{path}"
    if requests.get(url, headers=headers).status_code == 200: return
    parent, folder_name = "/".join(path.split("/")[:-1]), path.split("/")[-1]
    create_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root{':/'+parent+':' if parent else ''}/children"
    requests.post(create_url, headers=headers, json={"name": folder_name, "folder": {}})

# ========= Gestión de Correo y Contador =========
def enviar_correo_con_adjuntos(asunto, cuerpo, archivos_adjuntos):
    msg = EmailMessage()
    msg["Subject"], msg["From"], msg["To"] = asunto, EMAIL_FROM, EMAIL_TO
    msg.set_content(cuerpo)
    for archivo_bytes, nombre_archivo in archivos_adjuntos:
        msg.add_attachment(archivo_bytes.getvalue(), maintype="application", 
                           subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=nombre_archivo)
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
        smtp.starttls()
        smtp.login(SMTP_USER, SMTP_PASS)
        smtp.send_message(msg)

def _leer_contador_hoy():
    fecha_hoy = datetime.now(ZoneInfo("America/Costa_Rica")).strftime("%d%m%Y")
    try:
        stream = get_file_from_sharepoint(f"{FOLDER_PATH}/contador_envios.txt")
        fecha_guardada, cnt = stream.read().decode("utf-8").strip().split(",")
        return (fecha_hoy, int(cnt)) if fecha_guardada == fecha_hoy else (fecha_hoy, 0)
    except: return fecha_hoy, 0

def _guardar_contador_hoy(fecha_ddmmaaaa, nuevo_contador):
    upload_file_to_sharepoint(f"{FOLDER_PATH}/contador_envios.txt", BytesIO(f"{fecha_ddmmaaaa},{nuevo_contador}".encode("utf-8")))

# ========= Lógica de Comparación =========
def normalize_df_for_compare(df):
    if df is None or df.empty: return df
    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]
    for c in out.columns:
        out[c] = out[c].astype(str).str.strip().replace(['nan', 'None', 'NaT'], '')
    return out

def detectar_cambios(df_original, df_modificado, tipo_archivo):
    do = normalize_df_for_compare(df_original.drop(columns=[ROWKEY], errors='ignore'))
    dm = normalize_df_for_compare(df_modificado.drop(columns=[ROWKEY], errors='ignore'))
    cambios = []
    for i in range(min(len(do), len(dm))):
        for col in do.columns:
            if col in dm.columns and do.iloc[i][col] != dm.iloc[i][col]:
                ident = f"Fila {i+1}"
                if "Stm" in do.columns: ident = f"Stm {do.iloc[i]['Stm']}"
                cambios.append(f"{ident}: {col} de {do.iloc[i][col]} → {dm.iloc[i][col]}")
    return cambios

# ========= Manejo de interfaz de archivo =========
def manejar_archivo(nombre_modo, nombre_archivo):
    file_stream = get_file_from_sharepoint(f"{FOLDER_PATH}/{nombre_archivo}")
    contenido_binario = file_stream.getvalue()
    file_stream.seek(0)
    
    # Lectura y Limpieza Radical para evitar errores de serialización
    df = pd.read_excel(file_stream).fillna('')
    df = df.astype(object) # Evita tipos Arrow complejos
    df[ROWKEY] = np.arange(len(df)).astype(str)

    col_msg, col_btn = st.columns([3, 1])
    with col_msg: st.success(f"📂 Cargado {nombre_archivo} ✅")
    with col_btn:
        st.download_button("Descargar última versión", data=contenido_binario, 
                           file_name=nombre_archivo, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                           key=f"dl_{nombre_modo}", use_container_width=True)
    
    # --- NUEVO: BUSCADOR / FILTRO DINÁMICO ---
    busqueda = st.text_input(f"🔍 Filtrar en {nombre_modo} (Escribe ID, Nombre o cualquier valor):", key=f"search_{nombre_modo}")
    
    if busqueda:
        # Filtra el dataframe si cualquier celda contiene el texto buscado
        mask = df.apply(lambda row: row.astype(str).str.contains(busqueda, case=False).any(), axis=1)
        df_mostrar = df[mask]
    else:
        df_mostrar = df

    # --- TABLA EDITABLE ---
    df_editado = st.data_editor(
        df_mostrar,
        hide_index=True,
        column_config={ROWKEY: None}, 
        use_container_width=True,
        height=500,
        key=f"editor_{nombre_modo}"
    )
    
    # Importante: Si filtramos, debemos devolver el dataframe completo con los cambios aplicados
    # para no perder las filas que no están visibles.
    if busqueda:
        df.update(df_editado)
        return df
        
    return df_editado



# ================== INTERFAZ PRINCIPAL ==================
try:
    tab_fijo, tab_movilidad = st.tabs(["📄 Masterfile Fijo", "📄 Masterfile Movilidad"])

    with tab_fijo:
        df_fijo_final = manejar_archivo("Fijo", ARCHIVOS["Fijo"])
    with tab_movilidad:
        df_movilidad_final = manejar_archivo("Movilidad", ARCHIVOS["Movilidad"])

    if st.button("💾 Guardar nueva versión de Masterfile", use_container_width=True):
        timestamp = datetime.now(ZoneInfo("America/Costa_Rica")).strftime("%Y%m%d_%H%M%S")
        archivos_adjuntos, cuerpo_correo = [], f"Cambios realizados el {timestamp}:\n\n"

        for modo, df_mod, n_arch in [("Fijo", df_fijo_final, ARCHIVOS["Fijo"]), ("Movilidad", df_movilidad_final, ARCHIVOS["Movilidad"])]:
            # Obtener original para comparar
            df_orig = pd.read_excel(get_file_from_sharepoint(f"{FOLDER_PATH}/{n_arch}")).fillna('')
            df_orig[ROWKEY] = np.arange(len(df_orig)).astype(str)
            
            cambios = detectar_cambios(df_orig, df_mod, modo)
            cuerpo_correo += f"📌 {modo}: {chr(10).join(['• '+c for c in cambios]) if cambios else 'Sin cambios'}\n\n"

            # Guardar Excel
            df_save = df_mod.drop(columns=[ROWKEY], errors='ignore')
            buf = BytesIO()
            df_save.to_excel(buf, index=False)
            buf.seek(0)
            
            # Subir a SharePoint (Backup y Principal)
            path_bkp = f"{FOLDER_PATH}/Backups/{modo}/{n_arch.replace('.xlsx','')}_{timestamp}.xlsx"
            ensure_folder(f"{FOLDER_PATH}/Backups/{modo}")
            upload_file_to_sharepoint(path_bkp, buf)
            buf.seek(0)
            upload_file_to_sharepoint(f"{FOLDER_PATH}/{n_arch}", buf)
            buf.seek(0)
            archivos_adjuntos.append((BytesIO(buf.getvalue()), f"{n_arch.replace('.xlsx','')}_{timestamp}.xlsx"))

        # Correo y Contador
        f_hoy, c_actual = _leer_contador_hoy()
        asunto = f"Masterfile Sutel {f_hoy}" + (f" V{c_actual+1}" if c_actual > 0 else "")
        enviar_correo_con_adjuntos(asunto, cuerpo_correo + "Un saludo", archivos_adjuntos)
        _guardar_contador_hoy(f_hoy, c_actual + 1)
        st.success("📧 Masterfile guardado y notificado por correo.")

except Exception as e:
    st.error(f"Error crítico: {e}")
