# ==============================================================
# VERSION FINAL MASTER - SUTEL MASTERFILE
# Reemplazo de AgGrid por st.data_editor (Solución LargeUtf8)
# Filtros dinámicos en Sidebar y Sincronización de Cambios
# ==============================================================

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
<div style="background: linear-gradient(90deg,#0f2027,#203a43,#2c5364); padding: 18px 24px; border-radius: 12px; margin-bottom: 10px;">
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

# ========= Autenticación y Graph API =========
@st.cache_data(ttl=3000)
def get_access_token_cached():
    app = msal.ConfidentialClientApplication(CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}", client_credential=CLIENT_SECRET)
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result: raise Exception(f"Error Token: {result}")
    return result["access_token"]

@st.cache_data(ttl=3600)
def get_site_drive_cached():
    token = get_access_token_cached()
    headers = {"Authorization": f"Bearer {token}"}
    sites = requests.get(f"https://graph.microsoft.com/v1.0/sites?search={SITE_NAME}", headers=headers).json().get("value", [])
    site = next((s for s in sites if SITE_HOST in s.get("webUrl", "")), sites[0])
    drives = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site['id']}/drives", headers=headers).json().get("value", [])
    return site["id"], drives[0]["id"]

def get_file_from_sharepoint(path):
    token = get_access_token_cached()
    s_id, d_id = get_site_drive_cached()
    url = f"https://graph.microsoft.com/v1.0/sites/{s_id}/drives/{d_id}/root:/{path}:/content"
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"})
    if resp.status_code != 200: raise Exception(f"Error descarga {path}")
    return BytesIO(resp.content)

def upload_file_to_sharepoint(path, file_bytes):
    token = get_access_token_cached()
    s_id, d_id = get_site_drive_cached()
    url = f"https://graph.microsoft.com/v1.0/sites/{s_id}/drives/{d_id}/root:/{path}:/content"
    resp = requests.put(url, headers={"Authorization": f"Bearer {token}"}, data=file_bytes.getvalue())
    if resp.status_code not in (200, 201): raise Exception(f"Error subida {path}")

def ensure_folder(path):
    token = get_access_token_cached()
    s_id, d_id = get_site_drive_cached()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    parts = path.split('/')
    current_path = ""
    for part in parts:
        parent = current_path
        current_path = f"{current_path}/{part}" if current_path else part
        url = f"https://graph.microsoft.com/v1.0/sites/{s_id}/drives/{d_id}/root:/{current_path}"
        if requests.get(url, headers=headers).status_code != 200:
            c_url = f"https://graph.microsoft.com/v1.0/sites/{s_id}/drives/{d_id}/root{':/'+parent+':' if parent else ''}/children"
            requests.post(c_url, headers=headers, json={"name": part, "folder": {}})

# ========= Lógica de Correo y Contador =========
def enviar_correo_con_adjuntos(asunto, cuerpo, archivos_adjuntos):
    msg = EmailMessage()
    msg["Subject"], msg["From"], msg["To"] = asunto, EMAIL_FROM, EMAIL_TO
    msg.set_content(cuerpo)
    for f_bytes, f_name in archivos_adjuntos:
        msg.add_attachment(f_bytes.getvalue(), maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=f_name)
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
        smtp.starttls()
        smtp.login(SMTP_USER, SMTP_PASS)
        smtp.send_message(msg)

def _leer_contador_hoy():
    fecha_hoy = datetime.now(ZoneInfo("America/Costa_Rica")).strftime("%d%m%Y")
    try:
        stream = get_file_from_sharepoint(f"{FOLDER_PATH}/contador_envios.txt")
        f_guardada, cnt = stream.read().decode("utf-8").strip().split(",")
        return (fecha_hoy, int(cnt)) if f_guardada == fecha_hoy else (fecha_hoy, 0)
    except: return fecha_hoy, 0

def _guardar_contador_hoy(fecha, nuevo_cnt):
    upload_file_to_sharepoint(f"{FOLDER_PATH}/contador_envios.txt", BytesIO(f"{fecha},{nuevo_cnt}".encode("utf-8")))

# ========= Detección de Cambios =========
def normalize_val(v):
    if v is None or (isinstance(v, float) and np.isnan(v)): return ""
    return str(v).strip()

def detectar_cambios(df_orig, df_mod, tipo):
    cambios = []
    # Aseguramos que ambos tengan el mismo index por ROWKEY para comparar fila a fila correctamente
    df_o = df_orig.set_index(ROWKEY)
    df_m = df_mod.set_index(ROWKEY)
    
    comunes = df_o.index.intersection(df_m.index)
    cols = [c for c in df_o.columns if c in df_m.columns]

    for idx in comunes:
        row_o = df_o.loc[idx]
        row_m = df_m.loc[idx]
        for c in cols:
            val_o = normalize_val(row_o[c])
            val_m = normalize_val(row_m[c])
            if val_o != val_m:
                ident = f"ID {row_o[ID_COL]}" if ID_COL in row_o else f"Fila {idx}"
                if "Stm" in row_o: ident = f"Stm {row_o['Stm']}"
                cambios.append(f"{ident}: {c} de '{val_o}' → '{val_m}'")
    return cambios

# ========= Manejo de Archivo y Filtros =========
def manejar_archivo(nombre_modo, nombre_archivo):
    # 1. Carga de datos
    file_stream = get_file_from_sharepoint(f"{FOLDER_PATH}/{nombre_archivo}")
    contenido_binario = file_stream.getvalue()
    
    df = pd.read_excel(file_stream).fillna('')
    df = df.astype(object)
    df[ROWKEY] = np.arange(len(df)).astype(str)

    # --- DISEÑO SUPERIOR ---
    col_msg, col_btn = st.columns([3, 1])
    with col_msg: st.success(f"📂 {nombre_archivo} cargado.")
    with col_btn: st.download_button("Descargar Excel", data=contenido_binario, file_name=nombre_archivo, key=f"dl_{nombre_modo}")

    # --- SECCIÓN DE FILTROS DINÁMICOS ---
    with st.expander(f"🔍 Panel de Filtros Personalizados - {nombre_modo}", expanded=True):
        # Permitimos al usuario elegir qué columnas quiere usar para filtrar
        columnas_disponibles = [c for c in df.columns if c != ROWKEY]
        cols_a_filtrar = st.multiselect(
            "Selecciona las columnas por las que deseas filtrar:",
            options=columnas_disponibles,
            default=columnas_disponibles[:3] if len(columnas_disponibles) > 3 else columnas_disponibles,
            key=f"selector_cols_{nombre_modo}"
        )

        df_filtrado = df.copy()
        
        if cols_a_filtrar:
            # Creamos filas de 3 columnas para que los filtros no ocupen demasiado espacio vertical
            filas_filtros = [cols_a_filtrar[i:i + 3] for i in range(0, len(cols_a_filtrar), 3)]
            
            for fila in filas_filtros:
                st_cols = st.columns(len(fila))
                for i, col_name in enumerate(fila):
                    opciones = sorted([str(x) for x in df[col_name].unique() if str(x).strip() != ''])
                    seleccion = st_cols[i].multiselect(
                        f"Filtrar {col_name}", 
                        options=opciones, 
                        key=f"filter_{nombre_modo}_{col_name}"
                    )
                    if seleccion:
                        df_filtrado = df_filtrado[df_filtrado[col_name].astype(str).isin(seleccion)]

    st.markdown(f"**Registros encontrados:** {len(df_filtrado)}")

    # --- TABLA EDITABLE ---
    df_editado_vista = st.data_editor(
        df_filtrado,
        hide_index=True,
        column_config={ROWKEY: None},
        use_container_width=True,
        height=500,
        key=f"ed_{nombre_modo}"
    )

    # Sincronización: Actualiza el dataframe original con los cambios hechos en la vista filtrada
    if not df_editado_vista.equals(df_filtrado):
        # Usamos el ROWKEY para asegurar que el cambio vaya a la fila correcta del original
        df.set_index(ROWKEY, inplace=True)
        df.update(df_editado_vista.set_index(ROWKEY))
        df.reset_index(inplace=True)
    
    return df

    # --- TABLA EDITABLE ---
    df_editado_vista = st.data_editor(
        df_filtrado,
        hide_index=True,
        column_config={ROWKEY: None},
        use_container_width=True,
        height=500,
        key=f"ed_{nombre_modo}"
    )

    # Sincronización de cambios
    if not df_editado_vista.equals(df_filtrado):
        df.set_index(ROWKEY, inplace=True)
        df.update(df_editado_vista.set_index(ROWKEY))
        df.reset_index(inplace=True)
    
    return df
    # --- INTERFAZ ---
    col_msg, col_btn = st.columns([3, 1])
    with col_msg: st.info(f"📂 {nombre_archivo} | {len(df_filtrado)} filas filtradas.")
    with col_btn:
        st.download_button("Descargar Excel", data=contenido_binario, file_name=nombre_archivo, key=f"dl_{nombre_modo}")

    df_editado_vista = st.data_editor(
        df_filtrado,
        hide_index=True,
        column_config={ROWKEY: None},
        use_container_width=True,
        height=500,
        key=f"ed_{nombre_modo}"
    )

    # Sincronización de cambios: Actualizar el DF original con lo editado en la vista filtrada
    if not df_editado_vista.equals(df_filtrado):
        df.set_index(ROWKEY, inplace=True)
        df.update(df_editado_vista.set_index(ROWKEY))
        df.reset_index(inplace=True)
    
    return df

# ================== MAIN UI ==================
try:
    tab1, tab2 = st.tabs(["📄 Masterfile Fijo", "📄 Masterfile Movilidad"])

    with tab1:
        df_fijo_final = manejar_archivo("Fijo", ARCHIVOS["Fijo"])
    with tab2:
        df_movilidad_final = manejar_archivo("Movilidad", ARCHIVOS["Movilidad"])

    st.markdown("---")
    if st.button("💾 GUARDAR CAMBIOS Y ENVIAR CORREO", use_container_width=True):
        with st.spinner("Procesando cambios y subiendo a SharePoint..."):
            timestamp = datetime.now(ZoneInfo("America/Costa_Rica")).strftime("%Y%m%d_%H%M%S")
            adjuntos = []
            cuerpo = f"Reporte de cambios - {timestamp}\n\n"

            for modo, df_mod, n_arc in [("Fijo", df_fijo_final, ARCHIVOS["Fijo"]), ("Movilidad", df_movilidad_final, ARCHIVOS["Movilidad"])]:
                # Obtener original puro para comparar cambios reales
                df_orig = pd.read_excel(get_file_from_sharepoint(f"{FOLDER_PATH}/{n_arc}")).fillna('')
                df_orig[ROWKEY] = np.arange(len(df_orig)).astype(str)

                lista_cambios = detectar_cambios(df_orig, df_mod, modo)
                cuerpo += f"📌 ENTORNO {modo.upper()}:\n"
                cuerpo += ("\n".join([f"• {c}" for c in lista_cambios]) if lista_cambios else "Sin cambios detectados.") + "\n\n"

                # Guardar en Excel
                df_save = df_mod.drop(columns=[ROWKEY], errors='ignore')
                buf = BytesIO()
                df_save.to_excel(buf, index=False)
                buf.seek(0)

                # Backups y Sobrescribir
                bkp_path = f"{FOLDER_PATH}/Backups/{modo}/{n_arc.replace('.xlsx','')}_{timestamp}.xlsx"
                ensure_folder(f"{FOLDER_PATH}/Backups/{modo}")
                upload_file_to_sharepoint(bkp_path, buf)
                buf.seek(0)
                upload_file_to_sharepoint(f"{FOLDER_PATH}/{n_arc}", buf)
                buf.seek(0)
                adjuntos.append((BytesIO(buf.getvalue()), f"{n_arc.replace('.xlsx','')}_{timestamp}.xlsx"))

            # Notificación Correo
            f_hoy, c_act = _leer_contador_hoy()
            asunto = f"Masterfile Sutel {f_hoy}" + (f" V{c_act+1}" if c_act > 0 else "")
            enviar_correo_con_adjuntos(asunto, cuerpo + "\nSaludos.", adjuntos)
            _guardar_contador_hoy(f_hoy, c_act + 1)
            
            st.success("✅ Guardado exitoso. Archivos actualizados y correo enviado.")
            st.balloons()

except Exception as e:
    st.error(f"❌ Error en la aplicación: {e}")
