# VERSION FINAL - MASTERFILE con MSAL y Microsoft Graph API
# ==============================================
# - Backups separados para FIJO y MOVILIDAD
# - Envío de correo con contador persistente por día
# - Detección de cambios valor viejo → valor nuevo
# ==============================================

import streamlit as st
import pandas as pd
import numpy as np
import time
from io import BytesIO
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode, JsCode
from zoneinfo import ZoneInfo
import smtplib
from email.message import EmailMessage
import requests
import msal
from config import get_secret
from st_aggrid import JsCode



# ------ Configuración de vista ----------
# ------ Configuración de vista ----------
st.set_page_config(
    page_title="Masterfile Sutel", # Nombre
    page_icon="📋",                # 
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
    <p style="margin:0;font-size:13px;opacity:0.8;">
        Gestión de archivos Fijo y Movilidad
    </p>
</div>
""", unsafe_allow_html=True)





# ================== CONFIGURACIÓN ==================
TENANT_ID = get_secret("tenant_id")
CLIENT_ID = get_secret("client_id")
CLIENT_SECRET = get_secret("client_secret")

# Ajusta si tu "library" real tiene otro nombre; la búsqueda es tolerante.
SITE_HOST = "caseonit.sharepoint.com"
SITE_NAME = "Sutel"
LIBRARY = "Documentos"  # se usa búsqueda parcial; "Documentos" / "Documentos compartidos" / "Shared Documents"
FOLDER_PATH = "01. Documentos MedUX/Automatizacion/Masterfile"

ARCHIVOS = {
    "Fijo": "MasterfileSutel.xlsx",
    "Movilidad": "MasterfileSutel_Movilidad.xlsx"
}

# ================== CONFIG SMTP ==================
SMTP_SERVER = get_secret("smtp_server")
SMTP_PORT = get_secret("smtp_port")
SMTP_USER = get_secret("smtp_user")
SMTP_PASS = get_secret("smtp_pass")
EMAIL_FROM = get_secret("email_from")
EMAIL_TO = get_secret("email_to")

# ================== Parámetros ==================
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
    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
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
    if resp.status_code != 200:
        raise Exception(f"Error buscando sites: {resp.status_code} {resp.text}")

    sites = resp.json().get("value", [])
    if not sites:
        raise Exception(f"No se encontraron sites con search='{SITE_NAME}'")

    site = None
    for s in sites:
        weburl = s.get("webUrl", "")
        if SITE_HOST in weburl and SITE_NAME in weburl:
            site = s
            break
    if site is None:
        site = sites[0]

    site_id = site["id"]

    drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    drives_resp = requests.get(drives_url, headers=headers)
    if drives_resp.status_code != 200:
        raise Exception(f"Error listando drives: {drives_resp.status_code} {drives_resp.text}")

    drives = drives_resp.json().get("value", [])
    drive = drives[0]

    return site_id, drive["id"]

def get_file_from_sharepoint(path):
    token = get_access_token_cached()
    site_id, drive_id = get_site_drive_cached()
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{path}:/content"
    resp = requests.get(url, headers=headers)
    if resp.status_code != 200:
        raise Exception(f"Error descargando archivo {path}: {resp.status_code} {resp.text}")
    return BytesIO(resp.content)

import time

def upload_file_to_sharepoint(path, file_bytes, max_retries=5):
    token = get_access_token_cached()
    site_id, drive_id = get_site_drive_cached()

    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{path}:/content"

    for intento in range(max_retries):
        resp = requests.put(
            url,
            headers=headers,
            data=file_bytes.getvalue(),
            timeout=300
        )

        if resp.status_code in (200, 201):
            return

        # 🔥 manejar lock
        if resp.status_code in (423, 429, 500, 503, 504):
            wait = 2 ** intento
            time.sleep(wait)
            continue

        raise Exception(f"Error subiendo archivo {path}: {resp.status_code} {resp.text}")

    raise Exception(f"Archivo bloqueado después de {max_retries} intentos: {path}")

def ensure_folder(path):
    token = get_access_token_cached()
    site_id, drive_id = get_site_drive_cached()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{path}"
    resp = requests.get(url, headers=headers)
    if resp.status_code == 200:
        return

    parent = "/".join(path.split("/")[:-1])
    folder_name = path.split("/")[-1]
    if parent == "":
        create_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root/children"
    else:
        create_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{parent}:/children"

    body = {"name": folder_name, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"}
    create_resp = requests.post(create_url, headers=headers, json=body)
    if create_resp.status_code not in (200, 201) and create_resp.status_code != 409:
        raise Exception(f"Error creando carpeta {path}: {create_resp.status_code} {create_resp.text}")

# ========= Envío de correo =========
def enviar_correo_con_adjuntos(asunto, cuerpo, archivos_adjuntos):
    msg = EmailMessage()
    msg["Subject"] = asunto
    msg["From"] = EMAIL_FROM
    msg["To"] = EMAIL_TO
    msg.set_content(cuerpo)

    for archivo_bytes, nombre_archivo in archivos_adjuntos:
        msg.add_attachment(
            archivo_bytes.getvalue(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=nombre_archivo
        )

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
        smtp.starttls()
        smtp.login(SMTP_USER, SMTP_PASS)
        smtp.send_message(msg)

# ========= Contador persistente =========
def _leer_contador_hoy():
    fecha_hoy = datetime.now(ZoneInfo("America/Costa_Rica")).strftime("%d%m%Y")
    contador_actual = 0
    try:
        stream = get_file_from_sharepoint(f"{FOLDER_PATH}/contador_envios.txt")
        contenido = stream.read().decode("utf-8").strip()
        partes = contenido.split(",")
        if len(partes) == 2:
            fecha_guardada, cnt = partes
            if fecha_guardada == fecha_hoy:
                contador_actual = int(cnt)
    except Exception:
        contador_actual = 0
    return fecha_hoy, contador_actual

def _guardar_contador_hoy(fecha_ddmmaaaa, nuevo_contador):
    contenido = f"{fecha_ddmmaaaa},{nuevo_contador}".encode("utf-8")
    out = BytesIO(contenido)
    upload_file_to_sharepoint(f"{FOLDER_PATH}/contador_envios.txt", out)

# ========= Normalización para comparar =========
PHANTOM_PATTERNS = [r"^Unnamed", r"::auto_unique_id::", r"^index$", r"^Index$"]

def drop_phantom_cols(df):
    if df is None or df.empty:
        return df
    mask = np.zeros(len(df.columns), dtype=bool)
    for pat in PHANTOM_PATTERNS:
        mask |= df.columns.astype(str).str.contains(pat, regex=True, na=False)
    return df.loc[:, ~mask]

def normalize_df_for_compare(df):
    if df is None or df.empty:
        return df
    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]

    def to_cmp(v):
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return ""
        s = str(v).strip()
        try:
            f = float(s.replace(",", ""))
            return str(int(f)) if f.is_integer() else str(f)
        except Exception:
            return s

    for c in out.columns:
        out[c] = out[c].map(to_cmp)
    return out

# ========= Comparación =========
def detectar_cambios(df_original, df_modificado, tipo_archivo):
    if df_original.empty or df_modificado.empty:
        return []

    df_o = drop_phantom_cols(df_original).copy()
    df_m = drop_phantom_cols(df_modificado).copy()

    have_rowkey_o = ROWKEY in df_o.columns
    have_rowkey_m = ROWKEY in df_m.columns
    use_rowkey = have_rowkey_o and have_rowkey_m

    if not use_rowkey and ID_COL not in df_o.columns:
        return []

    no = normalize_df_for_compare(df_o)
    nm = normalize_df_for_compare(df_m)

    def obtener_identificador(row, k):
        if tipo_archivo.lower() == "fijo" and "Stm" in row.index and pd.notna(row["Stm"]):
            return f"Stm {row['Stm']}"
        elif tipo_archivo.lower() == "movilidad" and "NOMBRE PANELISTA" in row.index and pd.notna(row["NOMBRE PANELISTA"]):
            return f"Panelista {row['NOMBRE PANELISTA']}"
        elif ID_COL in row.index:
            return f"ID {row[ID_COL]}"
        else:
            return f"Fila {k}"

    cambios = []
    if use_rowkey:
        no_idx = no.set_index(ROWKEY, drop=False)
        nm_idx = nm.set_index(ROWKEY, drop=False)
        comunes = sorted(set(no_idx.index) & set(nm_idx.index))
        cols = [c for c in no.columns if c in nm.columns and c != ROWKEY]
        for k in comunes:
            ro = no_idx.loc[k]
            rm = nm_idx.loc[k]
            if isinstance(ro, pd.DataFrame): ro = ro.iloc[0]
            if isinstance(rm, pd.DataFrame): rm = rm.iloc[0]
            for c in cols:
                if ro[c] != rm[c]:
                    ident = obtener_identificador(ro, k)
                    cambios.append(f"{ident}: {c} de {ro[c]} → {rm[c]}")
    else:
        no_idx = no.drop_duplicates(subset=[ID_COL]).set_index(ID_COL, drop=False)
        nm_idx = nm.drop_duplicates(subset=[ID_COL]).set_index(ID_COL, drop=False)
        comunes = sorted(set(no_idx.index) & set(nm_idx.index))
        cols = [c for c in no.columns if c in nm.columns and c != ID_COL]
        for k in comunes:
            ro = no_idx.loc[k]
            rm = nm_idx.loc[k]
            if isinstance(ro, pd.DataFrame): ro = ro.iloc[0]
            if isinstance(rm, pd.DataFrame): rm = rm.iloc[0]
            for c in cols:
                if ro[c] != rm[c]:
                    ident = obtener_identificador(ro, k)
                    cambios.append(f"{ident}: {c} de {ro[c]} → {rm[c]}")
    return cambios

# ========= Manejo de archivos =========
def manejar_archivo(nombre_modo, nombre_archivo, autosize=True):

    # 1. Obtenemos el flujo de datos desde SharePoint
    file_stream = get_file_from_sharepoint(f"{FOLDER_PATH}/{nombre_archivo}")
    
    # 2. Guardamos el contenido binario EXACTO para la descarga local
    contenido_binario = file_stream.getvalue()

    # 3. Leemos el DataFrame (importante hacer seek(0) para resetear el puntero)
    file_stream.seek(0)
    df_original = pd.read_excel(file_stream, dtype={0: str, 1: str})
    df_original[ROWKEY] = np.arange(len(df_original)).astype(str)

    st.success(f"📂 Cargado {nombre_archivo} ✅")

    # --- NUEVO: Botón de descarga del archivo original ---
    st.download_button(
        label=f"📥 Descargar versión cargada de ({nombre_modo})",
        data=contenido_binario,
        file_name=nombre_archivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"download_btn_{nombre_modo}" # Key única para evitar errores de Streamlit
    )
    # -----------------------------------------------------

    gb = GridOptionsBuilder.from_dataframe(df_original)

    gb.configure_default_column(
        editable=True, 
        resizable=True, 
        filter=True, 
        sortable=True, 
        suppressMovable=True,
        minWidth=200,
    )
    gb.configure_pagination(enabled=False)
    gb.configure_column(ROWKEY, hide=True, editable=False)

    gb.configure_grid_options(
        suppressSizeToFit=True,
        domLayout="normal",
        suppressHorizontalScroll=False,
        suppressColumnVirtualisation=False,
        alwaysShowHorizontalScroll=True,
    
        onGridReady=JsCode("""
            function(params) {
                setTimeout(function() {
                    const allColumnIds = [];
                    params.columnApi.getColumns().forEach(function(column) {
                        allColumnIds.push(column.getId());
                    });
                    params.columnApi.autoSizeColumns(allColumnIds);
                }, 300);
            }
        """)
    )

    grid_options = gb.build()

    grid_response = AgGrid(
        df_original,
        gridOptions=grid_options,
        height=500,
        fit_columns_on_grid_load=False,
        enable_enterprise_modules=False,
        update_mode=GridUpdateMode.MODEL_CHANGED,
        data_return_mode=DataReturnMode.AS_INPUT,
        allow_unsafe_jscode=True,
        theme="balham",
        reload_data=False,
        width="100%"
    )

    df_modificado = pd.DataFrame(grid_response["data"]).copy()

    return df_modificado
# ================== INTERFAZ PRINCIPAL ==================
try:
    tab_fijo, tab_movilidad = st.tabs(["📄 Masterfile Fijo", "📄 Masterfile Movilidad"])

    with tab_fijo:
        df_fijo = manejar_archivo("Fijo", ARCHIVOS["Fijo"])

    with tab_movilidad:
        df_movilidad = manejar_archivo("Movilidad", ARCHIVOS["Movilidad"])

    if st.button("💾 Guardar nueva versión de Masterfile"):
        timestamp = datetime.now(ZoneInfo("America/Costa_Rica")).strftime("%Y%m%d_%H%M%S")
        archivos_adjuntos = []
        cuerpo_correo = f"Buen día,\n\nSe adjunta nueva versión de Masterfile con los cambios realizados el {timestamp}.\n\n"

        for nombre_modo, df_modificado, nombre_archivo in [
            ("Fijo", df_fijo, ARCHIVOS["Fijo"]),
            ("Movilidad", df_movilidad, ARCHIVOS["Movilidad"])
        ]:
            df_original_stream = get_file_from_sharepoint(f"{FOLDER_PATH}/{nombre_archivo}")
            df_original = pd.read_excel(df_original_stream, dtype={0: str, 1: str})
            df_original[ROWKEY] = np.arange(len(df_original)).astype(str)

            cambios = detectar_cambios(df_original, df_modificado, nombre_modo)
            if cambios:
                filas_cambiadas = "\n" + "\n".join([f"• {c}" for c in cambios])
            else:
                filas_cambiadas = "Ningún cambio detectado"

            cuerpo_correo += f"📌 Cambios en entorno {nombre_modo}:\n{filas_cambiadas}\n\n"

            df_a_guardar = df_modificado.copy()
            if ROWKEY in df_a_guardar.columns:
                df_a_guardar = df_a_guardar.drop(columns=[ROWKEY])

            nuevo_nombre = f"{nombre_archivo.replace('.xlsx','')}_{timestamp}.xlsx"
            bytes_excel = BytesIO()
            df_a_guardar.to_excel(bytes_excel, index=False)
            bytes_excel.seek(0)

            backup_folder = f"{FOLDER_PATH}/Backups/{nombre_modo}"
            ensure_folder(backup_folder)
            upload_file_to_sharepoint(f"{backup_folder}/{nuevo_nombre}", bytes_excel)

            bytes_excel.seek(0)
            upload_file_to_sharepoint(f"{FOLDER_PATH}/{nombre_archivo}", bytes_excel)

            bytes_excel.seek(0)
            archivos_adjuntos.append((BytesIO(bytes_excel.getvalue()), nuevo_nombre))

        fecha_ddmmaaaa, contador_actual = _leer_contador_hoy()
        if contador_actual == 0:
            asunto_correo = f"Masterfile Sutel Fijo y Movilidad {fecha_ddmmaaaa}"
            siguiente_contador = 1
        else:
            asunto_correo = f"Masterfile Sutel y Movilidad {fecha_ddmmaaaa} V{contador_actual + 1}"
            siguiente_contador = contador_actual + 1

        try:
            enviar_correo_con_adjuntos(
                asunto=asunto_correo,
                cuerpo=cuerpo_correo + "Un saludo",
                archivos_adjuntos=archivos_adjuntos
            )
            _guardar_contador_hoy(fecha_ddmmaaaa, siguiente_contador)
            st.success("📧 Correo enviado notificando la nueva versión de ambos Masterfiles.")
        except Exception as e:
            st.error(f"Error al enviar correo: {e}")

except Exception as e:
    st.error(f"Error: {e}")
