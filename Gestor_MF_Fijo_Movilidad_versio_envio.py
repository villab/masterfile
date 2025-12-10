# VERSION FINAL - MASTERFILE con MSAL y Microsoft Graph API
# ==============================================
# - Backups separados para FIJO y MOVILIDAD
# - Envío de correo con contador persistente por día
# - Detección de cambios valor viejo → valor nuevo
# ==============================================

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

# ------ Configuración de vista ----------
st.set_page_config(layout="wide")
st.title("📋 Masterfile Entorno de medición Fijo y Movilidad")

# ================== CONFIGURACIÓN ==================
TENANT_ID = st.secrets["tenant_id"]
CLIENT_ID = st.secrets["client_id"]
CLIENT_SECRET = st.secrets["client_secret"]

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
SMTP_SERVER = st.secrets["smtp_server"]
SMTP_PORT = st.secrets["smtp_port"]
SMTP_USER = st.secrets["smtp_user"]
SMTP_PASS = st.secrets["smtp_pass"]
EMAIL_FROM = st.secrets["email_from"]
EMAIL_TO = st.secrets["email_to"]

# ================== Parámetros ==================
ID_COL = "ID SONDA"
ROWKEY = "_row_id"

# ========= Autenticación con MSAL =========
def get_access_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        # Mostrar detalle en UI para depuración
        st.error(f"❌ No se pudo obtener token de acceso: {result}")
        raise Exception("No se pudo obtener token de acceso")
    return result["access_token"]

# ========= Funciones SharePoint con Graph (más robustas) =========
def _get_site_and_drive(token):
    """
    Devuelve (site_id, drive_id).
    - Busca sites con ?search=SITE_NAME y elige el que contenga SITE_HOST y SITE_NAME si es posible.
    - Busca drive por coincidencia parcial del nombre LIBRARY (case-insensitive).
    """
    headers = {"Authorization": f"Bearer {token}"}

    # 1) Buscar sites que coincidan con SITE_NAME
    search_url = f"https://graph.microsoft.com/v1.0/sites?search={SITE_NAME}"
    resp = requests.get(search_url, headers=headers)
    if resp.status_code != 200:
        raise Exception(f"Error buscando sites: {resp.status_code} {resp.text}")

    sites = resp.json().get("value", [])
    if not sites:
        raise Exception(f"No se encontraron sites con search='{SITE_NAME}' (respuesta vacía)")

    # Intentar elegir el site correcto: preferir site cuyo webUrl contenga SITE_HOST y SITE_NAME
    site = None
    for s in sites:
        weburl = s.get("webUrl", "")
        if SITE_HOST in weburl and SITE_NAME in weburl:
            site = s
            break
    if site is None:
        # fallback: tomar el primero cuyo name contenga SITE_NAME (case-insensitive)
        for s in sites:
            if SITE_NAME.lower() in s.get("name", "").lower():
                site = s
                break
    if site is None:
        # si aún no hay coincidencia, tomar el primer site devuelto
        site = sites[0]

    site_id = site.get("id")
    if not site_id:
        raise Exception(f"Site encontrado no tiene 'id': {site}")

    # 2) Obtener drives (bibliotecas) del site
    drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    drives_resp = requests.get(drives_url, headers=headers)
    if drives_resp.status_code != 200:
        raise Exception(f"Error listando drives: {drives_resp.status_code} {drives_resp.text}")

    drives = drives_resp.json().get("value", [])
    if not drives:
        raise Exception("No se encontraron drives en el site.")

    # Buscar drive por coincidencia parcial en el nombre (case-insensitive)
    drive = None
    for d in drives:
        if LIBRARY.lower() in (d.get("name") or "").lower():
            drive = d
            break
    if drive is None:
        # intentar nombres comunes
        for d in drives:
            if "documents" in (d.get("name") or "").lower() or "documentos" in (d.get("name") or "").lower():
                drive = d
                break
    if drive is None:
        # fallback al primer drive
        drive = drives[0]

    drive_id = drive.get("id")
    if not drive_id:
        raise Exception(f"Drive encontrado no tiene 'id': {drive}")

    # devolver ids
    return site_id, drive_id

def get_file_from_sharepoint(path):
    """
    path: ruta relativa dentro del drive, por ejemplo:
      "01. Documentos MedUX/Automatizacion/Masterfile/MasterfileSutel.xlsx"
    devuelve BytesIO con el contenido.
    """
    token = get_access_token()
    site_id, drive_id = _get_site_and_drive(token)
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{path}:/content"
    resp = requests.get(url, headers=headers)
    if resp.status_code != 200:
        raise Exception(f"Error descargando archivo {path}: {resp.status_code} {resp.text}")
    return BytesIO(resp.content)

def upload_file_to_sharepoint(path, file_bytes):
    """
    path: ruta completa dentro del drive (incluyendo el nombre del archivo).
    file_bytes: BytesIO
    """
    token = get_access_token()
    site_id, drive_id = _get_site_and_drive(token)
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{path}:/content"
    # usar PUT para sobrescribir
    resp = requests.put(url, headers=headers, data=file_bytes.getvalue())
    if resp.status_code not in (200, 201):
        raise Exception(f"Error subiendo archivo {path}: {resp.status_code} {resp.text}")

def ensure_folder(path):
    """
    Asegura que exista la carpeta `path` bajo root. path relativo p.ej "Backups/Fijo".
    Crea la carpeta si no existe (intenta crear solo la última carpeta).
    """
    token = get_access_token()
    site_id, drive_id = _get_site_and_drive(token)
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    # comprobar si carpeta existe
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{path}"
    resp = requests.get(url, headers=headers)
    if resp.status_code == 200:
        return  # ya existe

    # si no existe, crearla en el padre
    parent = "/".join(path.split("/")[:-1])
    folder_name = path.split("/")[-1]
    if parent == "":
        create_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root/children"
    else:
        create_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{parent}:/children"

    body = {"name": folder_name, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"}
    create_resp = requests.post(create_url, headers=headers, json=body)
    if create_resp.status_code not in (200, 201):
        # si ya existe (conflict) podemos ignorar, sino lanzar error
        if create_resp.status_code == 409:
            return
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
def manejar_archivo(nombre_modo, nombre_archivo):
    file_stream = get_file_from_sharepoint(f"{FOLDER_PATH}/{nombre_archivo}")
    df_original = pd.read_excel(file_stream, dtype={0: str, 1: str})
    df_original[ROWKEY] = np.arange(len(df_original)).astype(str)

    st.success(f"📂 Cargado {nombre_archivo} ✅")

    gb = GridOptionsBuilder.from_dataframe(df_original)
    gb.configure_default_column(editable=True, resizable=True, filter=True, sortable=True, suppressMovable=True)
    gb.configure_pagination(enabled=False)
    gb.configure_column(ROWKEY, hide=True, editable=False)

    gb.configure_grid_options(
        onFirstDataRendered=JsCode("""
            function(params) {
                setTimeout(function() {
                    let allColumnIds = [];
                    params.columnApi.getAllColumns().forEach(function(column) {
                        allColumnIds.push(column.getId());
                    });
                    params.columnApi.autoSizeColumns(allColumnIds);
                }, 250);  // <-- delay para que Streamlit termine de renderizar
            }
        """)
    )

    grid_options = gb.build()

    grid_response = AgGrid(
        df_original,
        gridOptions=grid_options,
        height=500,
        fit_columns_on_grid_load=True,   # <--- cambiar a True
        enable_enterprise_modules=False,
        update_mode=GridUpdateMode.VALUE_CHANGED,
        data_return_mode=DataReturnMode.AS_INPUT,
        allow_unsafe_jscode=True,
        theme="balham",
        reload_data=False
    )

    df_modificado = pd.DataFrame(grid_response["data"])
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


