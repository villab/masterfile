# VERSION CON SUBCARPETAS DE BACKUP SEPARADAS PARA FIJO Y MOVILIDAD
# y envío de correo con ambas últimas versiones con fecha y viñetas de cambios por archivo

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode, JsCode
from zoneinfo import ZoneInfo
import smtplib
from email.message import EmailMessage

# ------ Configuración de vista de la pagina ----------
st.set_page_config(layout="wide")
st.title("📋 Masterfile Entorno de medición Fijo y Movilidad")

# ================== CONFIGURACIÓN ==================
USERNAME = st.secrets["sharepoint_user"]
APP_PASSWORD = st.secrets["app_password"]

SITE_URL = "https://caseonit.sharepoint.com/sites/Sutel"
FOLDER_URL = "/sites/Sutel/Documentos compartidos/01. Documentos MedUX/Automatizacion/Masterfile"

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
ID_COL = "ID SONDA"     # identificador “lógico” si hiciera falta
ROWKEY = "_row_id"      # identificador “físico” estable por fila para comparar

# ========= Envío de correo =========
def enviar_correo_con_adjuntos(asunto, cuerpo, archivos_adjuntos):
    """Envia un correo con múltiples adjuntos (archivo_bytes, nombre_archivo)"""
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

# ========= Contador persistente en SharePoint =========
def _leer_contador_hoy(ctx):
    fecha_hoy = datetime.now(ZoneInfo("America/Costa_Rica")).strftime("%d%m%Y")
    contador_url = f"{FOLDER_URL}/contador_envios.txt"
    contador_actual = 0
    try:
        stream = BytesIO()
        ctx.web.get_file_by_server_relative_url(contador_url).download(stream).execute_query()
        stream.seek(0)
        contenido = stream.read().decode("utf-8").strip()
        partes = contenido.split(",")
        if len(partes) == 2:
            fecha_guardada, cnt = partes
            if fecha_guardada == fecha_hoy:
                contador_actual = int(cnt)
            else:
                contador_actual = 0
        else:
            contador_actual = 0
    except Exception:
        contador_actual = 0
    return fecha_hoy, contador_actual

def _guardar_contador_hoy(ctx, fecha_ddmmaaaa, nuevo_contador):
    contenido = f"{fecha_ddmmaaaa},{nuevo_contador}".encode("utf-8")
    out = BytesIO(contenido)
    out.seek(0)
    folder = ctx.web.get_folder_by_server_relative_url(FOLDER_URL)
    folder.upload_file("contador_envios.txt", out).execute_query()

# ========= Limpieza/normalización para comparar =========
PHANTOM_PATTERNS = [r"^Unnamed", r"::auto_unique_id::", r"^index$", r"^Index$"]

def drop_phantom_cols(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    mask = np.zeros(len(df.columns), dtype=bool)
    for pat in PHANTOM_PATTERNS:
        mask |= df.columns.astype(str).str.contains(pat, regex=True, na=False)
    return df.loc[:, ~mask]

def normalize_df_for_compare(df: pd.DataFrame) -> pd.DataFrame:
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

# ========= Comparación usando _row_id (fallback por ID SONDA) =========
def detectar_cambios(df_original: pd.DataFrame, df_modificado: pd.DataFrame) -> list[str]:
    if df_original is None or df_modificado is None or df_original.empty or df_modificado.empty:
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

    if use_rowkey:
        no_idx = no.set_index(ROWKEY, drop=False)
        nm_idx = nm.set_index(ROWKEY, drop=False)
        comunes = sorted(set(no_idx.index) & set(nm_idx.index))
        cols = [c for c in no.columns if c in nm.columns and c != ROWKEY]
        cambios = []
        for k in comunes:
            ro = no_idx.loc[k]
            rm = nm_idx.loc[k]
            if isinstance(ro, pd.DataFrame): ro = ro.iloc[0]
            if isinstance(rm, pd.DataFrame): rm = rm.iloc[0]
            for c in cols:
                if ro[c] != rm[c]:
                    # 🔹 usar STM como identificador si existe
                    if "Stm" in ro.index:
                        ident = ro["Stm"]
                        cambios.append(f"STM {ident}: {c} de {ro[c]} → {rm[c]}")
                    else:
                        ident = ro.get(ID_COL, k)
                        cambios.append(f"Fila {ident}: {c} de {ro[c]} → {rm[c]}")
        return cambios

    else:
        no_idx = no.drop_duplicates(subset=[ID_COL]).set_index(ID_COL, drop=False)
        nm_idx = nm.drop_duplicates(subset=[ID_COL]).set_index(ID_COL, drop=False)
        comunes = sorted(set(no_idx.index) & set(nm_idx.index))
        cols = [c for c in no.columns if c in nm.columns and c != ID_COL]
        cambios = []
        for k in comunes:
            ro = no_idx.loc[k]
            rm = nm_idx.loc[k]
            if isinstance(ro, pd.DataFrame): ro = ro.iloc[0]
            if isinstance(rm, pd.DataFrame): rm = rm.iloc[0]
            for c in cols:
                if ro[c] != rm[c]:
                    if "STM" in ro.index:
                        ident = ro["STM"]
                        cambios.append(f"STM {ident}: {c} de {ro[c]} → {rm[c]}")
                    else:
                        cambios.append(f"Fila {k}: {c} de {ro[c]} → {rm[c]}")
        return cambios

# ========= Carga/edición (inyecta _row_id oculto) =========
def manejar_archivo(nombre_modo, nombre_archivo):
    ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, APP_PASSWORD))
    FILE_URL = f"{FOLDER_URL}/{nombre_archivo}"

    file = ctx.web.get_file_by_server_relative_url(FILE_URL)
    file_stream = BytesIO()
    file.download(file_stream).execute_query()
    file_stream.seek(0)

    df_original = pd.read_excel(file_stream, dtype={0: str, 1: str})
    df_original[ROWKEY] = np.arange(len(df_original)).astype(int).astype(str)

    st.success(f"📂 Cargado {nombre_archivo} ✅")

    gb = GridOptionsBuilder.from_dataframe(df_original)
    gb.configure_default_column(editable=True, resizable=True, filter=True, sortable=True, suppressMovable=True)
    gb.configure_pagination(enabled=False)
    gb.configure_column(ROWKEY, hide=True, editable=False)
    gb.configure_grid_options(
        onFirstDataRendered=JsCode("""
            function(params) {
                let allColumnIds = [];
                params.columnApi.getAllColumns().forEach(function(column) {
                    allColumnIds.push(column.getId());
                });
                params.columnApi.autoSizeColumns(allColumnIds);
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
        update_mode=GridUpdateMode.VALUE_CHANGED,
        data_return_mode=DataReturnMode.AS_INPUT,
        allow_unsafe_jscode=True,
        theme="balham",
        reload_data=False
    )

    df_modificado = pd.DataFrame(grid_response["data"])

    return df_modificado

# ================== INTERFAZ CON PESTAÑAS ==================
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

        ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, APP_PASSWORD))

        for nombre_modo, df_modificado, nombre_archivo in [
            ("Fijo", df_fijo, ARCHIVOS["Fijo"]),
            ("Movilidad", df_movilidad, ARCHIVOS["Movilidad"])
        ]:
            df_original_stream = BytesIO()
            ctx.web.get_file_by_server_relative_url(f"{FOLDER_URL}/{nombre_archivo}").download(df_original_stream).execute_query()
            df_original_stream.seek(0)
            df_original = pd.read_excel(df_original_stream, dtype={0: str, 1: str})
            df_original[ROWKEY] = np.arange(len(df_original)).astype(int).astype(str)

            cambios = detectar_cambios(df_original, df_modificado)

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

            backup_folder_url = f"{FOLDER_URL}/Backups/{nombre_modo}"
            try:
                ctx.web.get_folder_by_server_relative_url(backup_folder_url).expand(["Files"]).get().execute_query()
            except:
                ctx.web.folders.add(backup_folder_url).execute_query()

            backup_folder = ctx.web.get_folder_by_server_relative_url(backup_folder_url)
            backup_folder.upload_file(nuevo_nombre, bytes_excel).execute_query()

            bytes_excel.seek(0)
            main_folder = ctx.web.get_folder_by_server_relative_url(FOLDER_URL)
            main_folder.upload_file(nombre_archivo, bytes_excel).execute_query()

            bytes_excel.seek(0)
            archivos_adjuntos.append((BytesIO(bytes_excel.getvalue()), nuevo_nombre))

        fecha_ddmmaaaa, contador_actual = _leer_contador_hoy(ctx)
        if contador_actual == 0:
            asunto_correo = f"Masterfile Sutel Fijo y Movilidad {fecha_ddmmaaaa}"
            siguiente_contador = 1
        else:
            asunto_correo = f"Masterfile Sutel Fijo y Movilidad {fecha_ddmmaaaa} V{contador_actual + 1}"
            siguiente_contador = contador_actual + 1

        try:
            enviar_correo_con_adjuntos(
                asunto=asunto_correo,
                cuerpo=cuerpo_correo + "Un saludo",
                archivos_adjuntos=archivos_adjuntos
            )
            _guardar_contador_hoy(ctx, fecha_ddmmaaaa, siguiente_contador)
            st.success("📧 Correo enviado notificando la nueva versión de ambos Masterfiles.")
        except Exception as e:
            st.error(f"Error al enviar correo: {e}")

except Exception as e:
    st.error(f"Error: {e}")

