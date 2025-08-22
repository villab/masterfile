# VERSION CON SUBCARPETAS DE BACKUP SEPARADAS PARA FIJO Y MOVILIDAD
# y envío de correo con ambas últimas versiones con fecha y viñetas de cambios por archivo

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
import os
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode, JsCode
from zoneinfo import ZoneInfo
import smtplib
from email.message import EmailMessage
import re

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

# ================== Parámetros de comparación ==================
ID_COL = "ID SONDA"  # identificador único por fila para comparar ediciones

# -------- Helpers correo --------
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

# ========= Helpers para contador persistente en SharePoint =========
def _leer_contador_hoy(ctx):
    """Lee el contador de envíos de hoy desde contador_envios.txt en SharePoint.
    Devuelve (fecha_ddmmaaaa, contador_actual). Si no existe o es otro día, contador=0."""
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
        # Si no existe el archivo o hay error de lectura, asumimos contador 0.
        contador_actual = 0
    return fecha_hoy, contador_actual

def _guardar_contador_hoy(ctx, fecha_ddmmaaaa, nuevo_contador):
    """Guarda el contador de envíos de hoy en contador_envios.txt en SharePoint."""
    contenido = f"{fecha_ddmmaaaa},{nuevo_contador}".encode("utf-8")
    out = BytesIO(contenido)
    out.seek(0)
    folder = ctx.web.get_folder_by_server_relative_url(FOLDER_URL)
    folder.upload_file("contador_envios.txt", out).execute_query()

# -------- Normalización y comparación robusta --------
_PHANTOM_PATTERNS = [
    r"^Unnamed",                       # columnas de Excel sin nombre
    r"::auto_unique_id::",             # ids internos de ag-Grid
    r"^index$", r"^Index$",            # índices exportados
    r"^_?RowId$", r"^ag-Grid",         # otras variantes
]

def _drop_phantom_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    cols = df.columns
    mask = np.zeros(len(cols), dtype=bool)
    for pat in _PHANTOM_PATTERNS:
        mask |= cols.astype(str).str.contains(pat, regex=True, na=False)
    keep_cols = cols[~mask]
    return df[keep_cols]

def _normalize_for_compare(df: pd.DataFrame, id_col: str) -> pd.DataFrame:
    """Devuelve una copia limpia solo para comparar (no se usa para guardar).
    - Elimina columnas fantasma.
    - Asegura que id_col sea string.
    - Convierte todas las celdas a texto comparable (quita espacios, uniformiza NaN)."""
    if df is None or df.empty:
        return df
    clean = _drop_phantom_columns(df).copy()

    # Trim nombres de columnas (evita 'ID SONDA ' con espacio final)
    clean.columns = [str(c).strip() for c in clean.columns]

    # Forzar ID a string si existe
    if id_col in clean.columns:
        clean[id_col] = clean[id_col].astype(str).str.strip()

    # Uniformizar valores para comparación (no altera df real que se guarda)
    def to_cmp_series(s: pd.Series) -> pd.Series:
        # Mantener NaN coherente
        s2 = s.copy()
        # Si ambos lados son numéricos, compararemos por valor; pero como no sabemos el tipo final
        # generamos una versión string "limpia":
        # - rellenar NaN con vacío
        # - convertir a string
        # - quitar espacios
        # - normalizar '1.0' y '1' a mismo texto cuando es número
        s2 = s2.replace({np.nan: None})
        out = []
        for v in s2:
            if v is None:
                out.append("")
                continue
            # intentar numérico
            try:
                fv = float(str(v).strip().replace(",", ""))  # por si hay separadores
                # representar sin '.0' innecesario
                if fv.is_integer():
                    out.append(str(int(fv)))
                else:
                    out.append(str(fv))
            except Exception:
                out.append(str(v).strip())
        return pd.Series(out, index=s2.index)

    for c in clean.columns:
        clean[c] = to_cmp_series(clean[c])

    return clean

def _equalish(a, b) -> bool:
    """Comparación tolerante: NaN/None vacíos, números equivalentes, strings con trim."""
    if a == b:
        return True
    # Vacíos equivalentes
    if (a in [None, ""] and b in [None, ""]) or (pd.isna(a) and pd.isna(b)):
        return True
    # Intentar comparar numéricamente
    try:
        fa = float(str(a))
        fb = float(str(b))
        return abs(fa - fb) < 1e-9
    except Exception:
        pass
    # Fallback string
    return str(a).strip() == str(b).strip()

def detectar_cambios_por_edicion(df_original: pd.DataFrame, df_modificado: pd.DataFrame, id_col: str) -> list[str]:
    """Devuelve una lista de descripciones de cambios reales por edición de celdas.
    Ignora filtros/ordenamientos. No marca filas nuevas/eliminadas."""
    # Normalizar solo para comparar (no afecta guardado)
    orig_cmp = _normalize_for_compare(df_original, id_col)
    mod_cmp  = _normalize_for_compare(df_modificado, id_col)

    if orig_cmp is None or mod_cmp is None or orig_cmp.empty or mod_cmp.empty:
        return []

    if id_col not in orig_cmp.columns or id_col not in mod_cmp.columns:
        # Sin columna identificadora no se puede comparar de forma robusta
        return []

    # Filas comunes por ID
    ids_comunes = set(orig_cmp[id_col].dropna()) & set(mod_cmp[id_col].dropna())
    if not ids_comunes:
        return []

    cambios = []
    # Columnas comunes (excluye ID)
    cols_comunes = [c for c in orig_cmp.columns if c in mod_cmp.columns and c != id_col]

    # Crear índice por ID para acceso O(1)
    orig_idx = orig_cmp.set_index(id_col)
    mod_idx  = mod_cmp.set_index(id_col)

    for _id in sorted(ids_comunes):
        if _id not in orig_idx.index or _id not in mod_idx.index:
            continue
        row_o = orig_idx.loc[_id]
        row_m = mod_idx.loc[_id]
        for c in cols_comunes:
            if not _equalish(row_o[c], row_m[c]):
                cambios.append(f"Fila { _id }: { c } de { row_o[c] } → { row_m[c] }")
    return cambios

# -------- Carga/edición de archivos --------
def manejar_archivo(nombre_modo, nombre_archivo):
    """Carga, muestra, permite editar y devolver el DataFrame editado"""
    ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, APP_PASSWORD))
    FILE_URL = f"{FOLDER_URL}/{nombre_archivo}"

    # Descargar archivo original
    file = ctx.web.get_file_by_server_relative_url(FILE_URL)
    file_stream = BytesIO()
    file.download(file_stream).execute_query()
    file_stream.seek(0)

    # Leer Excel original; forzar primeras columnas a string si aplica (manteniendo tu lógica)
    df_original = pd.read_excel(file_stream, dtype={0: str, 1: str})
    st.success(f"📂 Cargado {nombre_archivo} ✅")

    # Mostrar tabla editable con auto-size de columnas
    gb = GridOptionsBuilder.from_dataframe(df_original)
    gb.configure_default_column(editable=True, resizable=True, filter=True, sortable=True, suppressMovable=True)
    gb.configure_pagination(enabled=False)
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
            # Descargar original para comparar (sin tocar vista)
            df_original_stream = BytesIO()
            ctx.web.get_file_by_server_relative_url(f"{FOLDER_URL}/{nombre_archivo}").download(df_original_stream).execute_query()
            df_original_stream.seek(0)
            df_original = pd.read_excel(df_original_stream, dtype={0: str, 1: str})

            # Detectar cambios SOLO por ediciones reales (ignora filtros/ordenamientos)
            cambios = detectar_cambios_por_edicion(df_original, df_modificado, ID_COL)

            if cambios:
                filas_cambiadas = "\n" + "\n".join([f"• {c}" for c in cambios])
            else:
                filas_cambiadas = "Ningún cambio detectado"

            cuerpo_correo += f"📌 Cambios en entorno {nombre_modo}:\n{filas_cambiadas}\n\n"

            # Guardar nuevo archivo (el que el usuario ve/editó)
            nuevo_nombre = f"{nombre_archivo.replace('.xlsx','')}_{timestamp}.xlsx"
            output_stream = BytesIO()
            df_modificado.to_excel(output_stream, index=False)
            output_stream.seek(0)

            # Crear carpeta de backup por modo
            backup_folder_url = f"{FOLDER_URL}/Backups/{nombre_modo}"
            try:
                ctx.web.get_folder_by_server_relative_url(backup_folder_url).expand(["Files"]).get().execute_query()
            except:
                ctx.web.folders.add(backup_folder_url).execute_query()

            # Subir copia a Backup
            backup_folder = ctx.web.get_folder_by_server_relative_url(backup_folder_url)
            backup_folder.upload_file(nuevo_nombre, output_stream).execute_query()

            # Subir archivo principal actualizado
            output_stream.seek(0)
            main_folder = ctx.web.get_folder_by_server_relative_url(FOLDER_URL)
            main_folder.upload_file(nombre_archivo, output_stream).execute_query()

            # Adjuntar para correo
            archivos_adjuntos.append((output_stream, nuevo_nombre))

        # ===== Asunto con fecha ddmmaaaa y versión basada en contador persistente =====
        fecha_ddmmaaaa, contador_actual = _leer_contador_hoy(ctx)
        if contador_actual == 0:
            asunto_correo = f"Masterfile Sutel Fijo y Movilidad {fecha_ddmmaaaa}"
            siguiente_contador = 1
        else:
            asunto_correo = f"Masterfile Sutel Fijo y Movilidad {fecha_ddmmaaaa} V{contador_actual + 1}"
            siguiente_contador = contador_actual + 1

        # Enviar correo con ambos archivos y viñetas
        try:
            enviar_correo_con_adjuntos(
                asunto=asunto_correo,
                cuerpo=cuerpo_correo + "Un saludo",
                archivos_adjuntos=archivos_adjuntos
            )
            # Actualizar contador SOLO si el correo se envió correctamente
            _guardar_contador_hoy(ctx, fecha_ddmmaaaa, siguiente_contador)
            st.success("📧 Correo enviado notificando la nueva versión de ambos Masterfiles.")
        except Exception as e:
            st.error(f"Error al enviar correo: {e}")

except Exception as e:
    st.error(f"Error: {e}")
