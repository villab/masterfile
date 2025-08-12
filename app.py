import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
import os
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

#------ Configuración de vista de la pagina----------
st.set_page_config(layout="wide") 

# ================== CONFIGURACIÓN ==================
USERNAME = st.secrets["sharepoint_user"]
APP_PASSWORD = st.secrets["app_password"]

SITE_URL = "https://caseonit.sharepoint.com/sites/Sutel"
FILE_URL = "/sites/Sutel/Documentos compartidos/01. Documentos MedUX/Automatizacion/MasterfileSutel.xlsx"
FOLDER_URL = "/sites/Sutel/Documentos compartidos/01. Documentos MedUX/Automatizacion"
BACKUP_FOLDER_URL = f"{FOLDER_URL}/Backups"

try:
    ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, APP_PASSWORD))

    # Obtener solo el nombre del archivo
    nombre_archivo = os.path.basename(FILE_URL)

    # Descargar archivo original
    file = ctx.web.get_file_by_server_relative_url(FILE_URL)
    file_stream = BytesIO()
    file.download(file_stream).execute_query()
    file_stream.seek(0)

    # ================== LECTURA DEL EXCEL ==================
    df = pd.read_excel(file_stream)
    st.success(f"📂 Cargado  {nombre_archivo} ✅") 

    # ================== Mostrar tabla editable ==================
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(
        editable=True,  # 🔹 Permitir edición en todas las columnas
        resizable=True,
        filter=True,
        sortable=True
    )
    gb.configure_pagination(enabled=False)  # ❌ Sin paginación
    grid_options = gb.build()

    grid_response = AgGrid(
        df,
        gridOptions=grid_options,
        height=500,
        fit_columns_on_grid_load=False,
        enable_enterprise_modules=False,
        update_mode=GridUpdateMode.VALUE_CHANGED,  # 🔹 Detectar cambios
        allow_unsafe_jscode=True,
        theme="balham",
        reload_data=False
    )

    # 🔹 Capturar cambios hechos en la tabla
    df = pd.DataFrame(grid_response["data"])

    # ================== GUARDAR CAMBIOS ==================
    if st.button("💾 Guardar nueva versión de Masterfile"):
        # Nombre con fecha y hora (YYYYMMDD_HHMMSS)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nuevo_nombre = f"MasterfileSutel_{timestamp}.xlsx"

        # Guardar DataFrame en memoria
        output_stream = BytesIO()
        df.to_excel(output_stream, index=False)
        output_stream.seek(0)

        # Verificar o crear carpeta Backups
        try:
            ctx.web.get_folder_by_server_relative_url(BACKUP_FOLDER_URL).expand(["Files"]).get().execute_query()
        except:
            ctx.web.folders.add(BACKUP_FOLDER_URL).execute_query()

        # Subir copia con fecha a Backups
        backup_folder = ctx.web.get_folder_by_server_relative_url(BACKUP_FOLDER_URL)
        backup_folder.upload_file(nuevo_nombre, output_stream).execute_query()

        # Volver a poner el puntero al inicio
        output_stream.seek(0)

        # Sobrescribir el archivo original
        main_folder = ctx.web.get_folder_by_server_relative_url(FOLDER_URL)
        main_folder.upload_file("MasterfileSutel.xlsx", output_stream).execute_query()

        st.success(f"✅ Cambios guardados y copia creada en 'Backups' como {nuevo_nombre}")

except Exception as e:
    st.error(f"Error: {e}")
