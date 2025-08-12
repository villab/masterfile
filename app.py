import streamlit as st
import pandas as pd
from io import BytesIO
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

# -------------------- CONFIG --------------------
#USERNAME = st.secrets["sharepoint_user"]
#PASSWORD = st.secrets["sharepoint_pass"]
#ACCESS_KEY = st.secrets["app_password"]

USERNAME = st.secrets["sharepoint_user"]
APP_PASSWORD = st.secrets["app_password"]

SITE_URL = "https://caseonit.sharepoint.com/sites/Sutel"  # Ajusta a tu sitio real

# -------------------- FUNCIONES SHAREPOINT --------------------
@st.cache_resource
def connect_sharepoint():
    """Conecta a SharePoint usando usuario y contraseña de aplicación"""
    return ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, APP_PASSWORD))

def list_libraries(ctx):
    """Lista todas las bibliotecas (listas tipo documento) del sitio"""
    try:
        lists = ctx.web.lists
        ctx.load(lists)
        ctx.execute_query()
        libraries = [l.properties["Title"] for l in lists if l.properties["BaseTemplate"] == 101]
        return libraries
    except Exception as e:
        st.error(f"Error al listar bibliotecas: {e}")
        return []

def list_files(ctx, library_name):
    """Lista los archivos XLSX de la biblioteca seleccionada"""
    try:
        library = ctx.web.lists.get_by_title(library_name)
        files = library.root_folder.files
        ctx.load(files)
        ctx.execute_query()
        return [f.properties["Name"] for f in files if f.properties["Name"].endswith(".xlsx")]
    except Exception as e:
        st.error(f"Error al listar archivos en {library_name}: {e}")
        return []

def download_file(ctx, library_name, file_name):
    """Descarga un archivo desde la biblioteca y lo devuelve en bytes"""
    try:
        # Ruta relativa correcta al sitio actual
        file_url = f"/sites/Sutel/{library_name}/{file_name}"

        buffer = BytesIO()
        file = ctx.web.get_file_by_server_relative_url(file_url)
        file.download(buffer)
        ctx.execute_query()

        return buffer.getvalue()
    except Exception as e:
        st.error(f"Error al descargar {file_name}: {e}")
        return None

def upload_file(ctx, library_name, file_name, content):
    """Sube o reemplaza un archivo en la biblioteca"""
    try:
        target_folder = ctx.web.lists.get_by_title(library_name).root_folder
        target_folder.upload_file(file_name, content)
        ctx.execute_query()
    except Exception as e:
        st.error(f"Error al subir {file_name}: {e}")

# -------------------- LOGIN --------------------
st.set_page_config(page_title="Gestor Multiusuario XLSX", layout="wide")
st.title("🔐 Gestor Multiusuario de Datos en SharePoint")

password_input = st.text_input("Ingresa la clave de acceso", type="password")
if password_input != APP_PASSWORD:
    st.warning("Introduce la clave correcta para acceder.")
    st.stop()
st.success("✅ Acceso concedido")

# -------------------- APP --------------------
ctx = connect_sharepoint()

# 1. Selección de biblioteca
st.subheader("Selecciona la biblioteca de documentos")
libraries = list_libraries(ctx)
if not libraries:
    st.stop()

library_choice = st.selectbox("Bibliotecas disponibles", libraries)

# 2. Archivos de la biblioteca
files = list_files(ctx, library_choice)
file_choice = None  # Inicializar variable para evitar NameError

if files:
    file_choice = st.selectbox("Selecciona un archivo de SharePoint", [""] + files)
else:
    st.warning("No se encontraron archivos XLSX en esta biblioteca.")

# 3. Subir archivo nuevo
uploaded_file = st.file_uploader("O carga un archivo nuevo", type=["xlsx"])
if uploaded_file:
    upload_file(ctx, library_choice, uploaded_file.name, uploaded_file.getvalue())
    st.success(f"Archivo '{uploaded_file.name}' cargado a SharePoint ✅")
    file_choice = uploaded_file.name  # Reemplaza la selección con el archivo recién subido

# 4. Previsualizar y editar solo si hay archivo
if file_choice:
    file_bytes = download_file(ctx, library_choice, file_choice)
    if file_bytes:
        try:
            df = pd.read_excel(BytesIO(file_bytes))
            if df.empty:
                st.warning("El archivo está vacío o no tiene datos reconocibles.")
            else:
                st.subheader("Vista previa del archivo")
                st.dataframe(df.head(50), use_container_width=True)

                # -------------------- FILTROS --------------------
                st.subheader("Filtros dinámicos")
                filter_cols = st.multiselect("Selecciona columnas para filtrar", df.columns)
                filtered_df = df.copy()
                for col in filter_cols:
                    valores = st.multiselect(f"Filtrar {col}", df[col].unique())
                    if valores:
                        filtered_df = filtered_df[filtered_df[col].isin(valores)]

                # -------------------- EDICIÓN --------------------
                st.subheader("Editar datos")
                edited_df = st.data_editor(filtered_df, num_rows="dynamic", use_container_width=True)

              # Exportar
                st.subheader("Exportar datos editados")
                def to_excel(df):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                        df.to_excel(writer, index=False, sheet_name="Datos Editados")
                    return output.getvalue()

                excel_data = to_excel(edited_df)
                st.download_button(
                    label="📥 Descargar Excel editado",
                    data=excel_data,
                    file_name=f"editado_{file_choice}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                if st.button("💾 Guardar cambios en SharePoint"):
                    upload_file(ctx, library_choice, file_choice, excel_data)
                    st.success("Archivo actualizado en SharePoint ✅")

        except Exception as e:
            st.error(f"No se pudo leer el Excel: {e}")


