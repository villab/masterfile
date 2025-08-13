import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText

st.title("📋 Masterfile Sutel")

# Simular carga de datos (reemplaza por tu fuente real)
df = pd.DataFrame({
    "ID": [1, 2, 3],
    "Cliente": ["Cliente A", "Cliente B", "Cliente C"],
    "Estado": ["Activo", "Inactivo", "Activo"]
})

# Guardar copia original antes de edición
df_original = df.copy()

# Mostrar editor en Streamlit (activando edición)
df_editado = st.data_editor(df, num_rows="dynamic")

if st.button("💾 Guardar nueva versión de Masterfile"):
    try:
        # Resetear índice para evitar errores
        df_original_reset = df_original.reset_index(drop=True)
        df_editado_reset = df_editado.reset_index(drop=True)

        # Comparar cambios
        cambios = df_original_reset.compare(
            df_editado_reset,
            keep_shape=True,
            keep_equal=False
        ).dropna(how="all")

        if cambios.empty:
            st.info("No se detectaron cambios.")
        else:
            # Obtener nombre de la columna 2
            col2 = df_editado_reset.columns[1]

            # Construir cuerpo del correo
            cuerpo_mensaje = f"Se modificaron las siguientes filas (columna '{col2}'):\n\n"
            for idx in cambios.index.unique():
                valor_celda = df_editado_reset.loc[idx, col2]
                cambios_fila = cambios.loc[idx]
                # Si es una sola columna modificada
                if isinstance(cambios_fila, pd.Series):
                    cambios_fila = cambios_fila.to_frame().T
                for col in cambios_fila.columns.levels[0]:
                    antes = cambios_fila[col]["self"]
                    despues = cambios_fila[col]["other"]
                    cuerpo_mensaje += f"• {valor_celda}: '{antes}' → '{despues}' (columna: {col})\n"

            # Configuración SMTP desde secrets
            smtp_server = st.secrets["smtp_server"]
            smtp_port = st.secrets["smtp_port"]
            smtp_user = st.secrets["smtp_user"]
            smtp_pass = st.secrets["smtp_pass"]
            email_from = st.secrets["email_from"]
            email_to = st.secrets["email_to"].split(",")

            # Crear y enviar correo
            msg = MIMEText(cuerpo_mensaje, "plain", "utf-8")
            msg["Subject"] = "Cambios en MasterfileSutel"
            msg["From"] = email_from
            msg["To"] = ", ".join(email_to)

            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(smtp_user, smtp_pass)
                server.sendmail(email_from, email_to, msg.as_string())

            st.success("✅ Archivo guardado y correo enviado correctamente.")

    except Exception as e:
        st.error(f"⚠️ Error al enviar el correo: {e}")
