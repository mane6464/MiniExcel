import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
import io

# 1. Configuración de la página web
st.set_page_config(page_title="Mi Excel Ligero", page_icon="📊", layout="wide")
st.title("📊 Mi Generador de Tablas Ligero (Versión Estable)")
st.write("Modifica la estructura en la barra lateral, edita la tabla sin parpadeos y presiona 'Guardar cambios' para actualizar.")

# 2. Inicializar la tabla en la memoria de la sesión si no existe
if "df_datos" not in st.session_state:
    data_inicial = {
        "Columna A": ["", "", "", "", ""],
        "Columna B": ["", "", "", "", ""],
        "Columna C": ["", "", "", "", ""]
    }
    st.session_state.df_datos = pd.DataFrame(data_inicial)

# 3. Barra lateral para la estructura de la tabla
st.sidebar.header("🛠️ Estructura de la Tabla")

# --- SECCIÓN AÑADIR COLUMNA ---
nueva_columna = st.sidebar.text_input("Nombre de la nueva columna", placeholder="Ej. Teléfono")
if st.sidebar.button("➕ Agregar Columna"):
    if nueva_columna:
        if nueva_columna not in st.session_state.df_datos.columns:
            st.session_state.df_datos[nueva_columna] = [""] * len(st.session_state.df_datos)
            st.rerun()
        else:
            st.sidebar.error("Esa columna ya existe.")
    else:
        st.sidebar.warning("Escribe un nombre para la columna.")

st.sidebar.markdown("---")

# --- SECCIÓN RENOMBRAR COLUMNA ---
st.sidebar.subheader("✏️ Renombrar Columna")
columna_a_cambiar = st.sidebar.selectbox("Selecciona la columna a modificar", st.session_state.df_datos.columns)
nuevo_nombre = st.sidebar.text_input("Nuevo nombre de columna", placeholder="Ej. Cantidad")

if st.sidebar.button("🔄 Cambiar Nombre"):
    if nuevo_nombre:
        if nuevo_nombre not in st.session_state.df_datos.columns:
            st.session_state.df_datos = st.session_state.df_datos.rename(columns={columna_a_cambiar: nuevo_nombre})
            st.success(f"¡Cambiado '{columna_a_cambiar}' por '{nuevo_nombre}'!")
            st.rerun()
        else:
            st.sidebar.error("Ya existe una columna con ese nombre.")
    else:
        st.sidebar.warning("Escribe el nuevo nombre.")


st.sidebar.markdown("---")
st.sidebar.header("🎨 Configuración del Diseño")

# Selector de color
color_elegido = st.sidebar.color_picker("Color de fondo del encabezado", "#365F91")
color_hex = color_elegido.replace("#", "")

# Ajustes de texto
tamano_letra_datos = st.sidebar.slider("Tamaño de letra de los datos", min_value=10, max_value=20, value=12)
alineacion = st.sidebar.selectbox("Alineación del texto", ["Centrado", "Izquierda", "Derecha"])

dict_alineacion = {
    "Centrado": Alignment(horizontal="center", vertical="center"),
    "Izquierda": Alignment(horizontal="left", vertical="center"),
    "Derecha": Alignment(horizontal="right", vertical="center")
}
alineacion_final = dict_alineacion[alineacion]

# 4. Creación de la tabla interactiva DENTRO de un Formulario aislado
st.subheader("📝 Edita tus datos aquí abajo:")

# Al meterlo en un st.form, evitamos que Streamlit recargue la página cada vez que cambias de celda
with st.form("contenedor_tabla"):
    tabla_editada = st.data_editor(
        st.session_state.df_datos, 
        num_rows="dynamic", 
        use_container_width=True
    )
    
    # Este botón procesa de golpe todos los cambios que hayas hecho en las celdas
    boton_guardar = st.form_submit_button("💾 Guardar cambios en la tabla")
    
    if boton_guardar:
        st.session_state.df_datos = tabla_editada
        st.success("¡Datos guardados con éxito en la memoria!")

# 5. Función para procesar la tabla y aplicar los estilos seleccionados
def generar_excel(dataframe, color_bg, tamano_fuente, alineacion_obj):
    wb = Workbook()
    ws = wb.active
    ws.title = "Datos"

    columnas = list(dataframe.columns)
    ws.append(columnas)

    for index, fila in dataframe.iterrows():
        ws.append(list(fila))

    fill_encabezado = PatternFill(start_color=color_bg, end_color=color_bg, fill_type="solid")
    fuente_encabezado = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    fuente_datos = Font(name="Arial", size=tamano_fuente, bold=False)

    for num_fila, fila_celdas in enumerate(ws.iter_rows(min_row=1, max_row=len(dataframe)+1, min_col=1, max_col=len(columnas)), start=1):
        for celda in fila_celdas:
            celda.alignment = alineacion_obj
            if num_fila == 1:
                celda.fill = fill_encabezado
                celda.font = fuente_encabezado
            else:
                celda.font = fuente_datos

    # Autoajustar el ancho de las columnas
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for celda in col:
            if celda.value is not None:
                len_texto = len(str(celda.value))
                if len_texto > max_len:
                    max_len = len_texto
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# 6. Botón de descarga (fuera del formulario)
st.markdown("---")
if st.button("🚀 Generar y preparar descarga de Excel"):
    archivo_excel = generar_excel(st.session_state.df_datos, color_hex, tamano_letra_datos, alineacion_final)
    
    st.download_button(
        label="📥 Descargar archivo .xlsx",
        data=archivo_excel,
        file_name="mi_tabla_ligera.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.success("¡Tu archivo está listo para descargar!")
