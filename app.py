import io
import pandas as pd
import streamlit as st

# Configuración inicial de la página
st.set_page_config(
    page_title="Procesador y Reexpresión de Estados Financieros",
    page_icon="📊",
    layout="wide",
)

st.title("ESTADO FINANCIERO - REEXPRESIÓN MONETARIA")
st.markdown(
    "*Factor de reconversión aplicado: 1 / 1,000,000.00 (Eliminación de 6"
    " ceros)*"
)

# Datos reales extraídos del Balance de Comprobación (Trattoria-Pizzeria Da Lucy, C.A.)
datos_iniciales = [
    {
        "Cuenta / Descripción": (
            "Efectivo y Equivalentes de Efectivo (Caja y Bancos)"
        ),
        "Categoría": "Activo Circulante",
        "Monto Original": 6621552093.99,
    },
    {
        "Cuenta / Descripción": "Cuentas por Cobrar Comerciales",
        "Categoría": "Activo Circulante",
        "Monto Original": 6621798283.38,
    },
    {
        "Cuenta / Descripción": "Inventarios",
        "Categoría": "Activo Circulante",
        "Monto Original": 1425807.49,
    },
    {
        "Cuenta / Descripción": "Propiedad, Planta y Equipos (Maquinarias y Muebles)",
        "Categoría": "Activo No Circulante",
        "Monto Original": 3615457.65,
    },
    {
        "Cuenta / Descripción": "Cuentas por Pagar Comerciales",
        "Categoría": "Pasivo Circulante",
        "Monto Original": 4494460472.00,
    },
    {
        "Cuenta / Descripción": "Impuestos y Retenciones por Pagar",
        "Categoría": "Pasivo Circulante",
        "Monto Original": 471349054.12,
    },
    {
        "Cuenta / Descripción": "Capital Social y Actualización",
        "Categoría": "Patrimonio",
        "Monto Original": 381131.23,
    },
    {
        "Cuenta / Descripción": "Resultados Acumulados",
        "Categoría": "Patrimonio",
        "Monto Original": 274.57,
    },
    {
        "Cuenta / Descripción": "Ingresos por Ventas (Sector Privado)",
        "Categoría": "Estado de Resultados",
        "Monto Original": 1872384718.69,
    },
    {
        "Cuenta / Descripción": "Costo de Ventas",
        "Categoría": "Estado de Resultados",
        "Monto Original": -5965501128.19,
    },
    {
        "Cuenta / Descripción": "Gastos Operativos y Generales",
        "Categoría": "Estado de Resultados",
        "Monto Original": -231552001.76,
    },
]

df_base = pd.DataFrame(datos_iniciales)

# Sidebar de control
st.sidebar.header("Parámetros de Reexpresión")
factor_reconversion = st.sidebar.number_input(
    "Factor de división (Eliminación de ceros):",
    value=1000000.0,
    step=100000.0,
    format="%.2f",
)

st.sidebar.info(
    "💡 Puedes editar directamente las celdas de la tabla para ajustar montos"
    " o descripciones según tus requerimientos."
)

# Editor interactivo para mantener flexibilidad total idéntica al Excel deseado
st.subheader("Edición y Validación de Cuentas")
df_editado = st.data_editor(
    df_base,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "Monto Original": st.column_config.NumberColumn(
            "Monto Original", format="Bs. %.2f"
        )
    },
)

# Cálculo automático de la reconversión monetaria (-6 ceros)
df_editado["Monto Reexpresado (-6 Ceros)"] = (
    df_editado["Monto Original"] / factor_reconversion
)

# Mostrar la tabla final formateada de manera idéntica al modelo que necesitas
st.markdown("---")
st.subheader("Vista Previa del Resultado Reexpresado")


def formatear_monto(val):
    try:
        # Formato contable con separador de miles y paréntesis para negativos
        if val < 0:
            return f"({abs(val):,.2f})".replace(",", "X").replace(".", ",").replace("X", ".")
        else:
            return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return val


df_visual = pd.DataFrame()
df_visual["Cuenta / Descripción"] = df_editado["Cuenta / Descripción"]
df_visual["Categoría"] = df_editado["Categoría"]
df_visual["Monto Original"] = df_editado["Monto Original"].apply(
    formatear_monto
)
df_visual["Monto Reexpresado (-6 Ceros)"] = df_editado[
    "Monto Reexpresado (-6 Ceros)"
].apply(formatear_monto)

st.dataframe(df_visual, use_container_width=True)

# Cálculo de Totales
total_original = df_editado["Monto Original"].sum()
total_reexpresado = df_editado["Monto Reexpresado (-6 Ceros)"].sum()

col1, col2 = st.columns(2)
with col1:
    st.metric(
        label="TOTAL GENERAL (Original)",
        value=f"Bs. {total_original:,.2f}".replace(",", "X")
        .replace(".", ",")
        .replace("X", "."),
    )
with col2:
    st.metric(
        label="TOTAL GENERAL (Reexpresado)",
        value=f"Bs. {total_reexpresado:,.2f}".replace(",", "X")
        .replace(".", ",")
        .replace("X", "."),
    )


# Función para generar el archivo Excel limpio con openpyxl
def generar_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Estado Financiero")
    processed_data = output.getvalue()
    return processed_data


st.markdown("---")
excel_data = generar_excel(df_editado)

st.download_button(
    label="📥 Descargar Estado Financiero en Excel Limpio (.xlsx)",
    data=excel_data,
    file_name="Estado_Financiero_Reexpresado.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

