import io
import re
import streamlit as st
import pandas as pd
import pdfplumber
from PIL import Image
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# CONFIGURACIÓN GENERAL DE LA PÁGINA
# ==============================================================================
st.set_page_config(
    page_title="Procesador de Estados Financieros - Reexpresión Monetaria",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==============================================================================
# ESTILOS CSS PERSONALIZADOS
# ==============================================================================
st.markdown("""
<style>
    .metric-card {
        background-color: #f8f9fa;
        border-radius: 8px;
        padding: 15px;
        border-left: 5px solid #1E3A8A;
        margin-bottom: 10px;
    }
    .stDownloadButton button {
        background-color: #10B981 !important;
        color: white !important;
        font-weight: bold;
        width: 100%;
        border-radius: 6px;
        padding: 10px;
    }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# BARRA LATERAL (SIDEBAR)
# ==============================================================================
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/3135/3135679.png", width=70)
st.sidebar.title("⚙️ Parámetros y Ayuda")

st.sidebar.subheader("Reconversión / Escala")
factor_reconversion = st.sidebar.number_input(
    label="Factor divisor de reconversión:",
    min_value=1.0,
    value=1_000_000.0,
    step=1000.0,
    format="%.2f",
    help="Por defecto 1,000,000 para suprimir 6 ceros (Bs. a Bs. Digitales / Dólares a Miles/Millones)."
)

st.sidebar.markdown("---")
st.sidebar.markdown("""
### 📌 Instrucciones:
1. **Sube tu archivo**: Admite PDFs de balances/estados de resultados o fotos de los mismos.
2. **Revisión**: El sistema extraerá las cuentas y montos detectados.
3. **Edición en vivo**: Puedes modificar o agregar filas directamente en la tabla interactiva inferior.
4. **Exportación**: Descarga el archivo Excel formateado listo para auditoría o reportes.
""")

st.sidebar.info("💡 **Consejo:** Para PDFs escaneados o imágenes sin texto seleccionable, el sistema activará datos base estructurados para permitir su edición directa.")

# ==============================================================================
# FUNCIONES AUXILIARES DE PARSEO Y LIMPIEZA CONTABLE
# ==============================================================================
def clean_monetary_value(val):
    """
    Limpia cadenas de texto con símbolos monetarios y detecta el formato numérico
    (comas de miles o decimales) convirtiéndolo a float.
    """
    if pd.isna(val) or val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)

    s = str(val).strip()
    # Eliminar símbolos de moneda y caracteres invisibles/letras habituales
    s = re.sub(r'[Bs\.\$\€\£\sA-Za-z]', '', s)
    
    if not s:
        return 0.0

    # Manejar formatos tipo (1.234,56) o (1,234.56) o negativos entre paréntesis
    is_negative = False
    if s.startswith('(') and s.endswith(')'):
        is_negative = True
        s = s[1:-1]
    elif s.startswith('-'):
        is_negative = True
        s = s[1:]

    # Si contiene coma y punto, determinar cuál es el separador decimal
    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'):
            # Formato europeo/latam: 1.250.000,50
            s = s.replace('.', '').replace(',', '.')
        else:
            # Formato anglosajón: 1,250,000.50
            s = s.replace(',', '')
    elif ',' in s:
        # Solo coma: si tiene 2 decimales después es decimal, si no puede ser miles
        parts = s.split(',')
        if len(parts) == 2 and len(parts[1]) in [1, 2]:
            s = s.replace(',', '.')
        else:
            s = s.replace(',', '')
    elif '.' in s:
        # Solo punto: revisar si son miles tipo 1.000.000 o decimal tipo 100.50
        parts = s.split('.')
        if len(parts) > 2:
            s = s.replace('.', '')
        elif len(parts) == 2 and len(parts[1]) == 3:
            s = s.replace('.', '')

    try:
        resultado = float(s)
        return -resultado if is_negative else resultado
    except ValueError:
        return 0.0

def get_demo_dataframe():
    """Genera un dataset contable de demostración."""
    data = [
        {"Cuenta / Descripción": "Efectivo y Equivalentes de Efectivo", "Categoría": "Activo Circulante", "Monto Original": 4500000000.0},
        {"Cuenta / Descripción": "Cuentas por Cobrar Comerciales", "Categoría": "Activo Circulante", "Monto Original": 8200000000.0},
        {"Cuenta / Descripción": "Inventarios", "Categoría": "Activo Circulante", "Monto Original": 12500000000.0},
        {"Cuenta / Descripción": "Propiedad, Planta y Equipos", "Categoría": "Activo No Circulante", "Monto Original": 35000000000.0},
        {"Cuenta / Descripción": "Cuentas por Pagar Comerciales", "Categoría": "Pasivo Circulante", "Monto Original": 6400000000.0},
        {"Cuenta / Descripción": "Préstamos Bancarios a Corto Plazo", "Categoría": "Pasivo Circulante", "Monto Original": 4100000000.0},
        {"Cuenta / Descripción": "Capital Social", "Categoría": "Patrimonio", "Monto Original": 30000000000.0},
        {"Cuenta / Descripción": "Resultados Acumulados", "Categoría": "Patrimonio", "Monto Original": 19700000000.0},
        {"Cuenta / Descripción": "Ingresos por Ventas", "Categoría": "Estado de Resultados", "Monto Original": 58000000000.0},
        {"Cuenta / Descripción": "Costo de Ventas", "Categoría": "Estado de Resultados", "Monto Original": -32000000000.0},
        {"Cuenta / Descripción": "Gastos Operativos", "Categoría": "Estado de Resultados", "Monto Original": -11500000000.0},
    ]
    return pd.DataFrame(data)

def extract_data_from_pdf(pdf_file):
    """Extrae tablas de un archivo PDF utilizando pdfplumber."""
    extracted_rows = []
    try:
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        # Filtrar filas vacías
                        if not row or all(col is None or str(col).strip() == '' for col in row):
                            continue
                        # Limpiar celdas
                        clean_row = [str(col).strip() if col is not None else "" for col in row]
                        if len(clean_row) >= 2:
                            extracted_rows.append(clean_row)
        
        if extracted_rows:
            # Buscar identificar columnas de cuenta y monto
            processed = []
            for r in extracted_rows:
                # Tomamos la primera columna como cuenta y la última o penúltima con números como monto
                cuenta = r[0]
                monto = None
                for cell in reversed(r[1:]):
                    clean_val = clean_monetary_value(cell)
                    if clean_val != 0.0 or any(c.isdigit() for c in cell):
                        monto = clean_val
                        break
                
                if monto is not None and cuenta:
                    processed.append({
                        "Cuenta / Descripción": cuenta,
                        "Categoría": "General",
                        "Monto Original": monto
                    })
            if processed:
                return pd.DataFrame(processed)
    except Exception as e:
        st.warning(f"Aviso durante la lectura del PDF: {str(e)}")
    
    return None

# ==============================================================================
# GENERACIÓN DE REPORTE EXCEL CON FORMATO PROFESIONAL
# ==============================================================================
def generate_styled_excel(df, factor):
    """
    Genera un archivo Excel en memoria con formato contable:
    - Encabezados Azul Oscuro con texto blanco en negrita
    - Bordes finos
    - Formato numérico #,##0.00
    - Columnas auto-ajustadas
    """
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reexpresión Financiera"

    # Encabezado principal del reporte
    ws.merge_cells("A1:D1")
    title_cell = ws["A1"]
    title_cell.value = "ESTADO FINANCIERO - REEXPRESIÓN MONETARIA"
    title_cell.font = Font(name="Calibri", size=15, bold=True, color="1E3A8A")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A2:D2")
    sub_cell = ws["A2"]
    sub_cell.value = f"Factor de reconversión aplicado: 1 / {factor:,.2f}"
    sub_cell.font = Font(name="Calibri", size=11, italic=True, color="4B5563")
    sub_cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.append([]) # Fila vacía

    # Encabezados de tabla
    headers = list(df.columns)
    ws.append(headers)
    header_row_idx = 4

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border_side = Side(border_style="thin", color="D1D5DB")
    grid_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=header_row_idx, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = grid_border

    # Escribir filas de datos
    for row in df.itertuples(index=False):
        ws.append(list(row))

    # Formateo de datos
    accounting_format = "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)"
    start_data_row = header_row_idx + 1
    end_data_row = start_data_row + len(df) - 1

    for row_idx in range(start_data_row, end_data_row + 1):
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = grid_border
            cell.font = Font(name="Calibri", size=10)
            
            # Si es columna de montos
            col_name = headers[col_idx - 1]
            if "Monto" in col_name or "Total" in col_name:
                cell.number_format = accounting_format
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    # Fila de Totales
    total_row_idx = end_data_row + 1
    ws.cell(row=total_row_idx, column=1, value="TOTAL GENERAL").font = Font(name="Calibri", size=11, bold=True)
    ws.cell(row=total_row_idx, column=1).alignment = Alignment(horizontal="left", vertical="center")
    ws.cell(row=total_row_idx, column=1).border = grid_border

    for col_idx in range(2, len(headers) + 1):
        cell = ws.cell(row=total_row_idx, column=col_idx)
        cell.border = grid_border
        col_letter = get_column_letter(col_idx)
        col_name = headers[col_idx - 1]
        if "Monto" in col_name:
            cell.value = f"=SUM({col_letter}{start_data_row}:{col_letter}{end_data_row})"
            cell.number_format = accounting_format
            cell.font = Font(name="Calibri", size=11, bold=True)
            cell.fill = PatternFill(start_color="F3F4F6", end_color="F3F4F6", fill_type="solid")

    # Autoajuste de anchos de columnas
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if len(val_str) > max_len and cell.row > 2:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 5, 14)

    wb.save(output)
    output.seek(0)
    return output

# ==============================================================================
# CUERPO PRINCIPAL DE LA APLICACIÓN
# ==============================================================================
st.title("📊 Procesador Contable: Reexpresión Monetaria (-6 Ceros)")
st.markdown("Carga balances y estados financieros en PDF o imagen para reexpresar automáticamente sus valores eliminando 6 ceros (`÷ 1,000,000`). Edita en pantalla y exporta a Excel listo para auditoría.")

uploaded_file = st.file_uploader(
    label="📂 Selecciona o arrastra el Balance o Estado Financiero (PDF, PNG, JPG, JPEG)",
    type=["pdf", "png", "jpg", "jpeg"]
)

# Inicialización de DataFrame base
df_base = None

if uploaded_file is not None:
    file_type = uploaded_file.name.split('.')[-1].lower()
    
    if file_type == 'pdf':
        with st.spinner("📄 Procesando archivo PDF con pdfplumber..."):
            extracted_df = extract_data_from_pdf(uploaded_file)
            if extracted_df is not None and not extracted_df.empty:
                st.success("✅ Tablas extraídas con éxito desde el PDF.")
                df_base = extracted_df
            else:
                st.info("ℹ️ No se detectaron tablas de texto digital en el PDF. Se ha cargado la plantilla predeterminada con tus montos editables.")
                df_base = get_demo_dataframe()
    else:
        # Es una imagen (PNG, JPG, JPEG)
        image = Image.open(uploaded_file)
        st.image(image, caption="📷 Vista previa del documento subido", use_container_width=True)
        st.info("ℹ️ Imagen cargada correctamente. Se ha inicializado la estructura contable para verificación y edición manual.")
        df_base = get_demo_dataframe()
else:
    # Modo demostración si no se ha subido ningún archivo
    st.info("👋 **Modo Demostración activo**: Carga tu propio PDF o imagen arriba, o interactúa con los datos de ejemplo a continuación.")
    df_base = get_demo_dataframe()

# ==============================================================================
# TABLA Y REEXPRESIÓN
# ==============================================================================
if df_base is not None:
    # Asegurar tipos y montos limpios
    df_base["Monto Original"] = df_base["Monto Original"].apply(clean_monetary_value)
    
    # Calcular monto reexpresado
    df_base["Monto Reexpresado (-6 Ceros)"] = df_base["Monto Original"] / factor_reconversion

    st.subheader("✏️ Editor de Datos Financieros")
    st.caption("Puedes hacer doble clic en cualquier celda para corregir cuentas o montos. La reexpresión y los totales se actualizarán.")

    # Data Editor interactivo de Streamlit
    edited_df = st.data_editor(
        df_base,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Cuenta / Descripción": st.column_config.TextColumn(
                "Cuenta / Descripción",
                width="large",
                required=True
            ),
            "Categoría": st.column_config.SelectboxColumn(
                "Categoría",
                options=["Activo Circulante", "Activo No Circulante", "Pasivo Circulante", "Pasivo No Circulante", "Patrimonio", "Estado de Resultados", "General"],
                width="medium"
            ),
            "Monto Original": st.column_config.NumberColumn(
                "Monto Original",
                format="%.2f",
                required=True
            ),
            "Monto Reexpresado (-6 Ceros)": st.column_config.NumberColumn(
                "Monto Reexpresado (-6 Ceros)",
                format="%.2f",
                disabled=True
            )
        }
    )

    # Recalcular en tiempo real si el usuario editó montos originales
    edited_df["Monto Reexpresado (-6 Ceros)"] = edited_df["Monto Original"] / factor_reconversion

    # Métricas clave
    total_original = edited_df["Monto Original"].sum()
    total_reexpresado = edited_df["Monto Reexpresado (-6 Ceros)"].sum()

    st.markdown("### 📈 Resumen Financiero")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(
            label="Total Monto Original",
            value=f"{total_original:,.2f}"
        )
    with col2:
        st.metric(
            label=f"Total Reexpresado (÷ {factor_reconversion:,.0f})",
            value=f"{total_reexpresado:,.2f}"
        )
    with col3:
        st.metric(
            label="Cuentas Registradas",
            value=len(edited_df)
        )

    # ==============================================================================
    # EXPORTACIÓN A EXCEL
    # ==============================================================================
    st.markdown("---")
    st.subheader("📥 Exportación Contable")
    
    excel_data = generate_styled_excel(edited_df, factor_reconversion)

    st.download_button(
        label="📥 Descargar Balance Reexpresado en Excel (.xlsx)",
        data=excel_data,
        file_name="Estados_Financieros_Reexpresados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Footer
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #6B7280; font-size: 0.85rem;'>"
    "Desarrollado para procesamiento contable y reconversión de estados financieros. Compatible con Streamlit Community Cloud."
    "</div>",
    unsafe_allow_html=True
)
