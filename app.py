import io
import os
import re
import json
import base64
import requests
import streamlit as st
import pandas as pd
import pdfplumber
import pypdfium2 as pdfium
from PIL import Image
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# CONFIGURACIÓN GENERAL DE LA PÁGINA
# ==============================================================================
st.set_page_config(
    page_title="Procesador y Reexpresión de Estados Financieros",
    page_icon="📊",
    layout="wide",
)

st.title("ESTADOS FINANCIEROS - REEXPRESIÓN MONETARIA (-6 CEROS)")
st.markdown("*Sistema Automatizado con IA (Gemini Vision OCR) para Balances y Estados Financieros Escaneados*")

# ==============================================================================
# SIDEBAR: CONFIGURACIÓN Y CLAVE DE GEMINI
# ==============================================================================
st.sidebar.header("⚙️ 1. Configuración de IA y Reconversión")

api_key_default = os.environ.get("GEMINI_API_KEY", "")
if not api_key_default and "GEMINI_API_KEY" in st.secrets:
    api_key_default = st.secrets["GEMINI_API_KEY"]

gemini_api_key = st.sidebar.text_input(
    "Clave Gemini API (Google AI Studio):",
    value=api_key_default,
    type="password",
    help="Consigue tu API Key gratis en aistudio.google.com para procesar escaneos con IA."
)

factor_reconversion = st.sidebar.number_input(
    "Factor de división (Eliminación de 6 ceros):",
    min_value=1.0,
    value=1000000.0,
    step=100000.0,
    format="%.2f",
)

st.sidebar.markdown("---")
st.sidebar.header("📂 2. Cargar Documento")
archivo_subido = st.sidebar.file_uploader(
    "Sube tu Balance o Estado Financiero (PDF escaneado, Imagen, Excel o CSV):", 
    type=["pdf", "png", "jpg", "jpeg", "xlsx", "xls", "csv"]
)

# ==============================================================================
# FUNCIONES DE EXTRACCIÓN Y LIMPIEZA CONTABLE
# ==============================================================================
def limpiar_monto_contable(val):
    """Convierte cadenas con símbolos monetarios, comas, puntos o paréntesis a float."""
    if pd.isna(val) or val is None or val == "":
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    
    s = str(val).strip()
    s = re.sub(r'[Bs\.\$\€\£\sA-Za-z]', '', s)
    if not s:
        return 0.0
    
    es_negativo = False
    if (s.startswith('(') and s.endswith(')')) or s.startswith('-'):
        es_negativo = True
        s = s.replace('(', '').replace(')', '').replace('-', '')
        
    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif ',' in s:
        partes = s.split(',')
        if len(partes) == 2 and len(partes[1]) in [1, 2]:
            s = s.replace(',', '.')
        else:
            s = s.replace(',', '')
    elif '.' in s:
        partes = s.split('.')
        if len(partes) > 2 or (len(partes) == 2 and len(partes[1]) == 3):
            s = s.replace('.', '')
            
    try:
        num = float(s)
        return -num if es_negativo else num
    except ValueError:
        return 0.0

def convertir_pdf_a_imagenes(archivo_bytes, max_paginas=5):
    """Convierte las páginas del PDF en imágenes PIL para enviar a Gemini Vision."""
    pdf = pdfium.PdfDocument(archivo_bytes)
    imagenes = []
    total_pags = min(len(pdf), max_paginas)
    for i in range(total_pags):
        pagina = pdf[i]
        bitmap = pagina.render(scale=2.0)
        img = bitmap.to_pil()
        imagenes.append(img)
    return imagenes

def extraer_con_gemini_vision(imagenes, api_key):
    """
    Envía las imágenes a la API REST de Gemini (sin depender de SDKs externos)
    para extraer y estructurar el Balance de Comprobación en JSON exacto.
    """
    prompt = """
    Eres un Contador Público y Auditor Senior experto en digitalización de estados financieros escaneados.
    Analiza meticulosamente las imágenes adjuntas del Balance de Comprobación y/o Estados Financieros.
    
    Extrae TODAS las filas de cuentas contables, códigos, nombres y montos con máxima fidelidad.
    Devuelve ÚNICAMENTE un array JSON válido con la siguiente estructura para cada fila encontrada:
    [
      {
        "Cuenta / Descripción": "Código y nombre de la cuenta (ej. 1.1.1.00 - CAJA Y BANCO)",
        "Saldo Inicial": número o 0 si no existe,
        "Monto Debe": número o 0 si no existe,
        "Monto Haber": número o 0 si no existe,
        "Saldo Actual": número con saldo final (positivo o negativo)
      }
    ]
    
    Reglas estrictas:
    1. Si una columna no aparece en el documento, pon 0.
    2. Los montos deben ser números numéricos (float o int), respetando el valor original escaneado.
    3. Si detectas totales o subtotales de grupo (ej: TOTAL ACTIVO, TOTAL PASIVO), inclúyelos claramente.
    4. Responde SOLO con el bloque JSON sin texto adicional ni bloques markdown.
    """
    
    parts = [{"text": prompt}]
    for img in imagenes:
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=85)
        img_b64 = base64.b64encode(buf.getvalue()).decode("utf-8")
        parts.append({
            "inline_data": {
                "mime_type": "image/jpeg",
                "data": img_b64
            }
        })
    
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
    payload = {
        "contents": [{"parts": parts}],
        "generationConfig": {
            "response_mime_type": "application/json",
            "temperature": 0.1
        }
    }
    
    resp = requests.post(url, json=payload, timeout=90)
    if resp.status_code != 200:
        raise ValueError(f"Error de Gemini API ({resp.status_code}): {resp.text}")
        
    resultado_json = resp.json()
    try:
        texto_generado = resultado_json["candidates"][0]["content"]["parts"][0]["text"].strip()
        datos_tabla = json.loads(texto_generado)
        return pd.DataFrame(datos_tabla)
    except Exception as e:
        # Fallback de limpieza por si incluye texto envolvente
        raw_text = resp.text
        match = re.search(r'\[.*\]', raw_text, re.DOTALL)
        if match:
            return pd.DataFrame(json.loads(match.group(0)))
        raise ValueError(f"No se pudo estructurar el JSON de la respuesta: {e}")

# ==============================================================================
# TABLAS BASE DE RESPALDO (POR SI NO HAY ARCHIVO SUBIDO)
# ==============================================================================
def obtener_tablas_predeterminadas():
    bc_data = [
        ["1-ACTIVO CIRCULANTE", 0.0, 0.0, 0.0, 0.0],
        ["1.1.1.00 - CAJA Y BANCO", 33.35, 6787.98, 6821.33, 0.00],
        ["1.1.1.05 - BANCO CORPBANCA", 0.22, 0.00, 0.22, 0.00],
        ["1.1.1.08 - BANCO NACIONAL DE CRÉDITO", 0.25, 0.00, 0.25, 0.00],
        ["1.1.2.50 - CTAS X COBRAR EMPLEADOS", 0.00, 0.23, 0.00, 0.23],
        ["1.1.3.00 - INVENTARIO DE MERCANCÍAS", 3.51, 1.42, 3.51, 1.42],
        ["1.1.5.01 - ISLR PREPAGADO", 0.00, 0.21, 0.00, 0.21],
        ["TOTAL ACTIVO CIRCULANTE", 37.33, 6789.41, 1422.18, 6825.30],
        ["2.1.2.10 - CTAS X PAGAR PROVEEDORES", 0.00, 25.03, 1051.40, -1025.35],
        ["2.1.4.01 - IVA POR PAGAR", -2696.93, 474.23, 471.35, -14.69],
        ["TOTAL PASIVO A CORTO PLAZO", -2696.93, 5178.85, 5176.94, -1041.05],
        ["3.1.0 - CAPITAL SOCIAL ACCIONISTAS", -300.00, 299.99, 0.00, -3.00],
        ["3.1.2 - RESERVA LEGAL", -30.00, 29.99, 0.00, -0.30],
        ["3.1.3 - RESULTADOS ACUMULADOS", -27457.30, 27457.02, 0.00, -274.57],
        ["TOTAL PATRIMONIO", -34563.24, 34562.96, 380.65, -381.13]
    ]
    df_bc = pd.DataFrame(bc_data, columns=["Cuenta / Descripción", "Saldo Inicial", "Monto Debe", "Monto Haber", "Saldo Actual"])
    return df_bc

# ==============================================================================
# LÓGICA DE PROCESAMIENTO
# ==============================================================================
df_bc = obtener_tablas_predeterminadas()
es_extraido_con_ia = False

if archivo_subido is not None:
    nombre = archivo_subido.name.lower()
    bytes_archivo = archivo_subido.getvalue()
    
    if nombre.endswith(".pdf"):
        if gemini_api_key:
            with st.spinner("🤖 Analizando y transcribiendo PDF escaneado con Gemini Vision OCR..."):
                try:
                    imagenes = convertir_pdf_a_imagenes(bytes_archivo)
                    df_gemini = extraer_con_gemini_vision(imagenes, gemini_api_key)
                    if df_gemini is not None and not df_gemini.empty:
                        df_bc = df_gemini
                        es_extraido_con_ia = True
                        st.sidebar.success(f"✨ ¡OCR completado! Se digitalizaron {len(df_bc)} cuentas del documento escaneado.")
                except Exception as e:
                    st.sidebar.error(f"Error con Gemini OCR: {e}")
        else:
            st.sidebar.warning("⚠️ Este es un documento escaneado. Ingresa tu API Key de Gemini en la barra lateral para leerlo automáticamente con IA.")
            try:
                with pdfplumber.open(archivo_subido) as pdf:
                    filas = []
                    for p in pdf.pages:
                        tb = p.extract_tables()
                        for t in tb:
                            for f in t:
                                if f and any(f):
                                    filas.append([str(c) if c else "" for c in f])
                    if len(filas) > 3:
                        df_bc = pd.DataFrame(filas[1:], columns=filas[0])
                        st.sidebar.success("Se extrajeron tablas digitales.")
            except Exception:
                pass
                
    elif nombre.endswith((".png", ".jpg", ".jpeg")):
        if gemini_api_key:
            with st.spinner("🤖 Leyendo imagen con Gemini Vision OCR..."):
                try:
                    img = Image.open(io.BytesIO(bytes_archivo))
                    df_gemini = extraer_con_gemini_vision([img], gemini_api_key)
                    if df_gemini is not None and not df_gemini.empty:
                        df_bc = df_gemini
                        es_extraido_con_ia = True
                        st.sidebar.success("✨ ¡Imagen leída correctamente!")
                except Exception as e:
                    st.sidebar.error(f"Error procesando imagen: {e}")
        else:
            st.sidebar.warning("⚠️ Ingresa tu API Key de Gemini en la barra lateral para leer la imagen escaneada.")

    elif nombre.endswith((".xlsx", ".xls")):
        try:
            excel_cargado = pd.read_excel(archivo_subido, sheet_name=None)
            hojas = list(excel_cargado.keys())
            st.sidebar.success(f"¡Excel cargado! Hojas encontradas: {', '.join(hojas)}")
            df_bc = excel_cargado[hojas[0]].dropna(how="all")
        except Exception as e:
            st.sidebar.error(f"Error leyendo Excel: {e}")

# ==============================================================================
# REEXPRESIÓN MONETARIA SEGURA (-6 CEROS)
# ==============================================================================
def reexpresar_dataframe(df, factor):
    df_copy = df.copy()
    for col in df_copy.columns:
        if col != df_copy.columns[0]:
            df_copy[col] = df_copy[col].apply(limpiar_monto_contable)
            df_copy[col] = df_copy[col] / factor
    return df_copy

df_bc_reexpresado = reexpresar_dataframe(df_bc, factor_reconversion)

# ==============================================================================
# INTERFAZ Y PESTAÑAS
# ==============================================================================
if es_extraido_con_ia:
    st.success("✅ **Documento escaneado digitalizado con éxito por Gemini.** Revisa las cifras y descarga el Excel listo para reimpresión.")

tab1, tab2 = st.tabs([
    "📋 Balance de Comprobación Reexpresado", 
    "✏️ Editor Interactivo / Corrección Manual"
])

with tab1:
    st.subheader(f"Balance Reexpresado (Valores ÷ {factor_reconversion:,.0f})")
    st.dataframe(df_bc_reexpresado, use_container_width=True)

with tab2:
    st.subheader("Modificar Cuentas o Montos en Pantalla")
    st.caption("Puedes corregir cualquier celda haciendo doble clic:")
    df_bc_editado = st.data_editor(df_bc, num_rows="dynamic", use_container_width=True)

# ==============================================================================
# EXPORTACIÓN A EXCEL LIMPIO Y FORMATEADO
# ==============================================================================
def generar_excel_impresion(df, factor):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Balance Reexpresado"

    # Encabezado institucional
    ws.merge_cells("A1:E1")
    t1 = ws["A1"]
    t1.value = "ESTADO FINANCIERO - BALANCE DE COMPROBACIÓN"
    t1.font = Font(name="Calibri", size=14, bold=True, color="1E3A8A")
    t1.alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A2:E2")
    t2 = ws["A2"]
    t2.value = f"Valores reexpresados aplicando factor de división: 1 / {factor:,.2f}"
    t2.font = Font(name="Calibri", size=10, italic=True, color="4B5563")
    t2.alignment = Alignment(horizontal="center", vertical="center")

    ws.append([]) # Espacio en blanco

    headers = list(df.columns)
    ws.append(headers)
    fila_h = 4

    h_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    h_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    borde_fino = Border(
        left=Side(style="thin", color="D1D5DB"),
        right=Side(style="thin", color="D1D5DB"),
        top=Side(style="thin", color="D1D5DB"),
        bottom=Side(style="thin", color="D1D5DB")
    )

    for c_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=fila_h, column=c_idx)
        cell.font = h_font
        cell.fill = h_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = borde_fino

    # Formato contable oficial
    formato_contable = "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)"
    for row in df.itertuples(index=False):
        ws.append(list(row))

    inicio_d = fila_h + 1
    fin_d = inicio_d + len(df) - 1

    for r in range(inicio_d, fin_d + 1):
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = borde_fino
            cell.font = Font(name="Calibri", size=9)
            if c > 1:
                cell.number_format = formato_contable
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    # Autoajuste de anchos de columna para impresión limpia
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_l = max(len(str(c.value or '')) for c in col if c.row >= fila_h)
        ws.column_dimensions[col_letter].width = max(max_l + 4, 15)

    wb.save(output)
    return output.getvalue()

st.markdown("---")
excel_descargable = generar_excel_impresion(df_bc_reexpresado, factor_reconversion)

st.download_button(
    label="📥 Descargar Balance Reexpresado en Excel (.xlsx)",
    data=excel_descargable,
    file_name="Balance_Reexpresado_Reimpresion.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
= formato_contable
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
    # Autoajuste de anchos de columna para impresión limpia
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_l = max(len(str(c.value or '')) for c in col if c.row >= fila_h)
        ws.column_dimensions[col_letter].width = max(max_l + 4, 15)
    wb.save(output)
    return output.getvalue()
st.markdown("---")
excel_descargable = generar_excel_impresion(df_bc_reexpresado, factor_reconversion)
st.download_button(
    label="📥 Descargar Balance Reexpresado en Excel (.xlsx)",
    data=excel_descargable,
    file_name="Balance_Reexpresado_Reimpresion.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
