import io
import json
import base64
import re
import requests
import pandas as pd
import streamlit as st
import pypdfium2 as pdfium
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Configuración inicial de la página
st.set_page_config(
    page_title="Procesador y Reexpresión de Estados Financieros",
    page_icon="📊",
    layout="wide",
)

st.title("ESTADOS FINANCIEROS - REEXPRESIÓN MONETARIA (-6 CEROS)")
st.markdown("*Sistema de Digitalización Inteligente por IA y Generación de Excel Contable*")

# 1. Obtener la API Key desde st.secrets o input manual en la barra lateral
gemini_api_key = st.secrets.get("GEMINI_API_KEY", "")

st.sidebar.header("1. Configuración de API y Documento")
if not gemini_api_key:
    gemini_api_key = st.sidebar.text_input(
        "Ingresa tu Gemini API Key:", 
        type="password",
        help="Obtén tu clave gratis en https://aistudio.google.com"
    )

archivo_subido = st.sidebar.file_uploader(
    "Sube tu Balance de Comprobación en PDF", type=["pdf"]
)

factor_reconversion = st.sidebar.number_input(
    "Factor de división (Eliminación de 6 ceros):",
    min_value=1.0,
    value=1000000.0,
    step=100000.0,
    format="%.2f",
)

# Función segura para limpiar montos numéricos
def a_numero(val):
    if pd.isna(val) or val is None or val == "":
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    s = re.sub(r'[Bs\.\$\€\£\sA-Za-z]', '', s)
    if not s:
        return 0.0
    es_neg = s.startswith('-') or (s.startswith('(') and s.endswith(')'))
    s = s.replace('(', '').replace(')', '').replace('-', '')
    if ',' in s and '.' in s:
        s = s.replace('.', '').replace(',', '.') if s.rfind(',') > s.rfind('.') else s.replace(',', '')
    elif ',' in s:
        s = s.replace(',', '.')
    try:
        n = float(s)
        return -n if es_neg else n
    except ValueError:
        return 0.0

# Función para procesar PDF con pypdfium2 y Gemini Vision API REST
def extraer_datos_con_gemini(pdf_bytes, api_key):
    try:
        # Renderizar PDF a imágenes optimizadas usando pypdfium2 (máximo 8 páginas por seguridad de payload)
        pdf = pdfium.PdfDocument(pdf_bytes)
        images_parts = []
        total_paginas = min(len(pdf), 8)
