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
