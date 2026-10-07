import streamlit as st
import pypdf
import pandas as pd
import io
import re
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment

st.set_page_config(page_title="Conversor Bancario Universal", page_icon="🏦", layout="wide")
st.title("🏦 Convertidor Universal de Estados de Cuenta")
st.write("Sube el PDF de tu estado de cuenta (Soporta BCP, BBVA, Interbank y Scotiabank) y descárgalo en Excel al instante.")

archivo_pdf = st.file_uploader("Sube tu Estado de Cuenta (PDF)", type=["pdf"])

if archivo_pdf is not None:
    with st.spinner("Procesando documento con Inteligencia de Datos..."):
        try:
            reader = pypdf.PdfReader(archivo_pdf)
            all_text_with_coords = []
            texto_completo = ""
            for page_num, page in enumerate(reader.pages):
                texto_completo += page.extract_text() + " "
                def visitor_extract(text, cm, tm, fontDict, fontSize):
                    x, y = tm[4], tm[5]
                    if text.strip():
                        all_text_with_coords.append((page_num + 1, round(x, 1), round(y, 1), text.strip()))
                page.extract_text(visitor_text=visitor_extract)
                
            lines_by_page_and_y = {}
            for p, x, y, text in all_text_with_coords:
                key = (p, y)
                lines_by_page_and_y.setdefault(key, []).append((x, text))
                
            sorted_keys = sorted(lines_by_page_and_y.keys(), key=lambda k: (k[0], -k[1]))
            
            # 1. DETECCIÓN AUTOMÁTICA DEL BANCO (Orden corregido para evitar conflictos)
            texto_upper = texto_completo.upper()
            banco = "DESCONOCIDO"
            if "ESTADO DE CUENTA CORRIENTE" in texto_upper and "BCP" in texto_upper:
                banco = "BCP_CORRIENTE"
            elif "ESTADO DE CUENTA DE AHORRO" in texto_upper and "BCP" in texto_upper:
                banco = "BCP_AHORROS"
            elif "INTERBANK" in texto_upper and "NEGOCIOS" in texto_upper:
                banco = "INTERBANK"
            elif "SCOTIA" in texto_upper or "20100043140" in texto_upper:
                banco = "SCOTIABANK"
            elif "BBVA" in texto_upper or "CONTIAHORRO" in texto_upper:
                banco = "BBVA"
                
            st.info(f"🏦 Banco detectado automáticamente: **{banco.replace('_', ' ')}**")
            
            # 2. EXTRACCIÓN DEL SALDO INICIAL EXACTO DECLARADO
            saldo_inicial_declarado = None
            
            for key in sorted_keys:
                items = sorted(lines_by_page_and_y[key], key=lambda i: i[0]) 
                combined = " ".join([i[1] for i in items])
                
                if banco == "SCOTIABANK" and "Saldo Final al" in combined and re.search(r'\d{4}', combined):
                    try: saldo_inicial_declarado = float(combined.split()[-1].replace(',', ''))
                    except: pass
                    break
                elif banco in ["BBVA", "BCP_AHORROS"] and "SALDO ANTERIOR" in combined:
                    try: saldo_inicial_declarado = float(combined.split()[-1].replace(',', ''))
                    except: pass
