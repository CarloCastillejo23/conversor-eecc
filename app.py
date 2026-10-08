import streamlit as st
import pypdf
import pandas as pd
import io
import re
import datetime
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment

st.set_page_config(page_title="Respaldo Tributario - Conversor Bancario", page_icon="📈", layout="wide")

# --- INYECCIÓN DE DISEÑO CORPORATIVO (CSS) ---
st.markdown("""
<style>
/* Aplicar tipografía oficial de la marca */
html, body, [class*="css"] {
    font-family: 'Bahnschrift', sans-serif !important;
}

/* Fondo principal */
.stApp {
    background-color: #F8F9FA;
}

/* Color de los Títulos */
h1, h2, h3 {
    color: #2E1E7E !important; 
    font-weight: bold !important;
}

/* Estilo para los botones principales (Verde Menta a Azul Oscuro) */
.stButton>button {
    background-color: #66CCA1 !important;
    color: #2E1E7E !important;
    font-weight: bold !important;
    font-size: 16px !important;
    border-radius: 8px !important;
    border: 2px solid #66CCA1 !important;
    transition: all 0.3s ease;
    width: 100%;
}
.stButton>button:hover {
    background-color: #2E1E7E !important;
    color: #FFFFFF !important;
    border: 2px solid #2E1E7E !important;
}

/* Zona de carga de archivos (Dropzone) */
[data-testid="stFileUploadDropzone"] {
    border: 2px dashed #2E1E7E !important;
    background-color: rgba(46, 30, 126, 0.05) !important;
    border-radius: 10px !important;
}

/* Cajas de información (st.info, st.success) */
.stAlert {
    border-left-color: #66CCA1 !important;
    background-color: #FFFFFF !important;
    box-shadow: 0px 4px 6px rgba(0,0,0,0.05);
    color: #2E1E7E !important;
}

/* Cajas de texto (Contraseña) */
.stTextInput>div>div>input {
    border: 1px solid #2E1E7E !important;
    border-radius: 6px !important;
}
</style>
""", unsafe_allow_html=True)

# --- CABECERA DE LA APLICACIÓN ---
st.markdown("<h1>📈 Convertidor Universal de Estados de Cuenta</h1>", unsafe_allow_html=True)
st.markdown("<p style='color: #2E1E7E; font-size: 18px;'>Sube el PDF de tu estado de cuenta (Soporta BCP, BBVA, Interbank, Scotiabank y BanBif) y descárgalo en Excel corporativo al instante.</p>", unsafe_allow_html=True)
st.write("---")

col1, col2 = st.columns([2, 1])
with col1:
    archivo_pdf = st.file_uploader("Sube tu Estado de Cuenta (PDF)", type=["pdf"])
with col2:
    st.write(" ")
    st.write(" ")
    password_pdf = st.text_input("🔑 Contraseña del PDF (Solo si está protegido)", type="password", help="Normalmente es tu RUC o DNI")

if archivo_pdf is not None:
    with st.spinner("Procesando documento con Inteligencia de Datos..."):
        try:
            reader = pypdf.PdfReader(archivo_pdf)
            
            if reader.is_encrypted:
                if not password_pdf:
                    st.warning("🔒 Este PDF está protegido con contraseña. Por favor, escribe la clave en la casilla superior y presiona Enter.")
                    st.stop()
                else:
                    decrypted = reader.decrypt(password_pdf)
                    if decrypted == 0:
                        st.error("❌ La contraseña ingresada es incorrecta. Inténtalo de nuevo.")
                        st.stop()
                    else:
                        st.success("🔓 PDF desbloqueado correctamente.")
            
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
                if key not in lines_by_page_and_y:
                    lines_by_page_and_y[key] = []
                lines_by_page_and_y[key].append((x, text))
                
            sorted_keys = sorted(lines_by_page_and_y.keys(), key=lambda k: (k[0], -k[1]))
            
            # 1. DETECCIÓN AUTOMÁTICA DEL BANCO
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
            elif "BANBIF" in texto_upper:
                banco = "BANBIF"
                
            st.info(f"🏦 Banco detectado automáticamente: **{banco.replace('_', ' ')}**")
            
            current_year = str(datetime.datetime.now().year)
            match_year = re.search(r'\b(202\d)\b', texto_completo)
            doc_year = match_year.group(1) if match_year else current_year
            
            # 2. EXTRACCIÓN DEL SALDO INICIAL EXACTO DECLARADO
            saldo_inicial_declarado = None
            for key in sorted_keys:
                items = sorted(lines_by_page_and_y[key], key=lambda i: i[0]) 
                combined = " ".join([i[1] for i in items])
                
                if banco == "SCOTIABANK" and "Saldo Final al" in combined and re.search(r'\d{4}', combined):
                    try: 
                        saldo_inicial_declarado = float(combined.split()[-1].replace(',', ''))
                    except: pass
                    break
                elif banco in ["BBVA", "BCP_AHORROS"] and "SALDO ANTERIOR" in combined:
                    try: 
                        saldo_inicial_declarado = float(combined.split()[-1].replace(',', ''))
                    except: pass
                    break
                elif banco == "INTERBANK":
                    if len(items) > 3 and not re.search(r'[a-zA-Z]', combined) and items[0][1].replace(',', '').replace('.', '').isdigit():
                        try:
                            primer_texto = items[0][1]
                            match = re.search(r'(\d+,\d+\.\d{2}|\d+\.\d{2})', primer_texto)
                            if match:
                                numeros = re.findall(r'\d{1,3}(?:,\d{3})*\.\d{2}', primer_texto)
                                if len(numeros) > 1:
                                    saldo_inicial_declarado = float(numeros[1].replace(',', '')) 
                                else:
                                    saldo_inicial_declarado = float(numeros[0].replace(',', ''))
                        except: pass
                        break
                elif banco == "BCP_CORRIENTE":
                    if len(items) > 5 and not re.search(r'[a-zA-Z]', combined):
                        try: 
                            saldo_inicial_declarado = float(items[0][1].replace(',', ''))
                        except: pass
                        break

            saldo_previo = saldo_inicial_declarado if saldo_inicial_declarado is not None else 0.0
            data_final = []
            
            # 3. PROCESAMIENTO DE TRANSACCIONES
            for key in sorted_keys:
                items = sorted(lines_by_page_and_y[key], key=lambda i: i[0]) 
                combined = " ".join([i[1] for i in items])
                pagina = key[0]
                
                fecha, desc, medio, lugar, sucursal, num_op, hora, cargo, abono, itf, saldo = [""]*11
                es_transaccion = False

                if banco == "BANBIF":
                    if re.match(r'^\d{2}/\d{2}/\d{2}', combined):
                        for x, t in items:
                            if x < 100: fecha = t
                            elif 100 <= x < 150: pass
                            elif 150 <= x < 380: desc += t + " "
                            elif 380 <= x < 450: cargo = t.replace(',', '')
                            elif 450 <= x < 510: abono = t.replace(',', '')
                            elif x >= 510: saldo = t.replace(',', '')
                        desc = desc.strip()
                        if "ITF" in desc.upper():
                            itf = cargo if cargo else abono
                            cargo = ""
                            abono = ""
                        numeros = re.findall(r'\b\d{6,10}\b', desc)
                        if numeros: num_op = numeros[-1]
                        es_transaccion = True

                elif banco == "SCOTIABANK":
                    if re.match(r'^\d{2}/\d{2}\s', combined):
                        tokens = combined.split()
                        fecha = tokens[0] 
                        saldo = tokens[-1]
                        monto_raw = tokens[-2]
                        try:
                            saldo_float = float(saldo.replace(',', ''))
                            monto_float = float(monto_raw.replace(',', ''))
                            if round(saldo_previo - monto_float, 2) == round(saldo_float, 2): cargo = monto_raw
                            elif round(saldo_previo + monto_float, 2) == round(saldo_float, 2): abono = monto_raw
                            saldo_previo = saldo_float
                        except: cargo = monto_raw
                        num_op = tokens[-3] 
                        medio = tokens[2] 
                        desc = " ".join(tokens[3:-3]) 
                        es_transaccion = True

                elif banco == "INTERBANK":
                    if re.match(r'^\d{2}/\d{2}\s+\d{2}/\d{2}', combined):
                        tokens = combined.split()
                        fecha = tokens[0]
                        saldo = tokens[-1]
                        monto_raw = tokens[-2]
                        cargo = monto_raw.replace('-', '') if '-' in monto_raw else ""
                        abono = monto_raw if '-' not in monto_raw else ""
                        num_op_match = re.search(r'\b\d{7}\b', combined)
                        if num_op_match: num_op = num_op_match.group(0)
                        middle_tokens = tokens[2:-2]
                        if "WEB" in middle_tokens: medio = "WEB"
                        elif "INTERNO" in middle_tokens: medio = "INTERNO"
                        desc_tokens = [t for t in middle_tokens if t != num_op and t != medio]
                        desc = " ".join(desc_tokens)
                        es_transaccion = True

                elif banco == "BCP_AHORROS":
                    if re.match(r'^\d{2}[A-Z]{3}\s', combined):
                        raw_str = items[0][1] if len(items)==1 else combined
                        tokens = raw_str.split()
                        fecha = tokens[0]
                        monto = tokens[-1]
                        idx = raw_str.rfind(monto)
                        if idx > 60: abono = monto
                        else: cargo = monto
                        inicio_desc = 2 if re.match(r'^\d{2}[A-Z]{3}$', tokens[1]) else 1
                        desc = " ".join(tokens[inicio_desc:-1])
                        es_transaccion = True

                elif banco == "BBVA":
                    if len(items) > 3 and re.match(r'^\d{2}-\d{2}$', items[0][1]):
                        for x, t in items:
                            if x < 60: fecha = t
                            elif 100 < x < 250: desc += t + " "
                            elif 250 < x < 310: sucursal = t 
                            elif 310 < x < 330: medio = t 
                            elif 330 < x < 380: num_op = t 
                            elif 380 < x < 430:
                                if '-' in t: cargo = t.replace('-', '')
                                else: abono = t
                            elif 430 < x < 490: itf = t 
                            elif 490 < x < 550: saldo = t
                        desc = desc.strip()
                        es_transaccion = True

                elif banco == "BCP_CORRIENTE":
                    if re.match(r'^\d{2}-\d{2}\s', combined):
                        tokens = combined.split()
                        fecha = tokens[0]
                        saldo = tokens[-1]
                        monto_raw = tokens[-2]
                        cargo = monto_raw.replace('-', '') if '-' in monto_raw else ""
                        abono = monto_raw if '-' not in monto_raw else ""
                        middle = tokens[1:-2] 
                        medios_conocidos = ["BPI", "POS", "VEN", "INT", "CAJ", "TLC", "BPT"]
                        medio_index = next((i for i, t in enumerate(middle) if t in medios_conocidos), -1)
                        if medio_index != -1:
                            desc = " ".join(middle[:medio_index])
                            rest = middle[medio_index+1:]
                            medio = middle[medio_index]
                        else:
                            desc = " ".join(middle[:3]) 
                            rest = middle[3:]
                        for t in rest:
                            if re.match(r'^\d{3}-\d{3}$', t): lugar = t 
                            elif re.match(r'^\d{2}:\d{2}$', t): hora = t 
                            elif re.match(r'^\d{6}$', t): num_op = t 
                            elif re.match(r'^\d{4}$', t): sucursal = t 
                            elif len(t) == 6 and t.isalnum(): hora += f" {t}"
                        es_transaccion = True

                if es_transaccion:
                    data_final.append([pagina, fecha, desc, medio, lugar, sucursal, num_op, hora.strip(), cargo, abono, itf, saldo])
                    
            # 4. FORMATEO DE DATOS Y EXCEL
            columnas = ["PAGINA", "FECHA", "DESCRIPCION", "MEDIO", "LUGAR", "SUCURSAL", "NUMERO DE OPERACION", "HORA U ORIGEN", "CARGO", "ABONO", "ITF", "SALDO"]
            
            def formatear_fecha(fecha_str):
                fecha_str = str(fecha_str).strip().upper()
                if not fecha_str: return ""
                if re.search(r'\d{4}', fecha_str): return fecha_str
                
                match_banbif = re.match(r'^(\d{2})/(\d{2})/(\d{2})$', fecha_str)
                if match_banbif: return f"{match_banbif.group(1)}/{match_banbif.group(2)}/20{match_banbif.group(3)}"
                
                meses = {'ENE':'01', 'FEB':'02', 'MAR':'03', 'ABR':'04', 'MAY':'05', 'JUN':'06', 
                         'JUL':'07', 'AGO':'08', 'SET':'09', 'SEP':'09', 'OCT':'10', 'NOV':'11', 'DIC':'12'}
                match1 = re.match(r'^(\d{2})[/-](\d{2})$', fecha_str)
                if match1: return f"{match1.group(1)}/{match1.group(2)}/{doc_year}"
                match2 = re.match(r'^(\d{2})-?([A-Z]{3})$', fecha_str)
                if match2: return f"{match2.group(1)}/{meses.get(match2.group(2), '01')}/{doc_year}"
                return fecha_str

            def limpiar_numero_espacios(val):
                if not val or str(val).strip() == "": return None
                val_str = str(val).replace(',', '').strip()
                try: 
                    num = float(val_str)
                    return num if num != 0 else None
                except: return None
            
            def parse_float_seguro(val):
                if not val or str(val).strip() == "": return 0.0
                try: return float(str(val).replace(',', '').strip())
                except: return 0.0

            df_final = pd.DataFrame(data_final, columns=columnas)
            
            if len(df_final) > 0:
                df_final['FECHA'] = df_final['FECHA'].apply(formatear_fecha)
                
                if saldo_inicial_declarado is None:
                    s_primero = parse_float_seguro(data_final[0][11])
                    c_primero = parse_float_seguro(data_final[0][8])
                    a_primero = parse_float_seguro(data_final[0][9])
                    i_primero = parse_float_seguro(data_final[0][10])
                    saldo_inicial = s_primero + c_primero + i_primero - a_primero
                else:
                    saldo_inicial = saldo_inicial_declarado

                df_final['CARGO'] = df_final['CARGO'].apply(limpiar_numero_espacios)
                df_final['ABONO'] = df_final['ABONO'].apply(limpiar_numero_espacios)
                df_final['ITF'] = df_final['ITF'].apply(limpiar_numero_espacios)
                df_final['SALDO'] = None 
                
                fila_saldo_anterior = pd.DataFrame([["", "", "SALDO ANTERIOR", "", "", "", "", "", None, None, None, None]], columns=columnas)
                df_final = pd.concat([fila_saldo_anterior, df_final], ignore_index=True)

                st.success(f"✅ ¡Éxito! Se extrajeron {len(df_final) - 1} transacciones con fórmulas integradas.")
                st.dataframe(df_final.head())
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False, startrow=8, sheet_name="Estado de Cuenta")
                    worksheet = writer.sheets["Estado de Cuenta"]
                    
                    # --- DISEÑO DEL EXCEL CON COLORES DE RESPALDO TRIBUTARIO ---
                    color_marca = PatternFill(start_color="2E1E7E", end_color="2E1E7E", fill_type="solid")
                    fuente_blanca = Font(color="FFFFFF", bold=True)
                    fuente_logo = Font(name="Bahnschrift", color="FFFFFF", bold=True, size=40, italic=True)
                    borde_blanco = Border(
                        left=Side(style='medium', color="FFFFFF"), right=Side(style='medium', color="FFFFFF"),
                        top=Side(style='medium', color="FFFFFF"), bottom=Side(style='medium', color="FFFFFF")
                    )
                    centro = Alignment(horizontal="center", vertical="center")
                    
                    ancho_tabla = len(df_final.columns)
                    
                    for row in range(1, 8):
                        for col in range(1, ancho_tabla + 1):
                            worksheet.cell(row=row, column=col).fill = color_marca
                            
                    celda_banco = worksheet.cell(row=4, column=2, value=f"  {banco.replace('_', ' ')}  ")
                    celda_banco.font = fuente_logo
                    
                    col_inicio_resumen = ancho_tabla - 3
                    if col_inicio_resumen < 6: 
                        col_inicio_resumen = 6
                    
                    titulos_resumen = ["SALDO DISPONIBLE", "TOTAL CARGOS", "TOTAL ABONOS", "SALDO FINAL"]
                    
                    for i in range(4):
                        c_tit = worksheet.cell(row=3, column=col_inicio_resumen + i, value=titulos_resumen[i])
                        c_tit.font, c_tit.fill, c_tit.border, c_tit.alignment = fuente_blanca, color_marca, borde_blanco, centro
                        
                        c_val = worksheet.cell(row=4, column=col_inicio_resumen + i)
                        c_val.font, c_val.fill, c_val.border, c_val.alignment = fuente_blanca, color_marca, borde_blanco, centro
                        c_val.number_format = '#,##0.00'
                    
                    last_row = 9 + len(df_final)
                    
                    worksheet.cell(row=4, column=9).value = saldo_inicial 
                    worksheet.cell(row=4, column=10).value = f"=SUM(I11:I{last_row}) + SUM(K11:K{last_row})" 
                    worksheet.cell(row=4, column=11).value = f"=SUM(J11:J{last_row})" 
                    worksheet.cell(row=4, column=12).value = "=I4+K4-J4" 
                    
                    for r in range(10, last_row + 1):
                        if r == 10:
                            worksheet.cell(row=r, column=12).value = "=I4" 
                        else:
                            worksheet.cell(row=r, column=12).value = f"=L{r-1}-I{r}+J{r}-K{r}"
                    
                    for r in range(10, last_row + 1):
                        for c in [9, 10, 11, 12]:
                            worksheet.cell(row=r, column=c).number_format = '#,##0.00'
                        
                    for col in worksheet.columns:
                        max_length = 0
                        column = col[0].column_letter
                        for cell in col:
                            try:
                                if len(str(cell.value)) > max_length: 
                                    max_length = len(str(cell.value))
                            except: pass
                        worksheet.column_dimensions[column].width = min(max_length + 2, 50)
                
                st.download_button(
                    label="📥 Descargar Excel Corporativo (Formato y Fórmulas)",
                    data=buffer.getvalue(),
                    file_name=f"Estado_Cuenta_{banco}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("⚠️ El documento es válido pero no se detectaron transacciones legibles.")
                
        except Exception as e:
            st.error(f"❌ Ocurrió un error procesando el archivo. Detalle: {e}")
