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
                        try: saldo_inicial_declarado = float(items[0][1].replace(',', ''))
                        except: pass
                        break

            saldo_previo = saldo_inicial_declarado if saldo_inicial_declarado is not None else 0.0
            data_final = []
            
            # 3. PROCESAMIENTO DE TRANSACCIONES POR BANCO
            for key in sorted_keys:
                items = sorted(lines_by_page_and_y[key], key=lambda i: i[0]) 
                combined = " ".join([i[1] for i in items])
                pagina = key[0]
                
                fecha, desc, medio, lugar, sucursal, num_op, hora, cargo, abono, itf, saldo = [""]*11
                es_transaccion = False

                # LOGICA: SCOTIABANK 
                if banco == "SCOTIABANK":
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

                # LOGICA: INTERBANK
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

                # LOGICA: BCP AHORROS
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

                # LOGICA: BBVA
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

                # LOGICA: BCP CORRIENTE
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
                    
            # 4. CREAR EXCEL EN MEMORIA CON DISEÑO Y CÁLCULOS
            columnas = ["PAGINA", "FECHA", "DESCRIPCION", "MEDIO", "LUGAR", "SUCURSAL", "NUMERO DE OPERACION", "HORA U ORIGEN", "CARGO", "ABONO", "ITF", "SALDO"]
            df_final = pd.DataFrame(data_final, columns=columnas)
            
            if len(df_final) > 0:
                def limpiar_numero(val):
                    if not val or str(val).strip() == "": return 0.0
                    val_str = str(val).replace(',', '').strip()
                    try: return float(val_str)
                    except: return 0.0
                
                df_final['CARGO'] = df_final['CARGO'].apply(limpiar_numero)
                df_final['ABONO'] = df_final['ABONO'].apply(limpiar_numero)
                df_final['SALDO'] = df_final['SALDO'].apply(limpiar_numero)
                
                total_cargos = df_final['CARGO'].sum()
                total_abonos = df_final['ABONO'].sum()
                saldo_final = df_final['SALDO'].iloc[-1]
                
                saldo_inicial = saldo_inicial_declarado if saldo_inicial_declarado is not None else (df_final['SALDO'].iloc[0] + df_final['CARGO'].iloc[0] - df_final['ABONO'].iloc[0])

                st.success(f"✅ ¡Éxito! Se extrajeron {len(df_final)} transacciones.")
                st.dataframe(df_final.head())
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False, startrow=8, sheet_name="Estado de Cuenta")
                    worksheet = writer.sheets["Estado de Cuenta"]
                    
                    color_azul_oscuro = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
                    fuente_blanca = Font(color="FFFFFF", bold=True)
                    fuente_logo = Font(color="FFFFFF", bold=True, size=40, italic=True)
                    borde_blanco = Border(
                        left=Side(style='medium', color="FFFFFF"), right=Side(style='medium', color="FFFFFF"),
                        top=Side(style='medium', color="FFFFFF"), bottom=Side(style='medium', color="FFFFFF")
                    )
                    centro = Alignment(horizontal="center", vertical="center")
                    
                    ancho_tabla = len(df_final.columns)
                    for row in range(1, 8):
                        for col in range(1, ancho_tabla + 1):
                            worksheet.cell(row=row, column=col).fill = color_azul_oscuro
                            
                    celda_banco = worksheet.cell(row=4, column=2, value=f"  {banco.replace('_', ' ')}  ")
                    celda_banco.font = fuente_logo
                    
                    titulos_resumen = ["SALDO DISPONIBLE", "TOTAL CARGOS", "TOTAL ABONOS", "SALDO FINAL"]
                    valores_resumen = [saldo_inicial, total_cargos, total_abonos, saldo_final]
                    
                    col_inicio_resumen = ancho_tabla - 3
                    if col_inicio_resumen < 6: col_inicio_resumen = 6
                    
                    for i in range(4):
                        c_tit = worksheet.cell(row=3, column=col_inicio_resumen + i, value=titulos_resumen[i])
                        c_tit.font, c_tit.fill, c_tit.border, c_tit.alignment = fuente_blanca, color_azul_oscuro, borde_blanco, centro
                        
                        c_val = worksheet.cell(row=4, column=col_inicio_resumen + i, value=valores_resumen[i])
                        c_val.font, c_val.fill, c_val.border, c_val.alignment = fuente_blanca, color_azul_oscuro, borde_blanco, centro
                        c_val.number_format = '#,##0.00'
                        
                    for col in worksheet.columns:
                        max_length = 0
                        column = col[0].column_letter
                        for cell in col:
                            try:
                                if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
                            except: pass
                        worksheet.column_dimensions[column].width = min(max_length + 2, 50)
                
                st.download_button(
                    label="📥 Descargar Excel con Resumen Automático",
                    data=buffer.getvalue(),
                    file_name=f"Estado_Cuenta_{banco}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("⚠️ El documento es válido pero no se detectaron transacciones legibles.")
                
        except Exception as e:
            st.error(f"❌ Ocurrió un error procesando el archivo. Detalle: {e}")
