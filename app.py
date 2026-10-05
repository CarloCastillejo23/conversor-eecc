import streamlit as st
import pypdf
import pandas as pd
import io
import re

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
            
            # 1. DETECCIÓN AUTOMÁTICA DEL BANCO
            texto_upper = texto_completo.upper()
            banco = "DESCONOCIDO"
            if "SCOTIA" in texto_upper or "20100043140" in texto_upper:
                banco = "SCOTIABANK"
            elif "INTERBANK" in texto_upper and "NEGOCIOS" in texto_upper:
                banco = "INTERBANK"
            elif "BBVA" in texto_upper:
                banco = "BBVA"
            elif "ESTADO DE CUENTA DE AHORROS" in texto_upper:
                banco = "BCP_AHORROS"
            elif "ESTADO DE CUENTA CORRIENTE" in texto_upper:
                banco = "BCP_CORRIENTE"
                
            st.info(f"🏦 Banco detectado automáticamente: **{banco.replace('_', ' ')}**")
            
            data_final = []
            saldo_previo = 0.0 
            
            # 2. PROCESAMIENTO POR BANCO
            for key in sorted_keys:
                items = sorted(lines_by_page_and_y[key], key=lambda i: i[0]) 
                combined = " ".join([i[1] for i in items])
                pagina = key[0]
                
                # Se añade ITF a las variables
                fecha, desc, medio, lugar, sucursal, num_op, hora, cargo, abono, itf, saldo = [""]*11
                es_transaccion = False

                # LOGICA: SCOTIABANK
                if banco == "SCOTIABANK":
                    if "Saldo Final al" in combined and re.search(r'\d{4}', combined):
                        try: saldo_previo = float(combined.split()[-1].replace(',', ''))
                        except: pass
                    
                    elif re.match(r'^\d{2}/\d{2}\s', combined):
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
                            
                        desc = " ".join(tokens[2:-2])
                        es_transaccion = True

                # LOGICA: INTERBANK (Actualizada con Num. Operación)
                elif banco == "INTERBANK":
                    if re.match(r'^\d{2}/\d{2}\s+\d{2}/\d{2}', combined):
                        tokens = combined.split()
                        fecha = tokens[0]
                        saldo = tokens[-1]
                        monto_raw = tokens[-2]
                        cargo = monto_raw.replace('-', '') if '-' in monto_raw else ""
                        abono = monto_raw if '-' not in monto_raw else ""
                        
                        # Extraer código de operación (7 dígitos)
                        num_op_match = re.search(r'\b\d{7}\b', combined)
                        if num_op_match: num_op = num_op_match.group(0)
                        
                        middle_tokens = tokens[2:-2]
                        if "WEB" in middle_tokens: medio = "WEB"
                        elif "INTERNO" in middle_tokens: medio = "INTERNO"
                        
                        # Unir descripción limpiando el medio y num_op
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

                # LOGICA: BBVA (Actualizada con ITF y Medio/Lugar separados)
                elif banco == "BBVA":
                    if len(items) > 3 and re.match(r'^\d{2}-\d{2}$', items[0][1]):
                        for x, t in items:
                            if x < 60: fecha = t
                            elif 100 < x < 250: desc += t + " "
                            elif 250 < x < 310: lugar = t # OFICINA BBVA
                            elif 310 < x < 330: medio = t # CANAL BBVA
                            elif 330 < x < 380: num_op = t # NUM. OPER. BBVA
                            elif 380 < x < 430:
                                if '-' in t: cargo = t.replace('-', '')
                                else: abono = t
                            elif 430 < x < 490: itf = t # COLUMNA ITF BBVA
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
                    
            # 3. CREAR EXCEL EN MEMORIA
            # Columna ITF añadida aquí
            columnas = ["PAGINA", "FECHA", "DESCRIPCION", "MEDIO", "LUGAR", "SUCURSAL", "NUMERO DE OPERACION", "HORA U ORIGEN", "CARGO", "ABONO", "ITF", "SALDO"]
            df_final = pd.DataFrame(data_final, columns=columnas)
            
            if len(df_final) > 0:
                st.success(f"✅ ¡Éxito! Se extrajeron {len(df_final)} transacciones.")
                st.dataframe(df_final.head())
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False)
                
                st.download_button(
                    label="📥 Descargar Excel Uniformizado",
                    data=buffer.getvalue(),
                    file_name=f"Estado_Cuenta_{banco}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("⚠️ El documento es válido pero no se detectaron transacciones legibles.")
                
        except Exception as e:
            st.error(f"❌ Ocurrió un error procesando el archivo. Detalle: {e}")
