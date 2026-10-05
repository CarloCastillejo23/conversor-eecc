import streamlit as st
import pypdf
import pandas as pd
import io
import re

st.set_page_config(page_title="Conversor BCP", page_icon="📄")
st.title("📄 Convertidor de Estado de Cuenta BCP")
st.write("Sube el PDF de tu estado de cuenta y descárgalo en Excel al instante.")

archivo_pdf = st.file_uploader("Sube tu Estado de Cuenta (PDF)", type=["pdf"])

if archivo_pdf is not None:
    with st.spinner("Procesando documento..."):
        try:
            reader = pypdf.PdfReader(archivo_pdf)
            all_text_with_coords = []
            for page_num, page in enumerate(reader.pages):
                def visitor_extract(text, cm, tm, fontDict, fontSize):
                    x, y = tm[4], tm[5]
                    if text.strip():
                        all_text_with_coords.append((page_num + 1, x, y, text.strip()))
                page.extract_text(visitor_text=visitor_extract)
                
            lines_by_page_and_y = {}
            for p, x, y, text in all_text_with_coords:
                key = (p, round(y, 1))
                lines_by_page_and_y.setdefault(key, []).append((x, text))
                
            sorted_keys = sorted(lines_by_page_and_y.keys(), key=lambda k: (k[0], -k[1]))
            data_final = []
            
            for key in sorted_keys:
                items = sorted(lines_by_page_and_y[key], key=lambda i: i[0]) 
                combined = " ".join([i[1] for i in items])
                
                if re.match(r'^\d{2}-\d{2}\s', combined):
                    tokens = combined.split()
                    pagina = key[0]
                    fecha = tokens[0]
                    saldo = tokens[-1]
                    monto_raw = tokens[-2]
                    
                    cargo = monto_raw.replace('-', '') if '-' in monto_raw else ""
                    abono = monto_raw if '-' not in monto_raw else ""
                    
                    middle = tokens[1:-2] 
                    descripcion, medio, lugar, sucursal, num_op, hora_origen = "", "", "", "", "", ""
                    
                    medios_conocidos = ["BPI", "POS", "VEN", "INT", "CAJ", "TLC", "BPT"]
                    medio_index = next((i for i, t in enumerate(middle) if t in medios_conocidos), -1)
                        
                    if medio_index != -1:
                        descripcion = " ".join(middle[:medio_index])
                        rest = middle[medio_index+1:]
                    else:
                        descripcion = " ".join(middle[:3]) 
                        rest = middle[3:]
                        
                    for t in rest:
                        if re.match(r'^\d{3}-\d{3}$', t): lugar = t 
                        elif re.match(r'^\d{2}:\d{2}$', t): hora_origen = t 
                        elif re.match(r'^\d{6}$', t): num_op = t 
                        elif re.match(r'^\d{4}$', t): sucursal = t 
                        elif len(t) == 6 and t.isalnum(): hora_origen += f" {t}"
                                 
                    data_final.append([pagina, fecha, descripcion, medio, lugar, sucursal, num_op, hora_origen.strip(), cargo, abono, saldo])
                    
            df_final = pd.DataFrame(data_final, columns=["PAGINA", "FECHA", "DESCRIPCION", "MEDIO", "LUGAR", "SUCURSAL", "NUMERO DE OPERACION", "HORA U ORIGEN", "CARGO", "ABONO", "SALDO"])
            
            if len(df_final) > 0:
                st.success(f"✅ ¡Éxito! Se extrajeron {len(df_final)} transacciones.")
                st.dataframe(df_final.head()) # Muestra una pequeña vista previa
                
                # Crear Excel en memoria
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False)
                
                st.download_button(
                    label="📥 Descargar Excel",
                    data=buffer.getvalue(),
                    file_name="Estado_Cuenta_Convertido.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No se encontraron transacciones legibles.")
        except Exception as e:
            st.error(f"Error procesando el archivo: {e}")
