import streamlit as st
import imaplib
import email
import openpyxl
import io
import zipfile
import xml.etree.ElementTree as ET

st.title("🤖 Robot Administrativo DIF")

def escribir_en_celda(ws, celda, valor):
    for range_ in ws.merged_cells.ranges:
        if celda in range_:
            ws.unmerge_cells(str(range_))
            break
    ws[celda] = valor

if st.button("🚀 PROCESAR FACTURAS DE FERRETERIA CALOTE"):
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(st.secrets["EMAIL_USER"], st.secrets["EMAIL_PASS"])
        mail.select("inbox")
        
        # Filtro corregido: busca correos con el asunto específico
        _, mensajes = mail.search(None, '(SUBJECT "Comprobante fiscal")')
        ids = mensajes[0].split()[-5:] # Últimos 5 correos
        
        for num in ids:
            _, data = mail.fetch(num, "(RFC822)")
            msg = email.message_from_bytes(data[0][1])
            
            for part in msg.walk():
                # Buscamos archivos que terminen en .xml
                if part.get_filename() and part.get_filename().lower().endswith(".xml"):
                    xml_data = part.get_payload(decode=True)
                    
                    # Extraer folio del XML
                    try:
                        root = ET.fromstring(xml_data)
                        ns = {'cfdi': 'http://www.sat.gob.mx/cfd/4'}
                        folio = root.find('.//cfdi:Comprobante', ns).get('Folio', 'SN')
                    except:
                        folio = "ERROR_XML"
                    
                    # Procesar Excels
                    wb_cot = openpyxl.load_workbook("cotizacion.xlsx")
                    escribir_en_celda(wb_cot.active, 'H2', folio)
                    
                    wb_req = openpyxl.load_workbook("requisicion.xlsx")
                    escribir_en_celda(wb_req.active, 'I7', folio)
                    
                    # Crear ZIP
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, 'w') as zf:
                        zf.writestr(f"Factura_{folio}.xml", xml_data)
                        
                        buf_cot = io.BytesIO()
                        wb_cot.save(buf_cot)
                        zf.writestr(f"Cotizacion_{folio}.xlsx", buf_cot.getvalue())
                        
                        buf_req = io.BytesIO()
                        wb_req.save(buf_req)
                        zf.writestr(f"Requisicion_{folio}.xlsx", buf_req.getvalue())
                    
                    st.download_button(
                        label=f"📥 Descargar Expediente: {folio}",
                        data=zip_buf.getvalue(),
                        file_name=f"Expediente_{folio}.zip",
                        mime="application/zip"
                    )
        mail.logout()
        st.success("Búsqueda finalizada.")
    except Exception as e:
        st.error(f"Error: {e}")
