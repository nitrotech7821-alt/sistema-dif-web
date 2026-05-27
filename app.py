import streamlit as st
import imaplib
import email
import openpyxl
import io
import zipfile
import xml.etree.ElementTree as ET

st.title("🤖 Robot Administrativo DIF - Procesador Total")

def escribir_en_celda(ws, celda, valor):
    for range_ in ws.merged_cells.ranges:
        if celda in range_:
            ws.unmerge_cells(str(range_))
            break
    ws[celda] = valor

if st.button("🚀 BUSCAR Y PROCESAR TODA LA BANDEJA"):
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(st.secrets["EMAIL_USER"], st.secrets["EMAIL_PASS"])
        mail.select("inbox")
        
        # Buscamos TODOS los correos
        _, mensajes = mail.search(None, 'ALL')
        ids = mensajes[0].split()[-10:] # Tomamos los últimos 10 para no colapsar la app
        
        for num in ids:
            _, data = mail.fetch(num, "(RFC822)")
            msg = email.message_from_bytes(data[0][1])
            
            for part in msg.walk():
                if part.get_filename() and part.get_filename().endswith(".xml"):
                    # Extraer datos
                    xml_data = part.get_payload(decode=True)
                    try:
                        root = ET.fromstring(xml_data)
                        ns = {'cfdi': 'http://www.sat.gob.mx/cfd/4'}
                        folio = root.find('.//cfdi:Comprobante', ns).get('Folio', 'SIN_FOLIO')
                    except:
                        folio = "ERROR_XML"
                    
                    # Procesar Excel
                    wb_cot = openpyxl.load_workbook("cotizacion.xlsx")
                    escribir_en_celda(wb_cot.active, 'H2', folio)
                    
                    wb_req = openpyxl.load_workbook("requisicion.xlsx")
                    escribir_en_celda(wb_req.active, 'I7', folio)
                    
                    buf_cot = io.BytesIO()
                    wb_cot.save(buf_cot)
                    buf_req = io.BytesIO()
                    wb_req.save(buf_req)
                    
                    # Empaquetar
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, 'w') as zf:
                        zf.writestr(f"Factura_{folio}.xml", xml_data)
                        zf.writestr(f"Cotizacion_{folio}.xlsx", buf_cot.getvalue())
                        zf.writestr(f"Requisicion_{folio}.xlsx", buf_req.getvalue())
                    
                    # Mostrar botón
                    st.download_button(
                        label=f"📥 Descargar: Factura {folio}",
                        data=zip_buf.getvalue(),
                        file_name=f"Expediente_{folio}.zip",
                        mime="application/zip"
                    )
        mail.logout()
        st.success("Búsqueda finalizada.")
    except Exception as e:
        st.error(f"Error crítico: {e}")
