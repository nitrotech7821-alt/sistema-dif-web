import streamlit as st
import imaplib
import email
import openpyxl
import io
import zipfile
import xml.etree.ElementTree as ET

st.title("🤖 Robot Administrativo DIF")

def escribir_en_celda(ws, celda, valor):
    """Escribe en una celda, manejando celdas combinadas."""
    for range_ in ws.merged_cells.ranges:
        if celda in range_:
            ws.unmerge_cells(str(range_))
            break
    ws[celda] = valor

def extraer_folio_xml(xml_data):
    """Extrae el folio del XML de forma robusta."""
    try:
        root = ET.fromstring(xml_data)
        # El atributo 'Folio' suele estar en la raíz del comprobante
        folio = root.get('Folio')
        return folio if folio else "SIN_FOLIO"
    except:
        return "ERROR_LECTURA"

if st.button("🚀 BUSCAR Y PROCESAR FACTURAS"):
    try:
        # 1. Conexión al correo
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(st.secrets["EMAIL_USER"], st.secrets["EMAIL_PASS"])
        mail.select("inbox")
        
        # 2. Búsqueda de facturas
        _, mensajes = mail.search(None, '(SUBJECT "Comprobante fiscal")')
        ids = mensajes[0].split()[-5:] # Últimos 5 correos
        
        for num in ids:
            _, data = mail.fetch(num, "(RFC822)")
            msg = email.message_from_bytes(data[0][1])
            
            for part in msg.walk():
                if part.get_filename() and part.get_filename().lower().endswith(".xml"):
                    xml_data = part.get_payload(decode=True)
                    folio = extraer_folio_xml(xml_data)
                    
                    # 3. Procesar Excels
                    wb_cot = openpyxl.load_workbook("cotizacion.xlsx")
                    escribir_en_celda(wb_cot.active, 'H2', folio)
                    
                    wb_req = openpyxl.load_workbook("requisicion.xlsx")
                    escribir_en_celda(wb_req.active, 'I7', folio)
                    
                    # 4. Crear ZIP en memoria
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, 'w') as zf:
                        zf.writestr(f"Factura_{folio}.xml", xml_data)
                        
                        buf_cot = io.BytesIO()
                        wb_cot.save(buf_cot)
                        zf.writestr(f"Cotizacion_{folio}.xlsx", buf_cot.getvalue())
                        
                        buf_req = io.BytesIO()
                        wb_req.save(buf_req)
                        zf.writestr(f"Requisicion_{folio}.xlsx", buf_req.getvalue())
                    
                    # 5. Botón de descarga
                    st.download_button(
                        label=f"📥 Descargar Expediente: {folio}",
                        data=zip_buf.getvalue(),
                        file_name=f"Expediente_{folio}.zip",
                        mime="application/zip"
                    )
        mail.logout()
        st.success("Búsqueda finalizada.")
    except Exception as e:
        st.error(f"Error técnico: {e}")
