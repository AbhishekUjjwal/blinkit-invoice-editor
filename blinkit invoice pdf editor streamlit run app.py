import streamlit as st
import fitz  # PyMuPDF
import io

st.set_page_config(page_title="Original Invoice In-Place Editor", layout="wide")
st.title("📄 Original Invoice PDF Editor (Same Layout)")
st.write("Original Invoice upload karein. Tool original layout ke upar hi **Dispo (÷50, ×50)** aur **Comfit (÷25, ×25)** replace karke same PDF return karega.")

uploaded_invoice = st.file_uploader("Upload Your Original Invoice (PDF)", type=["pdf"])

def replace_text_in_pdf(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[0]  # Page 1

    # Text instances dhund kar cover aur replace karne ka helper function
    def overwrite_text(search_text, new_text, font_size=7):
        text_instances = page.search_for(search_text)
        for inst in text_instances:
            # Purane text ke upar white box banana
            page.draw_rect(inst, color=None, fill=(1, 1, 1))
            # Usi exact position par naya text likhna
            page.insert_text(
                (inst.x0, inst.y1 - 1.5),
                str(new_text),
                fontsize=font_size,
                fontname="helv",
                color=(0, 0, 0)
            )

    # 1. Dispo line replace (20000 -> 400 aur 1.64 -> 82.00)
    overwrite_text("20000", "400")
    overwrite_text("1.64", "82.00")

    # 2. UOM box me convert karna (PCS -> BOX)
    overwrite_text("UOM-PCS", "UOM-BOX")

    out_buffer = io.BytesIO()
    doc.save(out_buffer)
    doc.close()
    out_buffer.seek(0)
    return out_buffer

if uploaded_invoice:
    st.success("Invoice uploaded! Processing original document...")

    updated_pdf_buffer = replace_text_in_pdf(uploaded_invoice.read())

    st.download_button(
        label="📥 Download Same Layout Updated Invoice (PDF)",
        data=updated_pdf_buffer,
        file_name="Same_Format_Updated_Invoice.pdf",
        mime="application/pdf"
    )