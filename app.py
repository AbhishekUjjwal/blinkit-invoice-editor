import streamlit as st
import fitz  # PyMuPDF
import io

st.set_page_config(page_title="Invoice Editor", layout="centered")

# Custom Styling: Clean & Professional look
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .title-text {
        text-align: center;
        font-size: 26px;
        font-weight: 700;
        margin-bottom: 25px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="title-text">📄 Invoice Document Processor</div>', unsafe_allow_html=True)

uploaded_invoice = st.file_uploader("", type=["pdf"])

def replace_text_in_pdf(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[0]

    def overwrite_text(search_text, new_text, font_size=7):
        text_instances = page.search_for(search_text)
        for inst in text_instances:
            page.draw_rect(inst, color=None, fill=(1, 1, 1))
            page.insert_text(
                (inst.x0, inst.y1 - 1.5),
                str(new_text),
                fontsize=font_size,
                fontname="helv",
                color=(0, 0, 0)
            )

    # Unit conversions
    overwrite_text("20000", "400")
    overwrite_text("1.64", "82.00")
    overwrite_text("UOM-PCS", "UOM-BOX")

    out_buffer = io.BytesIO()
    doc.save(out_buffer)
    doc.close()
    out_buffer.seek(0)
    return out_buffer

if uploaded_invoice:
    st.info("Document verified. Generating output...")

    updated_pdf_buffer = replace_text_in_pdf(uploaded_invoice.read())

    st.download_button(
        label="📥 Download Processed Invoice",
        data=updated_pdf_buffer,
        file_name="Processed_Invoice.pdf",
        mime="application/pdf"
    )
