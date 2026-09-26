import streamlit as st
import fitz  # PyMuPDF
import re
import io

st.set_page_config(page_title="Invoice Operations Suite", page_icon="📑", layout="centered")

# Custom Professional UI
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .hero-container {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        padding: 15px 0 10px 0;
    }
    .logo-badge {
        width: 60px;
        height: 60px;
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        border-radius: 16px;
        display: flex;
        align-items: center;
        justify-content: center;
        box-shadow: 0 8px 20px rgba(30, 60, 114, 0.3);
        margin-bottom: 10px;
    }
    .logo-icon { font-size: 28px; }
    .brand-title { font-size: 22px; font-weight: 700; color: #FFFFFF; }
    .brand-sub { font-size: 13px; color: #94a3b8; margin-bottom: 15px; }
    </style>
    <div class="hero-container">
        <div class="logo-badge"><span class="logo-icon">⚡</span></div>
        <div class="brand-title">Blinkit Invoice Gateway</div>
        <div class="brand-sub">Auto Unit Conversion & Layout Preservation Engine</div>
    </div>
""", unsafe_allow_html=True)

uploaded_invoice = st.file_uploader("", type=["pdf"])

def overwrite_area(page, rect, new_text, font_size=7):
    """Purane text ko white box se mask karke naya text print karta hai"""
    page.draw_rect(rect, color=None, fill=(1, 1, 1))
    page.insert_text(
        (rect.x0, rect.y1 - 1.2),
        str(new_text),
        fontsize=font_size,
        fontname="helv",
        color=(0, 0, 0)
    )

def process_dynamic_invoice(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    for page in doc:
        # Page ke saare words unke coordinates ke sath read karein
        # Word format: (x0, y0, x1, y1, word, block_no, line_no, word_no)
        words = page.get_text("words")
        
        # Line wise words group karein
        lines = {}
        for w in words:
            key = (w[5], w[6]) # block_no, line_no
            if key not in lines:
                lines[key] = []
            lines[key].append(w)

        for _, line_words in lines.items():
            line_str = " ".join([w[4] for w in line_words]).upper()
            
            # Determine conversion factor based on product
            factor = None
            if "DISPO" in line_str:
                factor = 50
            elif "COMFIT" in line_str:
                factor = 25

            if factor:
                for w in line_words:
                    clean_word = w[4].replace(",", "").strip()
                    rect = fitz.Rect(w[0], w[1], w[2], w[3])
                    
                    # 1. Quantity Detection (Whole number jaise 20000, 5000, 1000)
                    if clean_word.isdigit() and int(clean_word) >= factor:
                        original_qty = int(clean_word)
                        # Agar yeh line item quantity hai
                        new_qty = original_qty // factor
                        overwrite_area(page, rect, f"{new_qty:,}")
                        
                    # 2. Rate Detection (Decimal price jaise 1.64, 2.50)
                    elif re.match(r"^\d+\.\d{2}$", clean_word):
                        original_rate = float(clean_word)
                        # Filter to avoid GST rates like 5.00, 12.00, 18.00 if separate
                        new_rate = original_rate * factor
                        overwrite_area(page, rect, f"{new_rate:.2f}")

        # Standard UOM Replacements
        for uom_text, target_text in [("UOM-PCS", "UOM-BOX"), (" PCS", " BOX"), ("PCS", "BOX")]:
            instances = page.search_for(uom_text)
            for inst in instances:
                overwrite_area(page, inst, target_text)

    out_buffer = io.BytesIO()
    doc.save(out_buffer)
    doc.close()
    out_buffer.seek(0)
    return out_buffer

if uploaded_invoice:
    st.info("Invoice analyze ho raha hai aur Dispo/Comfit formula apply ho raha hai...")
    
    try:
        updated_pdf_buffer = process_dynamic_invoice(uploaded_invoice.read())
        
        st.success("Invoice successfully reconcile ho gaya!")
        st.download_button(
            label="📥 Download Updated Invoice",
            data=updated_pdf_buffer,
            file_name="Blinkit_Reconciled_Invoice.pdf",
            mime="application/pdf"
        )
    except Exception as e:
        st.error(f"Processing error: {str(e)}")
