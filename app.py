import streamlit as st
import fitz  # PyMuPDF
import re
import io
import zipfile
from datetime import datetime

st.set_page_config(page_title="Invoice Operations Suite", page_icon="📑", layout="centered")

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
        <div class="brand-title">Blinkit Bulk Invoice Gateway</div>
        <div class="brand-sub">Universal Dynamic Reconciliation Engine (Batch Mode)</div>
    </div>
""", unsafe_allow_html=True)

# Multiple files upload support
uploaded_invoices = st.file_uploader("Upload Blinkit Invoices (Single ya Multiple)", type=["pdf"], accept_multiple_files=True)

def overwrite_area(page, rect, new_text, font_size=7):
    """Purane text ko white mask karke naya value likhta hai"""
    pad_rect = fitz.Rect(rect.x0 - 2, rect.y0 - 1, rect.x1 + 2, rect.y1 + 1)
    page.draw_rect(pad_rect, color=None, fill=(1, 1, 1))
    page.insert_text(
        (rect.x0, rect.y1 - 1.2),
        str(new_text),
        fontsize=font_size,
        fontname="helv",
        color=(0, 0, 0)
    )

def extract_metadata(doc):
    full_text = ""
    for page in doc:
        full_text += page.get_text() + "\n"

    ext_order_id = None
    ext_match = re.search(r"Extern(?:al)?\s*Order\s*(?:No\.?|ID)?\s*[:\s]*([0-9\s]{8,25})", full_text, re.IGNORECASE)
    if ext_match:
        ext_order_id = "".join(ext_match.group(1).split())

    if not ext_order_id:
        digits_match = re.findall(r"\b(49\d{8,14}|5\d{8,14}|\d{12,18})\b", full_text)
        if digits_match:
            ext_order_id = digits_match[0]

    if not ext_order_id:
        ext_order_id = "ExtOrder"

    inv_match = re.search(r"Invoice\s*No\s*[:\s]*([A-Za-z0-9\-_]+)", full_text, re.IGNORECASE)
    invoice_no = inv_match.group(1).strip() if inv_match else "Invoice"

    date_match = re.search(r"Invoice\s*Date\s*[:\s]*([A-Za-z0-9,\s\.\-\/]+?)(?=\n|Ship\s*Date|$)", full_text, re.IGNORECASE)
    raw_date = date_match.group(1).strip() if date_match else "Date"

    clean_date = raw_date
    for fmt in ("%b %d, %Y", "%B %d, %Y", "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            d_obj = datetime.strptime(raw_date, fmt)
            clean_date = d_obj.strftime("%d-%m-%Y")
            break
        except ValueError:
            pass

    clean_ext = re.sub(r'[^A-Za-z0-9\-_]', '', ext_order_id)
    clean_inv = re.sub(r'[^A-Za-z0-9\-_]', '', invoice_no)
    clean_dt = re.sub(r'[^A-Za-z0-9\-_]', '', clean_date)

    return f"{clean_ext}_{clean_inv}_{clean_dt}.pdf"

def process_universal_blinkit_invoice(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    download_filename = extract_metadata(doc)

    for page in doc:
        words = page.get_text("words")
        
        qty_box = None
        price_box = None

        for w in words:
            w_txt = w[4].strip().lower()
            if w_txt == "qty":
                qty_box = (w[0] - 15, w[2] + 25)
            elif w_txt == "price" or "unit" in w_txt:
                if not price_box:
                    price_box = (w[0] - 15, w[2] + 40)

        if not qty_box:
            qty_box = (320, 390)
        if not price_box:
            price_box = (390, 480)

        product_rows = []
        for w in words:
            text = w[4].upper()
            if "DISPO" in text:
                product_rows.append({"factor": 50, "y": (w[1] + w[3]) / 2})
            elif "COMFIT" in text:
                product_rows.append({"factor": 25, "y": (w[1] + w[3]) / 2})

        for prod in product_rows:
            factor = prod["factor"]
            row_y = prod["y"]
            row_words = [w for w in words if abs(((w[1] + w[3]) / 2) - row_y) <= 25]

            # Qty update
            for w in row_words:
                val = w[4].replace(",", "").strip()
                rect = fitz.Rect(w[0], w[1], w[2], w[3])

                if qty_box[0] <= w[0] <= qty_box[1]:
                    if val.isdigit() and len(val) != 8:
                        orig_q = int(val)
                        if orig_q >= factor:
                            new_q = orig_q // factor
                            overwrite_area(page, rect, f"{new_q}")

            # Unit Price update
            for w in row_words:
                val = w[4].replace(",", "").strip()
                rect = fitz.Rect(w[0], w[1], w[2], w[3])

                if price_box[0] <= w[0] <= price_box[1]:
                    if re.match(r"^\d+\.\d{2}$", val):
                        orig_p = float(val)
                        if orig_p > 0.00:
                            new_p = round(orig_price if 'orig_price' in locals() else orig_p * factor, 2)
                            overwrite_area(page, rect, f"{new_p:.2f}")

        # UOM Updates
        for target in ["UOM-PC", "UOM-IBOX", "UOM-PCS", "UOM-BOX"]:
            for inst in page.search_for(target):
                overwrite_area(page, inst, "UOM-BOX")

        for s_inst in page.search_for("S"):
            if 140 <= s_inst.x0 <= 260:
                page.draw_rect(s_inst, color=None, fill=(1, 1, 1))

    out_buffer = io.BytesIO()
    doc.save(out_buffer)
    doc.close()
    out_buffer.seek(0)
    return out_buffer, download_filename

if uploaded_invoices:
    # 1 file ho ya bulk, dono handle karega
    if len(uploaded_invoices) == 1:
        file = uploaded_invoices[0]
        try:
            updated_pdf_buffer, out_filename = process_universal_blinkit_invoice(file.read())
            st.download_button(
                label=f"📥 Download {out_filename}",
                data=updated_pdf_buffer,
                file_name=out_filename,
                mime="application/pdf"
            )
        except Exception as e:
            st.error(f"Error processing {file.name}: {str(e)}")
    else:
        # Multiple files: Process all & bundle into ZIP
        zip_buffer = io.BytesIO()
        processed_files = []

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for file in uploaded_invoices:
                try:
                    pdf_buf, out_fname = process_universal_blinkit_invoice(file.read())
                    # Avoid duplicate filename inside zip
                    if out_fname in processed_files:
                        out_fname = f"{len(processed_files)+1}_{out_fname}"
                    zip_file.writestr(out_fname, pdf_buf.getvalue())
                    processed_files.append(out_fname)
                except Exception as e:
                    st.error(f"Error in {file.name}: {str(e)}")

        zip_buffer.seek(0)
        st.download_button(
            label=f"📦 Download All Invoices ({len(processed_files)} Files - ZIP)",
            data=zip_buffer,
            file_name=f"Blinkit_Processed_Invoices_{datetime.now().strftime('%d-%m-%Y')}.zip",
            mime="application/zip"
        )
