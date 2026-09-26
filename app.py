import streamlit as st
import fitz  # PyMuPDF
import re
import io
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
        <div class="brand-title">Blinkit Invoice Gateway</div>
        <div class="brand-sub">Universal Dynamic Reconciliation Engine</div>
    </div>
""", unsafe_allow_html=True)

uploaded_invoice = st.file_uploader("", type=["pdf"])

def overwrite_area(page, rect, new_text, font_size=7):
    pad_rect = fitz.Rect(rect.x0 - 1.5, rect.y0 - 0.5, rect.x1 + 1.5, rect.y1 + 0.5)
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

        qty_col_x = None
        price_col_x = None

        for w in words:
            text = w[4].strip()
            if text.lower() == "qty":
                qty_col_x = (w[0], w[2])
            elif "unit" in text.lower() or "price" in text.lower():
                if not price_col_x:
                    price_col_x = (w[0], w[2])

        qty_x_range = (qty_col_x[0] - 15, qty_col_x[1] + 25) if qty_col_x else (330, 390)
        price_x_range = (price_col_x[0] - 15, price_col_x[1] + 35) if price_col_x else (390, 460)

        blocks = page.get_text("blocks")
        detected_rows = []

        for b in blocks:
            b_text = b[4].upper()
            factor = None
            if "DISPO" in b_text:
                factor = 50
            elif "COMFIT" in b_text:
                factor = 25

            if factor:
                detected_rows.append({
                    "factor": factor,
                    "y_top": b[1] - 4,
                    "y_bottom": b[3] + 4
                })

        for row in detected_rows:
            factor = row["factor"]
            y0, y1 = row["y_top"], row["y_bottom"]

            for w in words:
                w_rect = fitz.Rect(w[0], w[1], w[2], w[3])
                w_val = w[4].replace(",", "").strip()

                if y0 <= w_rect.y0 and w_rect.y1 <= y1 + 8:
                    if qty_x_range[0] <= w_rect.x0 <= qty_x_range[1]:
                        if w_val.isdigit() and int(w_val) >= factor:
                            orig_qty = int(w_val)
                            new_qty = orig_qty // factor
                            overwrite_area(page, w_rect, f"{new_qty}")

                    elif price_x_range[0] <= w_rect.x0 <= price_x_range[1]:
                        if re.match(r"^\d+(\.\d+)?$", w_val):
                            orig_price = float(w_val)
                            if orig_price > 0:
                                new_price = round(orig_price * factor, 2)
                                overwrite_area(page, w_rect, f"{new_price:.2f}")

        for target in ["UOM-PC", "UOM-IBOX", "UOM-PCS", "UOM-BOX"]:
            for inst in page.search_for(target):
                overwrite_area(page, inst, "UOM-BOX")

        for s_inst in page.search_for("S"):
            if 150 <= s_inst.x0 <= 250:
                page.draw_rect(s_inst, color=None, fill=(1, 1, 1))

    out_buffer = io.BytesIO()
    doc.save(out_buffer)
    doc.close()
    out_buffer.seek(0)
    return out_buffer, download_filename

if uploaded_invoice:
    st.info("Invoice analyze ho raha hai...")
    try:
        updated_pdf_buffer, out_filename = process_universal_blinkit_invoice(uploaded_invoice.read())
        st.success(f"File ready: {out_filename}")
        st.download_button(
            label=f"📥 Download {out_filename}",
            data=updated_pdf_buffer,
            file_name=out_filename,
            mime="application/pdf"
        )
    except Exception as e:
        st.error(f"Processing Error: {str(e)}")
