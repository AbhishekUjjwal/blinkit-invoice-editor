import streamlit as st
import fitz  # PyMuPDF
import re
import io
import json
import zipfile
from datetime import datetime

st.set_page_config(page_title="Blinkit Operations Gateway", page_icon="⚡", layout="centered")

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
    .stTabs [data-baseweb="tab-list"] {
        gap: 20px;
        justify-content: center;
    }
    .stTabs [data-baseweb="tab"] {
        font-size: 16px;
        font-weight: 600;
        padding: 10px 20px;
    }
    </style>
    <div class="hero-container">
        <div class="logo-badge"><span class="logo-icon">⚡</span></div>
        <div class="brand-title">Blinkit Operations Gateway</div>
        <div class="brand-sub">PDF Editor & e-Invoice JSON Center</div>
    </div>
""", unsafe_allow_html=True)

# ----------------- HELPER FUNCTIONS -----------------

def overwrite_area(page, rect, new_text, font_size=7):
    pad_rect = fitz.Rect(rect.x0 - 2, rect.y0 - 1, rect.x1 + 2, rect.y1 + 1)
    page.draw_rect(pad_rect, color=None, fill=(1, 1, 1))
    page.insert_text((rect.x0, rect.y1 - 1.2), str(new_text), fontsize=font_size, fontname="helv", color=(0, 0, 0))

def extract_metadata(full_text):
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
    std_date_for_json = datetime.now().strftime("%d/%m/%Y")
    for fmt in ("%b %d, %Y", "%B %d, %Y", "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            d_obj = datetime.strptime(raw_date, fmt)
            clean_date = d_obj.strftime("%d-%m-%Y")
            std_date_for_json = d_obj.strftime("%d/%m/%Y")
            break
        except ValueError:
            pass

    clean_ext = re.sub(r'[^A-Za-z0-9\-_]', '', ext_order_id)
    clean_inv = re.sub(r'[^A-Za-z0-9\-_]', '', invoice_no)
    clean_dt = re.sub(r'[^A-Za-z0-9\-_]', '', clean_date)
    return f"{clean_ext}_{clean_inv}_{clean_dt}.pdf", invoice_no, std_date_for_json

def process_universal_blinkit_invoice(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    full_text = ""
    for p in doc:
        full_text += p.get_text() + "\n"

    download_filename, _, _ = extract_metadata(full_text)

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

            for w in row_words:
                val = w[4].replace(",", "").strip()
                rect = fitz.Rect(w[0], w[1], w[2], w[3])
                if qty_box[0] <= w[0] <= qty_box[1]:
                    if val.isdigit() and len(val) != 8:
                        orig_q = int(val)
                        if orig_q >= factor:
                            new_q = orig_q // factor
                            overwrite_area(page, rect, f"{new_q}")

            for w in row_words:
                val = w[4].replace(",", "").strip()
                rect = fitz.Rect(w[0], w[1], w[2], w[3])
                if price_box[0] <= w[0] <= price_box[1]:
                    if re.match(r"^\d+\.\d{2}$", val):
                        orig_p = float(val)
                        if orig_p > 0.00:
                            new_p = round(orig_p * factor, 2)
                            overwrite_area(page, rect, f"{new_p:.2f}")

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

def build_einvoice_from_edited_pdf(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    full_text = ""
    for p in doc:
        full_text += p.get_text() + "\n"

    _, invoice_no, doc_date = extract_metadata(full_text)

    gstins = re.findall(r"\b\d{2}[A-Z]{5}\d{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}\b", full_text)
    seller_gstin = gstins[0] if len(gstins) > 0 else "09AAFCG9846E1Z9"
    buyer_gstin = gstins[1] if len(gstins) > 1 else seller_gstin

    pincodes = re.findall(r"\b[1-9][0-9]{5}\b", full_text)
    seller_pin = int(pincodes[0]) if len(pincodes) > 0 else 226401
    buyer_pin = int(pincodes[1]) if len(pincodes) > 1 else seller_pin

    seller_state_code = seller_gstin[:2]
    buyer_state_code = buyer_gstin[:2]
    buyer_name = "BLINK COMMERCE PRIVATE LIMITED"

    line_items = []
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
                product_rows.append({"desc": "DGNW50 Dispo Guard Face Mask", "y": (w[1] + w[3]) / 2})
            elif "COMFIT" in text:
                product_rows.append({"desc": "C3DFM Comfit 3D Face Mask", "y": (w[1] + w[3]) / 2})

        for prod in product_rows:
            row_y = prod["y"]
            row_words = [w for w in words if abs(((w[1] + w[3]) / 2) - row_y) <= 25]

            final_q = None
            final_p = None

            for w in row_words:
                val = w[4].replace(",", "").strip()
                if qty_box[0] <= w[0] <= qty_box[1]:
                    if val.isdigit() and len(val) != 8:
                        final_q = int(val)

            for w in row_words:
                val = w[4].replace(",", "").strip()
                if price_box[0] <= w[0] <= price_box[1]:
                    if re.match(r"^\d+\.\d{2}$", val):
                        val_float = float(val)
                        if val_float > 0.00:
                            final_p = val_float

            if final_q and final_p:
                line_items.append({
                    "desc": prod["desc"],
                    "hsn": "63079091",
                    "qty": final_q,
                    "unit_price": final_p,
                    "gst_rate": 5.0
                })

    doc.close()

    item_list = []
    tot_taxable = 0.0
    tot_cgst = 0.0
    tot_sgst = 0.0
    tot_igst = 0.0

    for idx, item in enumerate(line_items, 1):
        qty = item["qty"]
        unit_price = item["unit_price"]
        taxable_amt = round(qty * unit_price, 2)
        tot_taxable += taxable_amt

        gst_rate = item.get("gst_rate", 5.0)
        is_interstate = (seller_state_code != buyer_state_code)

        if is_interstate:
            igst_amt = round((taxable_amt * gst_rate) / 100, 2)
            cgst_amt = 0.0
            sgst_amt = 0.0
            tot_igst += igst_amt
        else:
            cgst_amt = round((taxable_amt * (gst_rate / 2)) / 100, 2)
            sgst_amt = round((taxable_amt * (gst_rate / 2)) / 100, 2)
            igst_amt = 0.0
            tot_cgst += cgst_amt
            tot_sgst += sgst_amt

        tot_item_val = round(taxable_amt + cgst_amt + sgst_amt + igst_amt, 2)

        item_list.append({
            "SlNo": str(idx),
            "PrdDesc": item["desc"][:100],
            "IsServc": "N",
            "HsnCd": item["hsn"],
            "Qty": qty,
            "Unit": "BOX",
            "UnitPrice": unit_price,
            "TotAmt": taxable_amt,
            "Discount": 0.0,
            "AssAmt": taxable_amt,
            "GstRt": gst_rate,
            "IgstAmt": igst_amt,
            "CgstAmt": cgst_amt,
            "SgstAmt": sgst_amt,
            "CesRt": 0.0,
            "CesAmt": 0.0,
            "TotItemVal": tot_item_val
        })

    total_inv_val = round(tot_taxable + tot_cgst + tot_sgst + tot_igst, 2)

    payload = {
        "Version": "1.03",
        "TranDtls": {
            "TaxSch": "GST",
            "SupTyp": "B2B",
            "RegRev": "N",
            "EcmGstin": None,
            "IgstOnIntra": "N"
        },
        "DocDtls": {
            "Typ": "INV",
            "No": invoice_no,
            "Dt": doc_date
        },
        "SellerDtls": {
            "Gstin": seller_gstin,
            "LglNm": "SELLER TRADING CO",
            "TrdNm": "SELLER TRADING CO",
            "Addr1": "Warehouse Address",
            "Loc": "City",
            "Pin": seller_pin,
            "Stcd": seller_state_code
        },
        "BuyerDtls": {
            "Gstin": buyer_gstin,
            "LglNm": buyer_name,
            "TrdNm": buyer_name,
            "Pos": buyer_state_code,
            "Addr1": "Warehouse Facility, Junabganj Road",
            "Loc": "Lucknow",
            "Pin": buyer_pin,
            "Stcd": buyer_state_code
        },
        "ItemList": item_list,
        "ValDtls": {
            "AssVal": round(tot_taxable, 2),
            "CgstVal": round(tot_cgst, 2),
            "SgstVal": round(tot_sgst, 2),
            "IgstVal": round(tot_igst, 2),
            "CesVal": 0.0,
            "StCesVal": 0.0,
            "RndOffAmt": 0.0,
            "TotInvVal": total_inv_val
        }
    }
    return payload, invoice_no

# ----------------- TABS SETUP -----------------

tab1, tab2 = st.tabs(["📑 Step 1: Edit Blinkit Invoice", "🧾 Step 2: Create e-Invoice JSON"])

# TAB 1: Edit original invoices
with tab1:
    st.subheader("Invoice Editing (Qty & Price Fix)")
    uploaded_invoices = st.file_uploader("Original Invoices Upload Karein (Single / Bulk)", type=["pdf"], accept_multiple_files=True, key="pdf_tab_uploader")

    if uploaded_invoices:
        if len(uploaded_invoices) == 1:
            file = uploaded_invoices[0]
            try:
                updated_pdf_buffer, out_filename = process_universal_blinkit_invoice(file.read())
                st.download_button(label=f"📥 Download Edited PDF ({out_filename})", data=updated_pdf_buffer, file_name=out_filename, mime="application/pdf")
            except Exception as e:
                st.error(f"Error: {str(e)}")
        else:
            zip_buffer = io.BytesIO()
            processed_files = []
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for file in uploaded_invoices:
                    try:
                        pdf_buf, out_fname = process_universal_blinkit_invoice(file.read())
                        if out_fname in processed_files:
                            out_fname = f"{len(processed_files)+1}_{out_fname}"
                        zip_file.writestr(out_fname, pdf_buf.getvalue())
                        processed_files.append(out_fname)
                    except Exception as e:
                        st.error(f"Error in {file.name}: {str(e)}")
            zip_buffer.seek(0)
            st.download_button(label=f"📦 Download All Edited Invoices ({len(processed_files)} Files - ZIP)", data=zip_buffer, file_name=f"Blinkit_Edited_Invoices_{datetime.now().strftime('%d-%m-%Y')}.zip", mime="application/zip")

# TAB 2: Upload EDITED invoices to create government ready JSON
with tab2:
    st.subheader("Edited Invoice Upload Karein aur e-Invoice JSON Payein")
    st.info("💡 Yahan aap apni edit ki hui invoices upload kar sakte hain, portal ready JSON direct generate ho jayega.")
    uploaded_edited_invoices = st.file_uploader("Edited Invoices Upload Karein (Single / Bulk)", type=["pdf"], accept_multiple_files=True, key="edited_tab_uploader")

    if uploaded_edited_invoices:
        if len(uploaded_edited_invoices) == 1:
            file = uploaded_edited_invoices[0]
            try:
                einv_json, inv_num = build_einvoice_from_edited_pdf(file.read())
                st.download_button(label=f"🧾 Download e-Invoice JSON ({inv_num})", data=json.dumps(einv_json, indent=4), file_name=f"{inv_num}_eInvoice.json", mime="application/json")
            except Exception as e:
                st.error(f"Error: {str(e)}")
        else:
            batch_list = []
            for file in uploaded_edited_invoices:
                try:
                    einv_json, _ = build_einvoice_from_edited_pdf(file.read())
                    batch_list.append(einv_json)
                except Exception as e:
                    st.error(f"Error in {file.name}: {str(e)}")
            if batch_list:
                st.download_button(label=f"🧾 Download Bulk e-Invoice JSON ({len(batch_list)} Invoices)", data=json.dumps(batch_list, indent=4), file_name=f"Bulk_eInvoice_NIC_{datetime.now().strftime('%d-%m-%Y')}.json", mime="application/json")
