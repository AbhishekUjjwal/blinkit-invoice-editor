import streamlit as st
import fitz  # PyMuPDF
import re
import json
from datetime import datetime

st.set_page_config(page_title="e-Invoice JSON Gateway", page_icon="🧾", layout="centered")

st.markdown("""
    <div style="text-align: center; padding: 10px 0 20px 0;">
        <h2 style="color: #FFFFFF; margin-bottom: 4px;">⚡ Blinkit e-Invoice JSON Gateway</h2>
        <p style="color: #94a3b8; font-size: 14px;">Government Portal Ready JSON Generator (Schema v1.03)</p>
    </div>
""", unsafe_allow_html=True)

uploaded_invoices = st.file_uploader("Upload Blinkit Invoices (PDF)", type=["pdf"], accept_multiple_files=True)

def extract_metadata(full_text):
    inv_match = re.search(r"Invoice\s*No\s*[:\s]*([A-Za-z0-9\-_]+)", full_text, re.IGNORECASE)
    invoice_no = inv_match.group(1).strip() if inv_match else "INV-001"

    date_match = re.search(r"Invoice\s*Date\s*[:\s]*([A-Za-z0-9,\s\.\-\/]+?)(?=\n|Ship\s*Date|$)", full_text, re.IGNORECASE)
    raw_date = date_match.group(1).strip() if date_match else "Date"

    std_date_for_json = datetime.now().strftime("%d/%m/%Y")
    for fmt in ("%b %d, %Y", "%B %d, %Y", "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            d_obj = datetime.strptime(raw_date, fmt)
            std_date_for_json = d_obj.strftime("%d/%m/%Y")
            break
        except ValueError:
            pass

    return invoice_no, std_date_for_json

def build_einvoice_json_object(full_text, invoice_no, doc_date, line_items):
    gstins = re.findall(r"\b\d{2}[A-Z]{5}\d{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}\b", full_text)
    seller_gstin = gstins[0] if len(gstins) > 0 else "09AAFCG9846E1Z9"
    buyer_gstin = gstins[1] if len(gstins) > 1 else seller_gstin

    pincodes = re.findall(r"\b[1-9][0-9]{5}\b", full_text)
    seller_pin = int(pincodes[0]) if len(pincodes) > 0 else 226401
    buyer_pin = int(pincodes[1]) if len(pincodes) > 1 else seller_pin

    seller_state_code = seller_gstin[:2]
    buyer_state_code = buyer_gstin[:2]
    buyer_name = "BLINK COMMERCE PRIVATE LIMITED"

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

    return {
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

def process_pdf_for_json(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    full_text = ""
    for p in doc:
        full_text += p.get_text() + "\n"

    invoice_no, std_date = extract_metadata(full_text)
    extracted_items = []

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
                product_rows.append({"type": "DISPO", "factor": 50, "y": (w[1] + w[3]) / 2, "desc": "DGNW50 Dispo Guard Face Mask"})
            elif "COMFIT" in text:
                product_rows.append({"type": "COMFIT", "factor": 25, "y": (w[1] + w[3]) / 2, "desc": "C3DFM Comfit 3D Face Mask"})

        for prod in product_rows:
            factor = prod["factor"]
            row_y = prod["y"]
            row_words = [w for w in words if abs(((w[1] + w[3]) / 2) - row_y) <= 25]

            final_q = None
            final_p = None

            for w in row_words:
                val = w[4].replace(",", "").strip()
                if qty_box[0] <= w[0] <= qty_box[1]:
                    if val.isdigit() and len(val) != 8:
                        orig_q = int(val)
                        if orig_q >= factor:
                            final_q = orig_q // factor

                if price_box[0] <= w[0] <= price_box[1]:
                    if re.match(r"^\d+\.\d{2}$", val):
                        orig_p = float(val)
                        if orig_p > 0.00:
                            final_p = round(orig_p * factor, 2)

            if final_q and final_p:
                extracted_items.append({
                    "desc": prod["desc"],
                    "hsn": "63079091",
                    "qty": final_q,
                    "unit_price": final_p,
                    "gst_rate": 5.0
                })

    doc.close()
    return build_einvoice_json_object(full_text, invoice_no, std_date, extracted_items), invoice_no

if uploaded_invoices:
    if len(uploaded_invoices) == 1:
        file = uploaded_invoices[0]
        try:
            einv_json, inv_num = process_pdf_for_json(file.read())
            st.success(f"Invoice {inv_num} ready!")
            st.download_button(
                label="🧾 Download e-Invoice JSON",
                data=json.dumps(einv_json, indent=4),
                file_name=f"{inv_num}_eInvoice.json",
                mime="application/json"
            )
        except Exception as e:
            st.error(f"Error: {str(e)}")
    else:
        batch_list = []
        for file in uploaded_invoices:
            try:
                einv_json, _ = process_pdf_for_json(file.read())
                batch_list.append(einv_json)
            except Exception as e:
                st.error(f"Error in {file.name}: {str(e)}")

        if batch_list:
            st.success(f"Total {len(batch_list)} Invoices ready!")
            st.download_button(
                label=f"🧾 Download Bulk e-Invoice JSON ({len(batch_list)} Invoices)",
                data=json.dumps(batch_list, indent=4),
                file_name=f"Bulk_eInvoice_NIC_{datetime.now().strftime('%d-%m-%Y')}.json",
                mime="application/json"
            )
