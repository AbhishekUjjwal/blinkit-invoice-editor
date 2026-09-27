import streamlit as st
import fitz  # PyMuPDF
import re
import io
import json
import zipfile
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Universal Operations & e-Invoice Suite", page_icon="⚡", layout="wide")

st.markdown("""
    <div style="text-align: center; padding: 15px 0 20px 0;">
        <h2 style="color: #FFFFFF; margin-bottom: 6px;">⚡ Universal All-Invoice Operations Suite</h2>
        <p style="color: #94a3b8; font-size: 14px;">Dynamic HSN Column Anchoring | Pure Item Code Description | Exact Precision</p>
    </div>
""", unsafe_allow_html=True)

uploaded_invoices = st.file_uploader("Upload Invoices (PDF) - Single ya Bulk", type=["pdf"], accept_multiple_files=True)

STATE_CODE_MAP = {
    "01": "JAMMU AND KASHMIR", "02": "HIMACHAL PRADESH", "03": "PUNJAB", "04": "CHANDIGARH",
    "05": "UTTARAKHAND", "06": "HARYANA", "07": "DELHI", "08": "RAJASTHAN",
    "09": "UTTAR PRADESH", "10": "BIHAR", "11": "SIKKIM", "12": "ARUNACHAL PRADESH",
    "13": "NAGALAND", "14": "MANIPUR", "15": "MIZORAM", "16": "TRIPURA",
    "17": "MEGHALAYA", "18": "ASSAM", "19": "WEST BENGAL", "20": "JHARKHAND",
    "21": "ODISHA", "22": "CHATTISGARH", "23": "MADHYA PRADESH", "24": "GUJARAT",
    "26": "DADRA AND NAGAR HAVELI AND DAMAN AND DIU", "27": "MAHARASHTRA", "29": "KARNATAKA",
    "30": "GOA", "31": "LAKSHADWEEP", "32": "KERALA", "33": "TAMIL NADU",
    "34": "PUDUCHERRY", "35": "ANDAMAN AND NICOBAR ISLANDS", "36": "TELANGANA", "37": "ANDHRA PRADESH",
    "38": "LADAKH"
}

MAJOR_CITIES = [
    "New Delhi", "Delhi", "Varanasi", "Lucknow", "Jaipur", "Gurgaon", "Gurugram", 
    "Noida", "Ghaziabad", "Kanpur", "Bengaluru", "Bangalore", "Mumbai", "Pune", 
    "Kolkata", "Ahmedabad", "Patna", "Ranchi", "Chandigarh", "Faridabad", "Agra", "Meerut"
]

def clean_extracted_city(text, state_name):
    m = re.search(r"([A-Za-z\s]+),\s*(?:[A-Za-z\s]+)[\-\s]*[0-9]{6}", text)
    if m:
        c_cand = m.group(1).strip()
        parts = [p.strip() for p in c_cand.split(",") if p.strip()]
        last_p = parts[-1]
        if len(last_p) > 2 and not re.search(r"(road|street|nagar|colony|floor|block|house|marg)", last_p, re.I):
            return last_p.title()

    for city in MAJOR_CITIES:
        if re.search(rf"\b{city}\b", text, re.I):
            return city

    return state_name.title() if state_name else "Delhi"

def fmt_dec(val):
    try:
        f = float(str(val).replace(",", "").strip())
        return f"{f:.2f}"
    except (ValueError, TypeError):
        return "0.00"

def clean_description_completely(desc, hsn_code=None, qty=None, unit_price=None):
    if not desc:
        return ""
    if hsn_code:
        desc = re.sub(rf"\b{re.escape(str(hsn_code))}\b", "", desc)
        first_digit = str(hsn_code)[0]
        desc = re.sub(rf"(?<=[a-zA-Z\s,/\-_]){re.escape(first_digit)}$", "", desc)
        desc = re.sub(rf"\b{re.escape(first_digit)}\b", "", desc)

    desc = re.sub(r"\b\d{4,8}\b", "", desc)

    if qty:
        desc = re.sub(rf"\b{re.escape(str(qty))}\b", "", desc)
    if unit_price:
        p_str = f"{float(unit_price):.2f}"
        desc = re.sub(rf"\b{re.escape(p_str)}\b", "", desc)
        desc = re.sub(rf"\b{re.escape(str(unit_price))}\b", "", desc)

    desc = re.sub(r"\b\d+\.\d{2}\b", "", desc)
    desc = re.sub(r"\s+\d$", "", desc)
    desc = re.sub(r"\s+", " ", desc).strip()
    return desc

# ----------------- COORDINATE EXTRACTION -----------------

def extract_metadata_from_doc(doc):
    full_text = ""
    for p in doc:
        full_text += p.get_text() + "\n"

    page0 = doc[0]
    words = page0.get_text("words")

    ext_order_id = None
    ext_match = re.search(r"Extern(?:al)?\s*Order\s*(?:No\.?|ID)?\s*[:\s]*([0-9\sA-Za-z\-]+?)(?=\n|Invoice|Date|$)", full_text, re.IGNORECASE)
    if ext_match:
        ext_order_id = "".join(ext_match.group(1).split())
    if not ext_order_id:
        digits_match = re.findall(r"\b(49\d{8,14}|5\d{8,14}|\d{12,18})\b", full_text)
        if digits_match:
            ext_order_id = digits_match[0]
    if not ext_order_id:
        ext_order_id = "ExtOrder"

    inv_match = re.search(r"Invoice\s*No\s*[:\s]*([A-Za-z0-9\-_/]+)", full_text, re.IGNORECASE)
    invoice_no = inv_match.group(1).strip() if inv_match else "INV-001"

    date_match = re.search(r"Invoice\s*Date\s*[:\s]*([A-Za-z0-9,\s\.\-\/]+?)(?=\n|Ship\s*Date|Order|$)", full_text, re.IGNORECASE)
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
    pdf_filename = f"{clean_ext}_{clean_inv}_{clean_dt}.pdf"

    sold_header = [w for w in words if "sold" in w[4].lower()]
    bill_header = [w for w in words if "billing" in w[4].lower()]
    table_header = [w for w in words if w[4].lower() in ["s.no", "sl.no", "item", "hsn"]]

    y_sold = sold_header[0][1] if sold_header else 80
    y_bill = bill_header[0][1] if bill_header else 215
    y_table = table_header[0][1] if table_header else 340
    split_x = 300

    seller_rect = fitz.Rect(20, y_sold, split_x, y_bill)
    seller_text = page0.get_text("text", clip=seller_rect).strip()
    s_lines = [l.strip() for l in seller_text.split("\n") if l.strip() and not re.search(r"sold\s*by", l, re.I)]
    seller_name = s_lines[0] if s_lines else "Romsons Prime Private Limited"

    s_gst = re.search(r"GSTIN\s*[:\s]*([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})", seller_text, re.I)
    seller_gstin = s_gst.group(1).strip() if s_gst else "07AALCR5906L1ZW"

    s_pin = re.findall(r"\b[1-9][0-9]{5}\b", seller_text)
    seller_pin = int(s_pin[0]) if s_pin else 110028
    seller_state_code = seller_gstin[:2]
    seller_loc = clean_extracted_city(seller_text, STATE_CODE_MAP.get(seller_state_code, "Delhi"))

    billing_rect = fitz.Rect(20, y_bill, split_x, y_table)
    billing_text = page0.get_text("text", clip=billing_rect).strip()
    b_lines = [l.strip() for l in billing_text.split("\n") if l.strip() and not re.search(r"billing\s*addr", l, re.I)]

    buyer_name = b_lines[0] if b_lines else "BUYER ENTERPRISE"
    b_addr_candidates = [l for l in b_lines[1:] if not re.search(r"(@|contact|phone|mob|gstin|pan|india)", l, re.I)]
    buyer_addr1 = ", ".join(b_addr_candidates[:2]) if b_addr_candidates else (b_lines[1] if len(b_lines) > 1 else "Commercial Facility")

    em_m = re.search(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", billing_text)
    buyer_email = em_m.group(0).strip() if em_m else ""

    ph_m = re.search(r"(?:Contact|Phone|Mob)?\s*[:\s]*([6-9]\d{9})", billing_text, re.IGNORECASE)
    buyer_phone = ph_m.group(1).strip() if ph_m else ""

    b_gst_m = re.search(r"GSTIN\s*[:\s]*([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})", billing_text, re.IGNORECASE)
    buyer_gstin = b_gst_m.group(1).strip() if b_gst_m else ""

    if not buyer_gstin:
        all_gstins = re.findall(r"\b\d{2}[A-Z]{5}\d{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}\b", full_text)
        buyer_gstin = all_gstins[1] if len(all_gstins) > 1 else (all_gstins[0] if all_gstins else "07AACCT5680N1ZS")

    buyer_state_code = buyer_gstin[:2]
    buyer_state_name = STATE_CODE_MAP.get(buyer_state_code, "DELHI")

    b_pin_m = re.findall(r"\b[1-9][0-9]{5}\b", billing_text)
    buyer_pin = int(b_pin_m[0]) if b_pin_m else int(buyer_state_code + "0001")
    buyer_loc = clean_extracted_city(billing_text, buyer_state_name)

    shipping_rect = fitz.Rect(split_x + 5, y_bill, 585, y_table)
    shipping_text = page0.get_text("text", clip=shipping_rect).strip()
    shp_lines = [l.strip() for l in shipping_text.split("\n") if l.strip() and not re.search(r"shipping\s*addr", l, re.I)]

    ship_name = shp_lines[0] if shp_lines else buyer_name
    s_addr_candidates = [l for l in shp_lines[1:] if not re.search(r"(@|contact|phone|mob|gstin|pan|india)", l, re.I)]
    ship_addr1 = ", ".join(s_addr_candidates[:2]) if s_addr_candidates else (shp_lines[1] if len(shp_lines) > 1 else buyer_addr1)

    shp_gst_m = re.search(r"GSTIN\s*[:\s]*([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})", shipping_text, re.IGNORECASE)
    ship_gstin = shp_gst_m.group(1).strip() if shp_gst_m else buyer_gstin
    ship_state_code = ship_gstin[:2]
    ship_state_name = STATE_CODE_MAP.get(ship_state_code, buyer_state_name)

    shp_pin_m = re.findall(r"\b[1-9][0-9]{5}\b", shipping_text)
    ship_pin = int(shp_pin_m[0]) if shp_pin_m else buyer_pin
    ship_loc = clean_extracted_city(shipping_text, ship_state_name)

    clean_b_str = re.sub(r'[^A-Za-z0-9]', '', f"{buyer_addr1}{buyer_pin}").lower()
    clean_s_str = re.sub(r'[^A-Za-z0-9]', '', f"{ship_addr1}{ship_pin}").lower()
    shipping_is_different = (clean_b_str != clean_s_str) and (buyer_pin != ship_pin or buyer_addr1 != ship_addr1)

    pos_match = re.search(r"Place\s*of\s*Supply\s*[:\s]*[A-Za-z\s\-]*(\d{2})", full_text, re.IGNORECASE)
    buyer_pos = pos_match.group(1).strip() if pos_match else buyer_state_code

    is_blinkit = bool(re.search(r"blink\s*commerce|blinkit", f"{buyer_name} {full_text}", re.IGNORECASE))

    inv_total_m = re.search(r"(?:Total\s*Invoice\s*Value|Invoice\s*Total|Grand\s*Total|Total\s*Amount)[:\s]*[₹\s]*([0-9,]+\.\d{2})", full_text, re.I)
    printed_grand_total = float(inv_total_m.group(1).replace(",", "")) if inv_total_m else None

    return {
        "invoice_no": invoice_no,
        "doc_date": std_date_for_json,
        "pdf_filename": pdf_filename,
        "seller_name": seller_name,
        "seller_gstin": seller_gstin,
        "seller_pin": seller_pin,
        "seller_loc": seller_loc,
        "seller_state_code": seller_state_code,
        "buyer_name": buyer_name,
        "buyer_trade_name": buyer_name,
        "buyer_addr1": buyer_addr1,
        "buyer_loc": buyer_loc,
        "buyer_pin": buyer_pin,
        "buyer_phone": buyer_phone,
        "buyer_email": buyer_email,
        "buyer_gstin": buyer_gstin,
        "buyer_state_code": buyer_state_code,
        "buyer_state_name": buyer_state_name,
        "buyer_pos": buyer_pos,
        "shipping_is_different": shipping_is_different,
        "ship_name": ship_name,
        "ship_addr1": ship_addr1,
        "ship_loc": ship_loc,
        "ship_pin": ship_pin,
        "ship_gstin": ship_gstin,
        "ship_state_name": ship_state_name,
        "ship_state_code": ship_state_code,
        "is_blinkit": is_blinkit,
        "printed_grand_total": printed_grand_total
    }

# ----------------- PARSING WITH DYNAMIC HSN COLUMN ANCHOR -----------------

def overwrite_area(page, rect, new_text, font_size=7):
    pad_rect = fitz.Rect(rect.x0 - 2, rect.y0 - 1, rect.x1 + 2, rect.y1 + 1)
    page.draw_rect(pad_rect, color=None, fill=(1, 1, 1))
    page.insert_text((rect.x0, rect.y1 - 1.2), str(new_text), fontsize=font_size, fontname="helv", color=(0, 0, 0))

def process_and_reconcile_pdf(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    meta = extract_metadata_from_doc(doc)
    is_blinkit = meta["is_blinkit"]
    line_items = []

    for page in doc:
        words = page.get_text("words")

        # 1. DYNAMIC HEADER COLUMN ANCHORS
        sno_anchor_x = 48
        hsn_min_x = 185
        hsn_max_x = 265
        qty_left = 265
        price_left = 315
        disc_left = 375
        tax_left = 425

        for w in words:
            wt = w[4].lower().strip()
            if wt in ["s.no", "sl.no", "s.no."]:
                sno_anchor_x = w[2] + 4
            elif wt == "hsn":
                # Pinpoint exact HSN column boundaries from header position
                hsn_min_x = w[0] - 12
                hsn_max_x = w[2] + 35
            elif wt == "qty":
                qty_left = w[0] - 8
            elif "price" in wt or "unit" in wt:
                price_left = w[0] - 8
            elif "disc" in wt:
                disc_left = w[0] - 8
            elif "taxable" in wt:
                tax_left = w[0] - 8

        # 2. S.NO CANDIDATE ROWS
        sno_candidates = []
        for w in words:
            val = w[4].strip()
            if val.isdigit() and 1 <= int(val) <= 99:
                if w[0] <= sno_anchor_x and w[1] > 240:
                    if not any(abs(c["center_y"] - ((w[1]+w[3])/2)) < 8 for c in sno_candidates):
                        sno_candidates.append({
                            "sno": int(val),
                            "y0": w[1],
                            "y1": w[3],
                            "center_y": (w[1] + w[3]) / 2
                        })

        sno_candidates.sort(key=lambda x: x["center_y"])

        table_bottom_y = page.rect.height - 50
        total_indicators = [w for w in words if w[1] > 350 and any(k in w[4].lower() for k in ["total", "subtotal", "taxable value", "amount in words"])]
        if total_indicators:
            table_bottom_y = min([w[1] for w in total_indicators]) - 4

        if sno_candidates:
            for idx, item_anchor in enumerate(sno_candidates):
                y_top = item_anchor["y0"] - 3
                if idx + 1 < len(sno_candidates):
                    y_bottom = sno_candidates[idx + 1]["y0"] - 2
                else:
                    y_bottom = min(item_anchor["y1"] + 120, table_bottom_y)

                row_words = [w for w in words if y_top <= ((w[1] + w[3]) / 2) <= y_bottom]

                # 3. EXACT HSN CODE EXTRACTION (Directly clipped from HSN Column)
                hsn_box_rect = fitz.Rect(hsn_min_x, y_top, hsn_max_x, y_bottom)
                hsn_clip_text = page.get_text("text", clip=hsn_box_rect)
                hsn_matches = re.findall(r"\b(\d{6,8})\b", hsn_clip_text)

                if hsn_matches:
                    hsn_code = hsn_matches[0]
                else:
                    # Search inside row words aligned around HSN column
                    candidates = [
                        w[4].replace(",", "").strip() for w in row_words 
                        if (hsn_min_x - 15) <= w[0] <= (hsn_max_x + 15) and w[4].replace(",", "").strip().isdigit()
                    ]
                    valid_hsns = [c for c in candidates if len(c) in [6, 7, 8]]
                    hsn_code = valid_hsns[0] if valid_hsns else "63079091"

                # Find exact X position of this row's HSN to lock item code's right border
                actual_hsn_x = hsn_min_x
                for w in row_words:
                    if w[4].replace(",", "").strip() == hsn_code:
                        actual_hsn_x = min(actual_hsn_x, w[0])
                        break

                final_q = None
                final_p = None
                printed_taxable = None
                printed_discount = 0.0

                # 4. QUANTITY
                qty_box_rect = fitz.Rect(qty_left - 2, y_top, price_left - 2, y_bottom)
                for w in words:
                    if qty_box_rect.contains(fitz.Point((w[0]+w[2])/2, (w[1]+w[3])/2)):
                        val = w[4].replace(",", "").strip()
                        rect = fitz.Rect(w[0], w[1], w[2], w[3])
                        if val.isdigit() and len(val) != 8:
                            orig_q = int(val)
                            factor = 50 if (is_blinkit and "DISPO" in w[4].upper()) else (25 if (is_blinkit and "COMFIT" in w[4].upper()) else 1)
                            if factor > 1 and orig_q >= factor:
                                final_q = orig_q // factor
                                overwrite_area(page, rect, f"{final_q}")
                            else:
                                final_q = orig_q
                            break

                # 5. UNIT PRICE
                price_box_rect = fitz.Rect(price_left - 2, y_top, disc_left - 2, y_bottom)
                for w in words:
                    if price_box_rect.contains(fitz.Point((w[0]+w[2])/2, (w[1]+w[3])/2)):
                        val = w[4].replace(",", "").strip()
                        rect = fitz.Rect(w[0], w[1], w[2], w[3])
                        if re.match(r"^\d+\.\d{2}$", val):
                            orig_p = float(val)
                            if orig_p > 0.00:
                                factor = 50 if (is_blinkit and "DISPO" in w[4].upper()) else (25 if (is_blinkit and "COMFIT" in w[4].upper()) else 1)
                                if factor > 1:
                                    final_p = round(orig_p * factor, 2)
                                    overwrite_area(page, rect, f"{final_p:.2f}")
                                else:
                                    final_p = orig_p
                                break

                # 6. DISCOUNT
                disc_box_rect = fitz.Rect(disc_left - 2, y_top, tax_left - 2, y_bottom)
                raw_disc_text = page.get_text("text", clip=disc_box_rect)
                disc_m = re.findall(r"\b\d+\.\d{2}\b", raw_disc_text)
                if disc_m:
                    printed_discount = float(disc_m[0])

                # 7. TAXABLE VALUE
                tax_box_rect = fitz.Rect(tax_left - 2, y_top, tax_left + 85, y_bottom)
                raw_tax_text = page.get_text("text", clip=tax_box_rect)
                tax_m = re.findall(r"\b\d+\.\d{2}\b", raw_tax_text)
                if tax_m:
                    printed_taxable = float(tax_m[0])

                # 8. PURE ITEM CODE (Boundary strictly locked left of actual HSN start)
                item_right_boundary = actual_hsn_x - 4
                item_box_rect = fitz.Rect(sno_anchor_x, y_top, item_right_boundary, y_bottom)
                raw_item_text = page.get_text("text", clip=item_box_rect).strip()

                lines = [l.strip() for l in raw_item_text.split("\n") if l.strip()]
                raw_desc = " ".join(lines).strip()

                pure_desc = clean_description_completely(raw_desc, hsn_code=hsn_code, qty=final_q, unit_price=final_p)
                if not pure_desc:
                    pure_desc = f"Item {item_anchor['sno']}"

                factor = 1
                if is_blinkit:
                    upper_desc = pure_desc.upper()
                    if "DISPO" in upper_desc:
                        factor = 50
                    elif "COMFIT" in upper_desc:
                        factor = 25

                if final_q is not None and final_p is not None:
                    gross_amt = round(final_q * final_p, 2)
                    if printed_taxable is not None and printed_taxable > 0:
                        taxable_val = printed_taxable
                    else:
                        taxable_val = round(gross_amt - printed_discount, 2)

                    unit_label = "BOX" if (is_blinkit and factor > 1) else "PCS"
                    line_items.append({
                        "sno": item_anchor["sno"],
                        "desc": pure_desc,
                        "hsn": hsn_code,
                        "qty": final_q,
                        "unit": unit_label,
                        "unit_price": final_p,
                        "gross_amt": gross_amt,
                        "discount": printed_discount,
                        "taxable_val": taxable_val,
                        "gst_rate": 5.0
                    })

        if is_blinkit:
            for target in ["UOM-PC", "UOM-IBOX", "UOM-PCS", "UOM-BOX"]:
                for inst in page.search_for(target):
                    overwrite_area(page, inst, "UOM-BOX")

            for s_inst in page.search_for("S"):
                if 140 <= s_inst.x0 <= 260:
                    page.draw_rect(s_inst, color=None, fill=(1, 1, 1))

    out_pdf_buf = io.BytesIO()
    doc.save(out_pdf_buf)
    doc.close()
    out_pdf_buf.seek(0)

    meta["line_items"] = line_items
    return out_pdf_buf, meta

# ----------------- JSON v1.01 BUILDER -----------------

def build_einvoice_json_v101(data):
    seller_state_code = data["seller_state_code"]
    buyer_pos = data["buyer_pos"]

    item_list = []
    tot_taxable = 0.0
    tot_cgst = 0.0
    tot_sgst = 0.0
    tot_igst = 0.0

    for idx, item in enumerate(data["line_items"], 1):
        qty = item["qty"]
        unit_price = item["unit_price"]
        unit_label = item.get("unit", "BOX")
        taxable_amt = item.get("taxable_val", round(qty * unit_price, 2))
        tot_taxable += taxable_amt

        gst_rate = item.get("gst_rate", 5.0)
        is_interstate = (seller_state_code != buyer_pos)

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
            "PrdDesc": item["desc"][:250],
            "IsServc": "N",
            "HsnCd": item["hsn"],
            "Qty": qty,
            "Unit": unit_label,
            "UnitPrice": round(unit_price, 2),
            "TotAmt": round(item.get("gross_amt", taxable_amt), 2),
            "Discount": round(item.get("discount", 0.0), 2),
            "AssAmt": round(taxable_amt, 2),
            "GstRt": gst_rate,
            "IgstAmt": igst_amt,
            "CgstAmt": cgst_amt,
            "SgstAmt": sgst_amt,
            "CesRt": 0.0,
            "CesAmt": 0.0,
            "TotItemVal": tot_item_val
        })

    calc_total = round(tot_taxable + tot_cgst + tot_sgst + tot_igst, 2)
    final_total = data.get("printed_grand_total") if data.get("printed_grand_total") else calc_total

    ship_dtls = None
    if data.get("shipping_is_different", False):
        ship_dtls = {
            "Gstin": data["ship_gstin"],
            "LglNm": data["ship_name"],
            "TrdNm": data["ship_name"],
            "Addr1": data["ship_addr1"][:100],
            "Loc": data["ship_loc"],
            "Pin": data["ship_pin"],
            "Stcd": data["ship_state_code"]
        }

    return {
        "Version": "1.01",
        "TranDtls": {
            "TaxSch": "GST",
            "SupTyp": "B2B",
            "RegRev": "N",
            "EcmGstin": None,
            "IgstOnIntra": "N"
        },
        "DocDtls": {
            "Typ": "INV",
            "No": data["invoice_no"],
            "Dt": data["doc_date"]
        },
        "SellerDtls": {
            "Gstin": data["seller_gstin"],
            "LglNm": data["seller_name"],
            "TrdNm": data["seller_name"],
            "Addr1": f"Warehouse Facility, {data['seller_loc']}",
            "Loc": data["seller_loc"],
            "Pin": data["seller_pin"],
            "Stcd": seller_state_code
        },
        "BuyerDtls": {
            "Gstin": data["buyer_gstin"],
            "LglNm": data["buyer_name"],
            "TrdNm": data["buyer_trade_name"],
            "Pos": buyer_pos,
            "Addr1": data["buyer_addr1"][:100],
            "Loc": data["buyer_loc"],
            "Pin": data["buyer_pin"],
            "Stcd": buyer_pos
        },
        "DispDtls": None,
        "ShipDtls": ship_dtls,
        "ItemList": item_list,
        "ValDtls": {
            "AssVal": round(tot_taxable, 2),
            "CgstVal": round(tot_cgst, 2),
            "SgstVal": round(tot_sgst, 2),
            "IgstVal": round(tot_igst, 2),
            "CesVal": 0.0,
            "StCesVal": 0.0,
            "RndOffAmt": round(final_total - calc_total, 2),
            "TotInvVal": round(final_total, 2)
        }
    }

# ----------------- OFFICIAL NIC v1.01 EXCEL BUILDER (TEXT 2-DECIMAL) -----------------

def generate_official_nic_v101_excel(data_list):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "eInvoice"
    ws.views.sheetView[0].showGridLines = True

    sky_blue = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
    light_blue = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    peach = PatternFill(start_color="FDE9D9", end_color="FDE9D9", fill_type="solid")
    green = PatternFill(start_color="D8E4BC", end_color="D8E4BC", fill_type="solid")

    font_title = Font(name="Arial", size=18, bold=True)
    font_button = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    font_sec = Font(name="Arial", size=10, bold=True)
    font_col = Font(name="Arial", size=9, bold=True)
    font_data = Font(name="Arial", size=9)

    thin_border = Border(
        left=Side(style='thin', color='B0B0B0'),
        right=Side(style='thin', color='B0B0B0'),
        top=Side(style='thin', color='B0B0B0'),
        bottom=Side(style='thin', color='B0B0B0')
    )

    ws["G1"] = "E-Invoice System"
    ws["G1"].font = font_title
    ws["G1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("L1:N2")
    ws["L1"] = "Validate"
    ws["L1"].fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    ws["L1"].font = font_button
    ws["L1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("P1:R2")
    ws["P1"] = "Prepare JSON"
    ws["P1"].fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    ws["P1"].font = font_button
    ws["P1"].alignment = Alignment(horizontal="center", vertical="center")

    ws["T2"] = "* Indicates Mandatory Fields"
    ws["T2"].font = Font(name="Arial", size=8, italic=True, color="555555")

    sections = [
        ("Supply Details", 1, 4, light_blue),
        ("Document Details", 5, 7, peach),
        ("Buyer Details", 8, 18, light_blue),
        ("Dispatch Details", 19, 24, peach),
        ("Shipping Details", 25, 32, light_blue),
        ("Item Details", 33, 46, green),
        ("Invoice Value Details", 47, 52, sky_blue)
    ]

    for name, sc, ec, fill in sections:
        ws.merge_cells(start_row=3, start_column=sc, end_row=3, end_column=ec)
        c = ws.cell(row=3, column=sc, value=name)
        c.fill = fill
        c.font = font_sec
        c.alignment = Alignment(horizontal="center", vertical="center")
        for i in range(sc, ec + 1):
            ws.cell(row=3, column=i).border = thin_border

    columns = [
        "Supply Type *", "Reverse Charge", "e-Comm GSTIN", "Igst On Intra",
        "Document Type *", "Document Number *", "Document Date (DD/MM/YYYY) *",
        "Buyer GSTIN *", "Buyer Legal Name *", "Buyer Trade Name", "Buyer POS *", 
        "Buyer Addr1 *", "Buyer Addr2", "Buyer Location *", "Buyer Pin Code *", 
        "Buyer State *", "Buyer Phone Number", "Buyer Email Id",
        "Dispatch Name", "Dispatch Addr1", "Dispatch Addr2", "Dispatch Location", "Dispatch Pin Code", "Dispatch State",
        "Shipping GSTIN", "Shipping Legal Name", "Shipping Trade Name", "Shipping Addr1", "Shipping Addr2", "Shipping Location", "Shipping Pin Code", "Shipping State",
        "Sl. No *", "Product Description *", "Is Service *", "HSN Code *", "Quantity *", "Unit *", "Unit Price *", 
        "Gross Amount", "Taxable Value *", "GST Rate (%) *", "IGST Amount", "CGST Amount", "SGST Amount", "Total Item Value *",
        "Total Taxable Value *", "Total CGST Amount", "Total SGST Amount", "Total IGST Amount", "Round Off Amount", "Total Invoice Value *"
    ]

    for c_idx, col_name in enumerate(columns, 1):
        cell = ws.cell(row=4, column=c_idx, value=col_name)
        if c_idx <= 4:
            cell.fill = light_blue
        elif c_idx <= 7:
            cell.fill = peach
        elif c_idx <= 18:
            cell.fill = light_blue
        elif c_idx <= 24:
            cell.fill = peach
        elif c_idx <= 32:
            cell.fill = light_blue
        elif c_idx <= 46:
            cell.fill = green
        else:
            cell.fill = sky_blue
        cell.font = font_col
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    ws.row_dimensions[4].height = 42

    curr_row = 5
    for inv in data_list:
        seller_state_code = inv["seller_state_code"]
        buyer_pos = inv["buyer_pos"]

        tot_taxable = sum([it["taxable_val"] for it in inv["line_items"]])
        is_interstate = (seller_state_code != buyer_pos)

        if is_interstate:
            tot_igst = round(tot_taxable * 0.05, 2)
            tot_cgst = 0.0
            tot_sgst = 0.0
        else:
            tot_igst = 0.0
            tot_cgst = round(tot_taxable * 0.025, 2)
            tot_sgst = round(tot_taxable * 0.025, 2)

        calc_inv_val = round(tot_taxable + tot_cgst + tot_sgst + tot_igst, 2)
        final_inv_val = inv.get("printed_grand_total") if inv.get("printed_grand_total") else calc_inv_val

        if inv.get("shipping_is_different", False):
            ship_vals = [
                str(inv["ship_gstin"]), str(inv["ship_name"]), str(inv["ship_name"]),
                str(inv["ship_addr1"]), "", str(inv["ship_loc"]), str(inv["ship_pin"]), str(inv["ship_state_name"])
            ]
        else:
            ship_vals = ["", "", "", "", "", "", "", ""]

        for s_no, it in enumerate(inv["line_items"], 1):
            qty = it["qty"]
            price = it["unit_price"]
            gross_amt = it.get("gross_amt", round(qty * price, 2))
            taxable = it.get("taxable_val", gross_amt)
            gst_rate = it.get("gst_rate", 5.0)

            if is_interstate:
                igst = round(taxable * (gst_rate / 100), 2)
                cgst = 0.0
                sgst = 0.0
            else:
                cgst = round(taxable * (gst_rate / 200), 2)
                sgst = round(taxable * (gst_rate / 200), 2)
                igst = 0.0

            item_val = round(taxable + cgst + sgst + igst, 2)

            row_data = [
                # Supply Details
                "B2B", "N", "", "N",
                # Document Details
                "Tax Invoice", str(inv["invoice_no"]), str(inv["doc_date"]),
                # Buyer Details
                str(inv["buyer_gstin"]), str(inv["buyer_name"]), str(inv["buyer_trade_name"]), str(buyer_pos),
                str(inv["buyer_addr1"]), "", str(inv["buyer_loc"]), str(inv["buyer_pin"]),
                str(inv["buyer_state_name"]), str(inv["buyer_phone"]), str(inv["buyer_email"]),
                # Dispatch Details (Blank)
                "", "", "", "", "", "",
                # Shipping Details
                *ship_vals,
                # Item Details (Pure Description & Exact Dynamic HSN Code)
                str(s_no), str(it["desc"]), "N", str(it["hsn"]), str(qty), str(it.get("unit", "BOX")),
                fmt_dec(price), fmt_dec(gross_amt), fmt_dec(taxable), str(int(gst_rate)),
                fmt_dec(igst), fmt_dec(cgst), fmt_dec(sgst), fmt_dec(item_val),
                # Invoice Value Details
                fmt_dec(tot_taxable), fmt_dec(tot_cgst), fmt_dec(tot_sgst), fmt_dec(tot_igst),
                fmt_dec(final_inv_val - calc_inv_val), fmt_dec(final_inv_val)
            ]

            for c_idx, val in enumerate(row_data, 1):
                cell = ws.cell(row=curr_row, column=c_idx, value=str(val))
                cell.font = font_data
                cell.border = thin_border
                cell.number_format = '@'
                if c_idx in [1, 2, 4, 5, 7, 8, 11, 15, 16, 17, 25, 31, 32, 33, 35, 36, 38, 42]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif c_idx in [37, 39, 40, 41, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52]:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

            curr_row += 1

    tabs = ["Welcome", "Profile", "Master Codes", "Sample Invoice", "Format A,B,C,D", "Schema", "Validation", "Calculations", "FAQs"]
    for t in tabs:
        d_ws = wb.create_sheet(title=t)
        d_ws.sheet_view.showGridLines = True
        d_ws["A1"] = f"{t} - NIC e-Invoice Offline Utility v1.01"
        d_ws["A1"].font = Font(size=14, bold=True, color="1F497D")

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row < 3:
                continue
            v_str = str(cell.value or '')
            if len(v_str) > max_len:
                max_len = len(v_str)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 11)

    out_io = io.BytesIO()
    wb.save(out_io)
    out_io.seek(0)
    return out_io

# ----------------- MAIN STREAMLIT WORKFLOW -----------------

if uploaded_invoices:
    processed_docs = []
    for file in uploaded_invoices:
        try:
            pdf_bytes = file.read()
            reconciled_pdf_buf, meta = process_and_reconcile_pdf(pdf_bytes)
            processed_docs.append({
                "pdf_buf": reconciled_pdf_buf,
                "meta": meta
            })
        except Exception as e:
            st.error(f"Error processing {file.name}: {str(e)}")

    if processed_docs:
        st.success(f"✅ Successfully processed {len(processed_docs)} Invoices!")
        
        parsed_invoices = [doc["meta"] for doc in processed_docs]
        excel_buffer = generate_official_nic_v101_excel(parsed_invoices)

        # SINGLE INVOICE ACTIONS
        if len(processed_docs) == 1:
            doc = processed_docs[0]
            meta = doc["meta"]
            pdf_buf = doc["pdf_buf"]
            json_payload = build_einvoice_json_v101(meta)
            inv_no = meta["invoice_no"]
            pdf_name = meta["pdf_filename"]

            ship_status = "⚠️ Different (Shipping Details Filled)" if meta["shipping_is_different"] else "✅ Same (Shipping Blank in Excel)"

            tot_tax = sum([it['taxable_val'] for it in meta['line_items']])
            disp_total = meta['printed_grand_total'] if meta['printed_grand_total'] else round(tot_tax * 1.05, 2)

            st.markdown(f"""
                <div style="background-color: #1e293b; padding: 14px 18px; border-radius: 8px; margin-bottom: 15px; border-left: 5px solid #3b82f6;">
                    <b>Buyer Name:</b> <code>{meta['buyer_name']}</code> &nbsp;|&nbsp; <b>City:</b> <code>{meta['buyer_loc']}</code> ({meta['buyer_pin']})<br>
                    <b>Shipping Status:</b> <code>{ship_status}</code> &nbsp;|&nbsp; <b>Total Invoice Value:</b> <code>₹{disp_total:.2f}</code><br>
                    <b>Seller:</b> {meta['seller_name']} ({meta['seller_gstin']})
                </div>
            """, unsafe_allow_html=True)

            preview_data = [{
                "Sl": it["sno"],
                "Item Code (Pure Description)": it["desc"],
                "HSN Code": it["hsn"],
                "Qty": it["qty"],
                "Unit Price": fmt_dec(it['unit_price']),
                "Gross Amt": fmt_dec(it.get('gross_amt', 0.0)),
                "Discount": fmt_dec(it.get('discount', 0.0)),
                "Taxable Value": fmt_dec(it.get('taxable_val', 0.0))
            } for it in meta['line_items']]
            
            st.write("📋 **Verified Line Items (Dynamic HSN & Clean Description):**")
            st.table(preview_data)

            c1, c2, c3 = st.columns(3)
            with c1:
                st.download_button(
                    label=f"📥 Download Processed PDF",
                    data=pdf_buf,
                    file_name=pdf_name,
                    mime="application/pdf"
                )
            with c2:
                st.download_button(
                    label=f"📊 Download NIC Bulk Excel (v1.01)",
                    data=excel_buffer,
                    file_name=f"{inv_no}_NIC_v1.01.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with c3:
                st.download_button(
                    label=f"🧾 Download e-Invoice JSON (v1.01)",
                    data=json.dumps(json_payload, indent=4),
                    file_name=f"{inv_no}_eInvoice_v1.01.json",
                    mime="application/json"
                )

        # BULK INVOICE ACTIONS
        else:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for idx, doc in enumerate(processed_docs, 1):
                    pdf_name = doc["meta"]["pdf_filename"]
                    zip_file.writestr(pdf_name, doc["pdf_buf"].getvalue())
            zip_buffer.seek(0)

            bulk_json = [build_einvoice_json_v101(d["meta"]) for d in processed_docs]
            date_str = datetime.now().strftime('%d-%m-%Y')

            c1, c2, c3 = st.columns(3)
            with c1:
                st.download_button(
                    label=f"📦 Download All Processed PDFs ({len(processed_docs)} Files - ZIP)",
                    data=zip_buffer,
                    file_name=f"Invoices_Processed_{date_str}.zip",
                    mime="application/zip"
                )
            with c2:
                st.download_button(
                    label=f"📊 Download Bulk NIC Excel v1.01 ({len(processed_docs)} Invoices)",
                    data=excel_buffer,
                    file_name=f"Bulk_NIC_v1.01_{date_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with c3:
                st.download_button(
                    label=f"🧾 Download Bulk e-Invoice JSON ({len(bulk_json)} Invoices)",
                    data=json.dumps(bulk_json, indent=4),
                    file_name=f"Bulk_eInvoice_NIC_v1.01_{date_str}.json",
                    mime="application/json"
                )
