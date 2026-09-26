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

st.set_page_config(page_title="Universal Invoice & e-Invoice Suite", page_icon="⚡", layout="wide")

st.markdown("""
    <div style="text-align: center; padding: 15px 0 20px 0;">
        <h2 style="color: #FFFFFF; margin-bottom: 6px;">⚡ Universal e-Invoice & Operations Gateway</h2>
        <p style="color: #94a3b8; font-size: 14px;">100% Dynamic Buyer & Seller Extraction | PDF Reconcile + NIC v1.01 Bulk Excel + e-Invoice JSON</p>
    </div>
""", unsafe_allow_html=True)

uploaded_invoices = st.file_uploader("Upload Invoices (PDF) - Single ya Bulk", type=["pdf"], accept_multiple_files=True)

# GST State Code to State Name Master Dictionary
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

# ----------------- 100% DYNAMIC DATA EXTRACTION -----------------

def extract_metadata(full_text):
    # External Order ID
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

    # Invoice No
    inv_match = re.search(r"Invoice\s*No\s*[:\s]*([A-Za-z0-9\-_/]+)", full_text, re.IGNORECASE)
    invoice_no = inv_match.group(1).strip() if inv_match else "INV-001"

    # Invoice Date
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
    pdf_filename = f"{clean_ext}_{clean_inv}_{clean_dt}.pdf"

    # 1. DYNAMIC SELLER EXTRACTION (Sold By block)
    seller_name = ""
    seller_addr1 = ""
    seller_loc = ""
    seller_pin = 110028
    seller_gstin = ""

    seller_block_m = re.search(r"(?:Sold\s*By|Seller\s*Details|From)[:\s]*\n(.*?)(?=\n\s*(?:Billing|Billed|Order\s*No|Invoice\s*No|\Z))", full_text, re.DOTALL | re.IGNORECASE)
    if seller_block_m:
        s_block = seller_block_m.group(1).strip()
        s_lines = [l.strip() for l in s_block.split("\n") if l.strip()]
        if len(s_lines) > 0:
            seller_name = s_lines[0]
        if len(s_lines) > 1:
            seller_addr1 = s_lines[1]

        s_gst_m = re.search(r"GSTIN\s*[:\s]*([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})", s_block, re.IGNORECASE)
        if s_gst_m:
            seller_gstin = s_gst_m.group(1).strip()

        s_pin_m = re.findall(r"\b[1-9][0-9]{5}\b", s_block)
        if s_pin_m:
            seller_pin = int(s_pin_m[0])

    if not seller_gstin:
        all_gstins = re.findall(r"\b\d{2}[A-Z]{5}\d{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}\b", full_text)
        seller_gstin = all_gstins[0] if len(all_gstins) > 0 else "07AALCR5906L1ZW"

    seller_state_code = seller_gstin[:2]
    seller_loc = STATE_CODE_MAP.get(seller_state_code, "Delhi").title()

    # 2. DYNAMIC BUYER EXTRACTION (Billing Address / Billed To block)
    buyer_name = ""
    buyer_addr1 = ""
    buyer_loc = ""
    buyer_pin = None
    buyer_phone = ""
    buyer_email = ""
    buyer_gstin = ""

    buyer_block_m = re.search(r"(?:Billing\s*Addr[a-z]*|Billed\s*To|Buyer\s*Details|Customer\s*Details)[:\s]*\n(.*?)(?=\n\s*(?:Shipping|Ship\s*To|S\.No|Item Code|H\s*S\s*N|\Z))", full_text, re.DOTALL | re.IGNORECASE)
    if buyer_block_m:
        b_block = buyer_block_m.group(1).strip()
        b_lines = [l.strip() for l in b_block.split("\n") if l.strip()]

        if len(b_lines) > 0:
            buyer_name = b_lines[0]
        if len(b_lines) > 1:
            buyer_addr1 = b_lines[1]

        # Extract Email
        em_m = re.search(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", b_block)
        if em_m:
            buyer_email = em_m.group(0).strip()

        # Extract Phone
        ph_m = re.search(r"(?:Contact|Phone|Mob)?\s*[:\s]*([6-9]\d{9})", b_block, re.IGNORECASE)
        if ph_m:
            buyer_phone = ph_m.group(1).strip()

        # Extract GSTIN
        b_gst_m = re.search(r"GSTIN\s*[:\s]*([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})", b_block, re.IGNORECASE)
        if b_gst_m:
            buyer_gstin = b_gst_m.group(1).strip()

        # Extract Pin Code
        b_pin_m = re.findall(r"\b[1-9][0-9]{5}\b", b_block)
        if b_pin_m:
            buyer_pin = int(b_pin_m[0])

        # Extract City / Location dynamically from address lines
        for l in b_lines[1:]:
            clean_l = re.sub(r"[0-9\-,]", " ", l).strip()
            words = [w for w in clean_l.split() if len(w) > 3 and not re.search(r"(road|station|near|opp|street|nagar|floor|block|india|contact|gstin|pan|email)", w, re.I)]
            if words:
                buyer_loc = words[0].title()
                break

    # Fallbacks if buyer GSTIN wasn't in billing block
    if not buyer_gstin:
        all_gstins = re.findall(r"\b\d{2}[A-Z]{5}\d{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}\b", full_text)
        if len(all_gstins) > 1:
            buyer_gstin = all_gstins[1]
        elif len(all_gstins) == 1:
            buyer_gstin = all_gstins[0]
        else:
            buyer_gstin = "09AAFCG9846E1Z9"

    buyer_state_code = buyer_gstin[:2]
    buyer_state_name = STATE_CODE_MAP.get(buyer_state_code, "UTTAR PRADESH")
    if not buyer_loc:
        buyer_loc = buyer_state_name.title()
    if not buyer_pin:
        buyer_pin = int(buyer_state_code + "0001") if buyer_state_code.isdigit() else 226401

    # Place of Supply (POS)
    pos_match = re.search(r"Place\s*of\s*Supply\s*[:\s]*[A-Za-z\s\-]*(\d{2})", full_text, re.IGNORECASE)
    buyer_pos = pos_match.group(1).strip() if pos_match else buyer_state_code

    return {
        "invoice_no": invoice_no,
        "doc_date": std_date_for_json,
        "pdf_filename": pdf_filename,
        "seller_name": seller_name if seller_name else "SELLER ENTERPRISE",
        "seller_gstin": seller_gstin,
        "seller_pin": seller_pin,
        "seller_loc": seller_loc,
        "seller_state_code": seller_state_code,
        "buyer_name": buyer_name if buyer_name else "BUYER ENTERPRISE",
        "buyer_trade_name": buyer_name if buyer_name else "BUYER ENTERPRISE",
        "buyer_addr1": buyer_addr1 if buyer_addr1 else "Commercial Facility",
        "buyer_loc": buyer_loc,
        "buyer_pin": buyer_pin,
        "buyer_phone": buyer_phone,
        "buyer_email": buyer_email,
        "buyer_gstin": buyer_gstin,
        "buyer_state_code": buyer_state_code,
        "buyer_state_name": buyer_state_name,
        "buyer_pos": buyer_pos
    }

# ----------------- PDF RECONCILIATION -----------------

def overwrite_area(page, rect, new_text, font_size=7):
    pad_rect = fitz.Rect(rect.x0 - 2, rect.y0 - 1, rect.x1 + 2, rect.y1 + 1)
    page.draw_rect(pad_rect, color=None, fill=(1, 1, 1))
    page.insert_text((rect.x0, rect.y1 - 1.2), str(new_text), fontsize=font_size, fontname="helv", color=(0, 0, 0))

def process_and_reconcile_pdf(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    full_text = ""
    for p in doc:
        full_text += p.get_text() + "\n"

    meta = extract_metadata(full_text)
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
                rect = fitz.Rect(w[0], w[1], w[2], w[3])
                if qty_box[0] <= w[0] <= qty_box[1]:
                    if val.isdigit() and len(val) != 8:
                        orig_q = int(val)
                        if orig_q >= factor:
                            final_q = orig_q // factor
                            overwrite_area(page, rect, f"{final_q}")

            for w in row_words:
                val = w[4].replace(",", "").strip()
                rect = fitz.Rect(w[0], w[1], w[2], w[3])
                if price_box[0] <= w[0] <= price_box[1]:
                    if re.match(r"^\d+\.\d{2}$", val):
                        orig_p = float(val)
                        if orig_p > 0.00:
                            final_p = round(orig_p * factor, 2)
                            overwrite_area(page, rect, f"{final_p:.2f}")

            if final_q and final_p:
                line_items.append({
                    "desc": prod["desc"],
                    "hsn": "63079091",
                    "qty": final_q,
                    "unit_price": final_p,
                    "gst_rate": 5.0
                })

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
        taxable_amt = round(qty * unit_price, 2)
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
        "ShipDtls": None,
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

# ----------------- OFFICIAL NIC v1.01 EXCEL BUILDER -----------------

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

        tot_taxable = sum([round(it["qty"] * it["unit_price"], 2) for it in inv["line_items"]])
        is_interstate = (seller_state_code != buyer_pos)

        if is_interstate:
            tot_igst = round(tot_taxable * 0.05, 2)
            tot_cgst = 0.0
            tot_sgst = 0.0
        else:
            tot_igst = 0.0
            tot_cgst = round(tot_taxable * 0.025, 2)
            tot_sgst = round(tot_taxable * 0.025, 2)

        tot_inv_val = round(tot_taxable + tot_cgst + tot_sgst + tot_igst, 2)

        for s_no, it in enumerate(inv["line_items"], 1):
            qty = it["qty"]
            price = it["unit_price"]
            taxable = round(qty * price, 2)
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
                "Tax Invoice", inv["invoice_no"], inv["doc_date"],
                # Buyer Details (100% Dynamic)
                inv["buyer_gstin"], inv["buyer_name"], inv["buyer_trade_name"], buyer_pos,
                inv["buyer_addr1"], "", inv["buyer_loc"], inv["buyer_pin"],
                inv["buyer_state_name"], inv["buyer_phone"], inv["buyer_email"],
                # Dispatch Details (KEPT BLANK)
                "", "", "", "", "", "",
                # Shipping Details (KEPT BLANK)
                "", "", "", "", "", "", "", "",
                # Item Details
                s_no, it["desc"], "N", it["hsn"], qty, "BOX", price,
                taxable, taxable, gst_rate, igst, cgst, sgst, item_val,
                # Invoice Value Details
                tot_taxable, tot_cgst, tot_sgst, tot_igst, 0.00, tot_inv_val
            ]

            for c_idx, val in enumerate(row_data, 1):
                cell = ws.cell(row=curr_row, column=c_idx, value=val)
                cell.font = font_data
                cell.border = thin_border
                if c_idx in [1, 2, 4, 5, 7, 8, 11, 15, 16, 17, 33, 35, 36, 38]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif c_idx == 37:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.number_format = "#,##0"
                elif c_idx in [39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52]:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.number_format = "#,##0.00"
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

# ----------------- MAIN WORKFLOW -----------------

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

            st.markdown(f"""
                <div style="background-color: #1e293b; padding: 12px 18px; border-radius: 8px; margin-bottom: 15px;">
                    <b>Seller:</b> {meta['seller_name']} ({meta['seller_gstin']})<br>
                    <b>Buyer:</b> {meta['buyer_name']} ({meta['buyer_gstin']})<br>
                    <b>Address:</b> {meta['buyer_addr1']}, {meta['buyer_loc']} - {meta['buyer_pin']} ({meta['buyer_state_name']})
                </div>
            """, unsafe_allow_html=True)

            c1, c2, c3 = st.columns(3)
            with c1:
                st.download_button(
                    label=f"📥 Download Edited PDF",
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
                    label=f"📦 Download All Edited PDFs ({len(processed_docs)} Files - ZIP)",
                    data=zip_buffer,
                    file_name=f"Invoices_Edited_{date_str}.zip",
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
