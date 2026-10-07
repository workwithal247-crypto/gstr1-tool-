# =====================================================================
# GSTR-1 SUMMARY TOOL  (free, rule-based)  -  built for Tally-style invoices
# Reads PDF invoices -> writes GSTR1_Summary.xlsx
# Works in Google Colab (phone) or on a computer.
# =====================================================================
import subprocess, sys, re, os, glob
from collections import defaultdict

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ---------------- SETTINGS (change here if needed) -------------------
DEFAULT_UQC = "NOS-NUMBERS"        # invoices don't show unit, so we assume this
B2CL_LIMIT = 250000                # inter-state unregistered invoice limit
OUTPUT_FILE = "GSTR1_Summary.xlsx"
# ----------------------------------------------------------------------

STATES = {"01":"Jammu & Kashmir","02":"Himachal Pradesh","03":"Punjab","04":"Chandigarh",
"05":"Uttarakhand","06":"Haryana","07":"Delhi","08":"Rajasthan","09":"Uttar Pradesh",
"10":"Bihar","11":"Sikkim","12":"Arunachal Pradesh","13":"Nagaland","14":"Manipur",
"15":"Mizoram","16":"Tripura","17":"Meghalaya","18":"Assam","19":"West Bengal",
"20":"Jharkhand","21":"Odisha","22":"Chhattisgarh","23":"Madhya Pradesh","24":"Gujarat",
"26":"Dadra & Nagar Haveli and Daman & Diu","27":"Maharashtra","29":"Karnataka","30":"Goa",
"31":"Lakshadweep","32":"Kerala","33":"Tamil Nadu","34":"Puducherry","35":"Andaman & Nicobar",
"36":"Telangana","37":"Andhra Pradesh","38":"Ladakh"}

GSTIN_RE = re.compile(r"^\d{2}[A-Z]{5}\d{4}[A-Z][1-9A-Z]Z[0-9A-Z]$")
CODE = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"

def gstin_checksum_ok(g):
    total = 0
    for i, ch in enumerate(g[:14]):
        v = CODE.index(ch) * (1 if i % 2 == 0 else 2)
        total += v // 36 + v % 36
    return CODE[(36 - total % 36) % 36] == g[14]

def num(s):
    s = re.sub(r"[^\d.]", "", s or "")
    return float(s) if s else 0.0

def find(pattern, text, flags=0):
    m = re.search(pattern, text, flags)
    return m.group(1).strip() if m else ""

def parse_invoice(text, source):
    """Pull the fields out of one invoice's text."""
    inv = {"source": source, "notes": []}
    # Seller part = before 'Bill To', Buyer part = after
    parts = re.split(r"Bill To[^\n]*\n", text, maxsplit=1)
    seller, buyer = parts[0], (parts[1] if len(parts) > 1 else "")
    inv["seller_gstin"] = find(r"GSTIN:\s*([0-9A-Z]{15})", seller)
    inv["inv_no"] = find(r"Invoice No\.?:\s*(\S+)", text)
    d = find(r"Invoice Date:\s*(\d{2}-\d{2}-\d{4})", text)
    inv["date"] = d
    pos = find(r"Place of Supply:\s*([^\n]+)", text)
    inv["pos_code"] = find(r"\((\d{2})\)", pos)
    inv["reverse"] = "Y" if re.search(r"Reverse Charge:\s*Yes", text, re.I) else "N"
    blines = [l.strip() for l in buyer.split("\n") if l.strip()]
    inv["buyer"] = blines[0] if blines else ""
    inv["buyer_gstin"] = find(r"GSTIN:\s*([0-9A-Z]{15})", buyer)
    # line items
    items = []
    for m in re.finditer(r"^\s*(\d+)\s+(.+?)\s+(\d{4,8})\s+(\d+(?:\.\d+)?)\s+\D?([\d,]+\.\d{2})\s+\D?([\d,]+\.\d{2})\s*$",
                         text, re.M):
        items.append({"desc": m.group(2), "hsn": m.group(3), "qty": float(m.group(4)),
                      "amount": num(m.group(6))})
    inv["items"] = items
    # tax totals
    inv["taxable"] = num(find(r"Taxable Value\s*(?:</b>)?\s*\D?([\d,]+\.\d{2})", text))
    for t in ("CGST", "SGST", "IGST", "CESS"):
        inv[t] = num(find(t + r"[^\n]*?\D([\d,]+\.\d{2})\s*$", text, re.M | re.I)) \
            if re.search(r"^(<b>)?" + t, text, re.M | re.I) else 0.0
    inv["total"] = num(find(r"Grand Total\s*(?:</b>)?\s*\D?([\d,]+\.\d{2})", text))
    # tax rate
    m = re.search(r"(CGST|IGST)\s*@\s*(\d+(?:\.\d+)?)\s*%", text, re.I)
    if m:
        r = float(m.group(2))
        inv["rate"] = r * 2 if m.group(1).upper() == "CGST" else r
    else:
        inv["rate"] = 0.0
    return inv

def validate(inv, seen):
    n = inv["notes"]          # problems (go to Review sheet)
    for f, label in (("inv_no", "Invoice number"), ("date", "Invoice date"),
                     ("taxable", "Taxable value"), ("pos_code", "Place of supply")):
        if not inv[f]:
            n.append(f"{label} not found")
    if not inv["items"]:
        n.append("No line items read")
    if inv["buyer_gstin"]:
        g = inv["buyer_gstin"]
        if not GSTIN_RE.match(g):
            n.append(f"Buyer GSTIN format wrong: {g}")
        elif not gstin_checksum_ok(g):
            n.append(f"Buyer GSTIN fails check-digit test: {g}")
    if inv["items"]:
        s = round(sum(i["amount"] for i in inv["items"]), 2)
        if abs(s - inv["taxable"]) > 1:
            n.append(f"Items add to {s:,.2f} but Taxable Value is {inv['taxable']:,.2f}")
    exp_tax = round(inv["taxable"] * inv["rate"] / 100, 2)
    got_tax = round(inv["CGST"] + inv["SGST"] + inv["IGST"], 2)
    if abs(exp_tax - got_tax) > 1:
        n.append(f"Tax should be {exp_tax:,.2f} at {inv['rate']}% but invoice shows {got_tax:,.2f}")
    if abs(round(inv["taxable"] + got_tax + inv["CESS"], 2) - inv["total"]) > 1:
        n.append("Grand Total does not match Taxable + Tax")
    if inv["seller_gstin"] and inv["pos_code"]:
        same = inv["seller_gstin"][:2] == inv["pos_code"]
        if same and inv["IGST"] > 0:
            n.append("Same-state supply but IGST charged")
        if not same and (inv["CGST"] > 0 or inv["SGST"] > 0):
            n.append("Other-state supply but CGST/SGST charged")
    if inv["inv_no"] in seen:
        n.append(f"Duplicate invoice number (also in {seen[inv['inv_no']]})")
    seen.setdefault(inv["inv_no"], inv["source"])

def read_all(pdf_paths):
    invoices = []
    for path in pdf_paths:
        with pdfplumber.open(path) as pdf:
            texts = [(p.extract_text() or "") for p in pdf.pages]
        # each invoice starts on a page containing 'Invoice No'; extra pages are joined to it
        current = None
        for i, t in enumerate(texts, 1):
            if re.search(r"Invoice No", t):
                if current:
                    invoices.append(current)
                current = [t, f"{os.path.basename(path)} p{i}"]
            elif current:
                current[0] += "\n" + t
            elif t.strip():
                invoices.append([t, f"{os.path.basename(path)} p{i}"])
        if current:
            invoices.append(current)
        if not any(t.strip() for t in texts):
            invoices.append(["", f"{os.path.basename(path)} (scanned/no text)"])
    return invoices

# ----------------------------- EXCEL ---------------------------------
HEAD_FILL = PatternFill("solid", fgColor="1F4E78")
HEAD_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
BODY = Font(name="Arial", size=10)
BOLD = Font(name="Arial", size=10, bold=True)
MONEY = "#,##0.00"

def write_sheet(ws, headers, rows, money_cols=(), total_cols=(), widths=None):
    ws.append(headers)
    for c in ws[1]:
        c.font, c.fill = HEAD_FONT, HEAD_FILL
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for r in rows:
        ws.append(r)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = BODY
            if c.column in money_cols:
                c.number_format = MONEY
    if rows and total_cols:
        tr = len(rows) + 2
        ws.cell(tr, 1, "TOTAL").font = BOLD
        for col in total_cols:
            L = get_column_letter(col)
            c = ws.cell(tr, col, f"=SUM({L}2:{L}{tr-1})")
            c.font, c.number_format = BOLD, MONEY
    for i, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(i)].width = (widths or {}).get(i, max(14, len(h) + 2))
    ws.freeze_panes = "A2"

def build_excel(invoices, out):
    wb = Workbook()
    b2b, b2cs, b2cl, hsn, review, allinv = [], [], [], defaultdict(lambda: [0, 0.0, 0.0, 0.0, 0.0, 0.0, ""]), [], []
    b2cs_group = defaultdict(lambda: [0.0, 0.0, 0.0, 0.0])
    for inv in invoices:
        pos = f"{inv['pos_code']}-{STATES.get(inv['pos_code'], '?')}" if inv["pos_code"] else ""
        igst, cgst, sgst = inv["IGST"], inv["CGST"], inv["SGST"]
        status = "REVIEW" if inv["notes"] else "OK"
        allinv.append([inv["source"], inv["inv_no"], inv["date"], inv["buyer"], inv["buyer_gstin"],
                       pos, inv["rate"], inv["taxable"], igst, cgst, sgst, inv["CESS"], inv["total"], status])
        if inv["notes"]:
            review.append([inv["source"], inv["inv_no"], inv["buyer"], "; ".join(inv["notes"])])
        if not inv["inv_no"]:
            continue
        if inv["buyer_gstin"]:
            b2b.append([inv["buyer_gstin"], inv["buyer"], inv["inv_no"], inv["date"], inv["total"], pos,
                        inv["reverse"], "Regular", inv["rate"], inv["taxable"], igst, cgst, sgst, inv["CESS"]])
        else:
            inter = inv["seller_gstin"][:2] != inv["pos_code"] if inv["seller_gstin"] else False
            if inter and inv["total"] > B2CL_LIMIT:
                b2cl.append([inv["inv_no"], inv["date"], inv["total"], pos, inv["rate"], inv["taxable"], igst, inv["CESS"]])
            else:
                g = b2cs_group[(pos, inv["rate"], "INTRA" if not inter else "INTER")]
                g[0] += inv["taxable"]; g[1] += igst; g[2] += cgst; g[3] += sgst
        # HSN summary: spread invoice tax over items in proportion to their value
        s_items = sum(i["amount"] for i in inv["items"]) or 1
        for it in inv["items"]:
            share = it["amount"] / s_items
            h = hsn[(it["hsn"], inv["rate"])]
            h[0] += it["qty"]; h[2] += it["amount"]; h[3] += igst * share
            h[4] += cgst * share; h[5] += sgst * share
            h[6] = it["desc"] if not h[6] else h[6]

    ws = wb.active; ws.title = "B2B"
    write_sheet(ws, ["GSTIN/UIN of Recipient", "Receiver Name", "Invoice Number", "Invoice date", "Invoice Value",
                     "Place Of Supply", "Reverse Charge", "Invoice Type", "Rate", "Taxable Value",
                     "IGST", "CGST", "SGST", "Cess"], b2b,
                money_cols=(5, 10, 11, 12, 13, 14), total_cols=(5, 10, 11, 12, 13, 14), widths={2: 26})

    ws = wb.create_sheet("B2CS")
    rows = [[("OE" if k[2] == "INTER" else "OE"), k[0], k[1], v[0], v[1], v[2], v[3]]
            for k, v in sorted(b2cs_group.items())]
    write_sheet(ws, ["Type", "Place Of Supply", "Rate", "Taxable Value", "IGST", "CGST", "SGST"], rows,
                money_cols=(4, 5, 6, 7), total_cols=(4, 5, 6, 7))
    ws = wb.create_sheet("B2CL")
    write_sheet(ws, ["Invoice Number", "Invoice date", "Invoice Value", "Place Of Supply", "Rate",
                     "Taxable Value", "IGST", "Cess"], b2cl, money_cols=(3, 6, 7, 8), total_cols=(3, 6, 7))

    ws = wb.create_sheet("HSN Summary")
    rows = []
    for (code, rate), v in sorted(hsn.items()):
        rows.append([code, v[6], DEFAULT_UQC, v[0], round(v[2] + v[3] + v[4] + v[5], 2), v[2], rate,
                     round(v[3], 2), round(v[4], 2), round(v[5], 2)])
    write_sheet(ws, ["HSN", "Description", "UQC", "Total Quantity", "Total Value", "Taxable Value", "Rate",
                     "IGST", "CGST", "SGST"], rows, money_cols=(5, 6, 8, 9, 10), total_cols=(5, 6, 8, 9, 10),
                widths={2: 32})

    ws = wb.create_sheet("All Invoices")
    write_sheet(ws, ["Source", "Invoice No", "Date", "Buyer", "Buyer GSTIN", "Place of Supply", "Rate %",
                     "Taxable", "IGST", "CGST", "SGST", "Cess", "Invoice Total", "Status"], allinv,
                money_cols=(8, 9, 10, 11, 12, 13), total_cols=(8, 9, 10, 11, 12, 13), widths={4: 26})

    ws = wb.create_sheet("Review")
    if not review:
        review = [["", "", "", "No problems found. Still spot-check a few invoices."]]
    write_sheet(ws, ["Source", "Invoice No", "Buyer", "Problem found"], review, widths={3: 24, 4: 90})
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")

    ws = wb.create_sheet("Notes")
    for line in ["Assumptions in this workbook:",
                 f"1. Unit (UQC) is not printed on invoices, so '{DEFAULT_UQC}' is assumed in HSN Summary.",
                 "2. HSN tax split is proportional to item value within each invoice.",
                 "3. Rate % = CGST% x 2 (or IGST%) read from the invoice.",
                 "4. B2CS 'Type' is OE (Other than E-commerce). Change if you sell via e-commerce.",
                 "5. Credit/debit notes, exports, advances and nil-rated sales are NOT handled yet.",
                 "6. Always compare totals with your books before filing on the GST portal."]:
        ws.append([line]); ws.cell(ws.max_row, 1).font = BODY
    ws.column_dimensions["A"].width = 100
    wb.save(out)


# ------------------------- WEB APP SCREEN ------------------------------
import tempfile
import streamlit as st

st.set_page_config(page_title="GSTR-1 Summary Tool", page_icon="🧾")
st.title("🧾 GSTR-1 Summary Tool")
st.write("Upload your sales invoice PDFs and get an Excel summary for GSTR-1.")

uploaded = st.file_uploader("Choose invoice PDFs", type="pdf", accept_multiple_files=True)

if uploaded and st.button("Build Excel summary", type="primary"):
    with tempfile.TemporaryDirectory() as tmp:
        paths = []
        for f in uploaded:
            p = os.path.join(tmp, f.name)
            with open(p, "wb") as fh:
                fh.write(f.getbuffer())
            paths.append(p)
        with st.spinner("Reading invoices..."):
            invoices = [parse_invoice(t, s) for t, s in read_all(paths)]
            seen = {}
            for inv in invoices:
                validate(inv, seen)
            out = os.path.join(tmp, OUTPUT_FILE)
            build_excel(invoices, out)
            data = open(out, "rb").read()
    bad = [i for i in invoices if i["notes"]]
    c1, c2, c3 = st.columns(3)
    c1.metric("Invoices read", len(invoices))
    c2.metric("Need review", len(bad))
    c3.metric("Taxable value", f"{sum(i['taxable'] for i in invoices):,.2f}")
    st.download_button("⬇️ Download Excel", data, file_name=OUTPUT_FILE,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if bad:
        st.warning("Please check these invoices (also listed in the Review sheet):")
        for i in bad:
            st.write(f"**{i['inv_no'] or i['source']}**: " + "; ".join(i["notes"]))
    else:
        st.success("No problems found. Still spot-check a few invoices.")
st.caption("Always compare totals with your books before filing on the GST portal.")
