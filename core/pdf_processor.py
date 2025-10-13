# -*- coding: utf-8 -*-
"""
PDF Processing Module
Converted from ExportNike_XuatKhau_AllInOne.py for web app use
"""

import os
import re
import logging
from datetime import datetime
import pdfplumber
import json

logger = logging.getLogger(__name__)

# ========================= Utils =========================
def norm(s: str) -> str:
    """Normalize string by removing extra whitespace"""
    return re.sub(r"\s+", " ", (s or "").strip())

def to_dmy(date_str: str) -> str:
    """Convert date string to dd/mm/yyyy format"""
    if not date_str:
        return ""
    
    # Support multiple date formats
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%b %d, %Y", "%B %d, %Y"):
        try:
            dt = datetime.strptime(date_str.strip(), fmt)
            return f"{dt.day:02d}/{dt.month:02d}/{dt.year}"
        except ValueError:
            continue
    
    try:
        # Handle formats like "August 20, 2025"
        dt = datetime.strptime(date_str.replace("  ", " ").replace(" ,", ",").strip(), "%B %d, %Y")
        return f"{dt.day:02d}/{dt.month:02d}/{dt.year}"
    except Exception:
        return date_str

# ==================== Marks & DEST ======================
def group_lines(words, tol=3):
    """Group words into lines based on y-coordinate"""
    if not words:
        return []
    words = sorted(words, key=lambda w: (round(w["top"]), w["x0"]))
    lines, cur, cur_y = [], [], None
    for w in words:
        y = round(w["top"])
        if cur_y is None or abs(y - cur_y) <= tol:
            cur.append(w)
            cur_y = y if cur_y is None else cur_y
        else:
            lines.append(sorted(cur, key=lambda t: t["x0"]))
            cur, cur_y = [w], y
    if cur:
        lines.append(sorted(cur, key=lambda t: t["x0"]))
    return lines

def linetxt(line_words):
    """Convert line words to text"""
    return " ".join(w["text"] for w in line_words)

RE_MARKS_CODE = re.compile(r"[A-Z0-9][A-Z0-9\-/# ]{5,}", re.IGNORECASE)

def extract_marks_by_words(page) -> str:
    """Extract marks from PDF page using word analysis"""
    try:
        words = page.extract_words(x_tolerance=1, y_tolerance=1, keep_blank_chars=False)
    except Exception:
        words = []
    if not words:
        return ""
    
    lines = group_lines(words, tol=2)
    x_left = None
    
    # Find Country of Origin section
    for line in lines:
        txt = linetxt(line)
        if re.search(r"\bCountry\s+Of\s+Origin\b", txt, re.IGNORECASE):
            x_left = min(w["x0"] for w in line if re.search(r"Country|Of|Origin", w["text"], re.I)) - 6
            break
    
    if x_left is None:
        x_left = page.width * 0.70

    right_lines, texts = [], []
    for line in lines:
        right_words = [w for w in line if w["x0"] >= x_left]
        if right_words:
            right_lines.append(right_words)
            texts.append(linetxt(right_words))
    
    if not right_lines:
        return ""

    r_coo_idx = next((i for i, t in enumerate(texts) if re.search(r"\bCountry\s+Of\s+Origin\b", t, re.IGNORECASE)), None)
    r_start = max(0, (r_coo_idx - 3 if r_coo_idx is not None else 0))
    
    for k in range((r_coo_idx or len(texts)) - 1, max(-1, (r_coo_idx or len(texts)) - 6), -1):
        if k < 0:
            break
        if re.search(r"\bNOCAB\b|\bMARKS?\b", texts[k], re.IGNORECASE):
            r_start = k
            break
    
    r_end = r_coo_idx + 1 if r_coo_idx is not None and r_coo_idx + 1 < len(texts) else (r_coo_idx or (r_start + 2))
    r_end = min(r_end, len(texts) - 1)
    picked = norm(" ".join(texts[r_start:r_end + 1]))
    
    m = re.search(r"(.*?Country\s+Of\s+Origin\s*:\s*[A-Za-z ]+)", picked, re.IGNORECASE)
    if m:
        return norm(m.group(1))
    
    m2 = RE_MARKS_CODE.search(picked)
    if m2:
        return norm(m2.group(0))
    
    return picked

def extract_marks(page) -> str:
    """Extract marks from PDF page"""
    mk = extract_marks_by_words(page)
    return mk or ""

# Country name simplification
COUNTRY_SIMPLE = [
    "Canada", "Belgium", "Poland", "Taiwan", "Singapore", "United Kingdom", "United States",
    "Netherlands", "France", "Germany", "Italy", "Spain", "Portugal", "Mexico", "Brazil", "China",
    "Japan", "Korea", "Turkey", "Australia", "New Zealand", "Thailand", "Indonesia", "Malaysia",
    "Philippines", "Vietnam", "Cambodia", "Laos", "Myanmar", "India", "Pakistan", "Bangladesh",
    "Sri Lanka", "U.A.E", "United Arab Emirates",
]

def clean_long_country_name(s: str) -> str:
    """Clean and simplify country names"""
    t = s.strip()
    t = re.sub(r"^(kingdom|republic|state|federal|commonwealth)\s+of\s+", "", t, flags=re.I)
    t = re.sub(r"\s*\(.*?\)\s*$", "", t)
    t = re.sub(r"^people's republic of\s+", "", t, flags=re.I)
    t = (t.replace("Viet Nam", "Vietnam")
           .replace("Korea, Republic of", "Korea")
           .replace("U.S.A", "United States"))
    return t.strip()

def dest_from_marks(marks: str) -> str:
    """Extract destination country from marks"""
    if not marks:
        return ""
    
    pre = marks
    if "Country Of Origin" in marks:
        pre = marks.split("Country Of Origin", 1)[0]
    
    found = ""
    for name in COUNTRY_SIMPLE:
        if re.search(rf"\b{name}\b", pre, flags=re.I):
            found = name
    
    if found:
        return clean_long_country_name(found)
    
    tail = re.findall(r"([A-Z][a-z]+(?:\s+[A-Z][a-z]+){0,2})\s*$", pre)
    return clean_long_country_name(tail[-1]) if tail else ""

# ==================== PL Parser ====================
# Regex patterns for PO and Reference PO extraction
RE_REFPO = re.compile(
    r"Reference\s+PO#\s*:\s*(\d{7,})"
    r"(?=\s*(?:Item\s+Seq\.|Material|Desc|Customer\s+Ship\s+To|Plant|Total\s+CBM|Total\s+Net\s+Kgs|Total\s+Gross\s+Kgs|\Z))",
    re.IGNORECASE,
)
RE_PO_ANY = re.compile(r"\bPO\W*#?\W*[: ]\s*(\d{7,})", re.IGNORECASE)

def parse_pl_pdf(pdf_path: str) -> dict:
    """Parse PL PDF and extract invoice data"""
    row = {
        "Invoice Number": "", "Date": "", "PO": "", "Reference PO#": "", "Item Seq.": "",
        "Material": "", "Desc": "", "Customer Ship To #": "", "Plant": "",
        "Total Gross Kgs": "", "Total Cartons": "", "Total Units": "", "Marks": "",
    }
    refpo_hits, po_hits = [], []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                txt = page.extract_text() or ""

                # Extract Invoice Number
                m = re.search(r"Invoice\s+Number\.?\s*:\s*([A-Z0-9\-\/]+)", txt, flags=re.IGNORECASE)
                if m: 
                    row["Invoice Number"] = norm(m.group(1))

                # Extract Date
                m = re.search(r"\bDate\s*:\s*([0-3]?\d[\/\-][01]?\d[\/\-]\d{4})", txt, flags=re.IGNORECASE)
                if not m:
                    m = re.search(r"\bDate\s*:\s*([A-Za-z]+ \d{1,2},\s*\d{4})", txt, flags=re.IGNORECASE)
                if m: 
                    row["Date"] = to_dmy(m.group(1))

                # Extract PO numbers
                for mm in RE_REFPO.finditer(txt):
                    refpo_hits.append(mm.group(1))
                for mm in RE_PO_ANY.finditer(txt):
                    po_hits.append(mm.group(1))

                def grab(pat):
                    m2 = re.search(pat, txt, flags=re.IGNORECASE)
                    return norm(m2.group(1)) if m2 else ""

                # Extract other fields
                row["Item Seq."]          = row["Item Seq."]          or grab(r"\bItem\s+Seq\.?\s*:\s*([A-Z0-9]+)")
                row["Material"]           = row["Material"]           or grab(r"\bMaterial\s*:\s*([A-Z0-9\-]+)")
                row["Desc"]               = row["Desc"]               or grab(
                    r"\bDesc\s*:\s*(.+?)(?=\s+(?:Total\s+Net\s+Kgs|Total\s+Gross\s+Kgs|Reference\s+PO#|Plant|Total\s+CBM|Item\s+Seq\.|Customer\s+Ship\s+To)\b)"
                )
                row["Customer Ship To #"] = row["Customer Ship To #"] or grab(r"\bCustomer\s+Ship\s+To\s*:?\s*#?\s*:\s*(\d{6,})")
                row["Plant"]              = row["Plant"]              or grab(r"\bPlant\s*:\s*([A-Z0-9\-]+)")
                row["Total Gross Kgs"]    = row["Total Gross Kgs"]    or grab(r"\bTotal\s+Gross\s+Kgs\s*:\s*([\d\.,]+)")
                row["Total Cartons"]      = row["Total Cartons"]      or grab(r"\bTotal\s+Cartons\s*:\s*([\d\.,]+)")
                row["Total Units"]        = row["Total Units"]        or grab(r"\bTotal\s+Units\s*:\s*([\d\.,]+)")

                # Extract marks
                mk = extract_marks(page)
                if mk:
                    row["Marks"] = norm(mk)
                    # PO might be in marks
                    for mm in RE_PO_ANY.finditer(mk):
                        po_hits.append(mm.group(1))

        # Set PO values
        if refpo_hits:
            row["Reference PO#"] = refpo_hits[0]
        elif po_hits:
            row["Reference PO#"] = po_hits[0]

        if po_hits:
            row["PO"] = po_hits[0]

    except Exception as e:
        logger.error(f"Error parsing PL PDF {pdf_path}: {str(e)}")
        
    return row

# ==================== ASI (BOOKING) DESCRIPTION ====================
DESC_ROW1 = re.compile(r"(?i)\bNIKE\s+\d{7,}(?:-\d+)?\s+([A-Z][A-Z0-9 /&\-]{2,})\s+\d+\s+\d+\s+\d+(?:\.\d+)?")
DESC_ROW2 = re.compile(r"(?i)\b\d{7,}(?:-\d+)?\s+([A-Z][A-Z0-9 /&\-]{2,})\s+\d+\s+\d+\s+\d+(?:\.\d+)?")
DESC_LABEL = re.compile(r"(?i)\bDESCRIPTION\s*[:：]?\s*([A-Z][A-Z0-9 /&\-]{2,})")

def _clean_desc(d: str) -> str:
    """Clean extracted description"""
    d = norm(d).upper()
    d = re.sub(r"\b\d+[A-Z]*\b", "", d)          # Remove quantities
    d = re.sub(r"[^A-Z ]", " ", d)
    d = re.sub(r"\s+", " ", d).strip()
    
    for k in ("SKULL CAP", "HAT", "CAP", "BEANIE", "SCULL CAP"):
        if k in d:
            return "SKULL CAP" if k == "SCULL CAP" else ("BEANIE" if k == "BEANIE" else k)
    return d

def extract_asi_description(pdf_path: str):
    """Extract description from ASI (BOOKING) PDF"""
    invs, descs = set(), set()
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                txt = page.extract_text() or ""
                
                # Find invoice numbers
                invs |= {m.group(0).upper() for m in re.finditer(r"\bA\d{6,8}[A-Z]?\b", txt)}
                
                flat = txt.replace("\u00A0", " ").replace("\n", " ")

                # Extract descriptions using patterns
                for rx in (DESC_ROW1, DESC_ROW2):
                    for m in rx.finditer(flat):
                        descs.add(_clean_desc(m.group(1)))

                for m in DESC_LABEL.finditer(txt):
                    descs.add(_clean_desc(m.group(1)))

        # If no invoice found in content, try filename
        if not invs:
            m = re.search(r"\bA\d{6,8}[A-Z]?\b", os.path.basename(pdf_path), re.IGNORECASE)
            if m: 
                invs.add(m.group(0).upper())

    except Exception as e:
        logger.error(f"Error extracting ASI description from {pdf_path}: {str(e)}")

    # Prioritize certain descriptions
    if descs:
        priority = ["SKULL CAP", "HAT", "CAP"]
        desc = sorted(descs, key=lambda s: priority.index(s) if s in priority else 99)[0]
    else:
        desc = ""
    
    return list(invs), desc

# ==================== Main Processing Functions ====================
def process_pdf_files(pl_file=None, booking_file=None, invoice_file=None):
    """Process PDF files and return extracted data"""
    result = {
        'pl_data': {},
        'asi_description': '',
        'invoice_numbers': [],
        'dest': '',
        'errors': []
    }
    
    try:
        # Process PL file
        if pl_file and os.path.exists(pl_file):
            logger.info(f"Processing PL file: {pl_file}")
            result['pl_data'] = parse_pl_pdf(pl_file)
            
            # Extract destination from marks
            marks = result['pl_data'].get('Marks', '')
            if marks:
                result['dest'] = dest_from_marks(marks)
        
        # Process BOOKING file
        if booking_file and os.path.exists(booking_file):
            logger.info(f"Processing BOOKING file: {booking_file}")
            invs, desc = extract_asi_description(booking_file)
            result['asi_description'] = desc
            result['invoice_numbers'].extend(invs)
        
        # Add invoice number from PL data if available
        if result['pl_data'].get('Invoice Number'):
            inv_num = result['pl_data']['Invoice Number']
            if inv_num not in result['invoice_numbers']:
                result['invoice_numbers'].append(inv_num)
                
    except Exception as e:
        error_msg = f"Error processing PDF files: {str(e)}"
        logger.error(error_msg)
        result['errors'].append(error_msg)
    
    return result

def get_invoice_data_dict(pl_data, asi_description='', dest=''):
    """Convert parsed data to dictionary format for template replacement"""
    data = {}
    
    # Copy all PL data
    for key, value in pl_data.items():
        data[key] = value
    
    # Add additional fields
    if asi_description:
        data['DESCRIPTION'] = asi_description
    
    if dest:
        data['DEST'] = dest
        data['Destination'] = dest
    
    # Convert Total Cartons to words (similar to VBA function)
    try:
        if 'Total Cartons' in data and data['Total Cartons']:
            cartons = float(str(data['Total Cartons']).replace(',', ''))
            data['Total Cartons In Words'] = number_to_words(int(cartons))
    except (ValueError, TypeError):
        data['Total Cartons In Words'] = ''
    
    # Ensure Marks is uppercase
    if 'Marks' in data:
        data['Marks'] = str(data['Marks']).upper()
    
    return data

def number_to_words(n):
    """Convert number to words (basic implementation)"""
    if n == 0:
        return "ZERO"
    
    ones = ["", "ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX", "SEVEN", "EIGHT", "NINE"]
    teens = ["TEN", "ELEVEN", "TWELVE", "THIRTEEN", "FOURTEEN", "FIFTEEN", 
             "SIXTEEN", "SEVENTEEN", "EIGHTEEN", "NINETEEN"]
    tens = ["", "", "TWENTY", "THIRTY", "FORTY", "FIFTY", "SIXTY", "SEVENTY", "EIGHTY", "NINETY"]
    
    def convert_hundreds(num):
        result = ""
        
        if num >= 100:
            result += ones[num // 100] + " HUNDRED "
            num %= 100
        
        if num >= 20:
            result += tens[num // 10] + " "
            num %= 10
        elif num >= 10:
            result += teens[num - 10] + " "
            return result.strip()
        
        if num > 0:
            result += ones[num] + " "
        
        return result.strip()
    
    if n < 1000:
        return convert_hundreds(n)
    elif n < 1000000:
        thousands = n // 1000
        remainder = n % 1000
        result = convert_hundreds(thousands) + " THOUSAND"
        if remainder > 0:
            result += " " + convert_hundreds(remainder)
        return result
    else:
        return str(n)  # For very large numbers, just return the string representation