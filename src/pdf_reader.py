# ==========================================================
# src/pdf_reader.py
# FIXED: ROBUST PDF COLUMN / SECTION EXTRACTION
# ==========================================================

import os
import re
import logging
from typing import List, Dict, Any, Optional

import pdfplumber
import pandas as pd

from src.logger import info, warning, error
from src.exceptions import PDFReadError

logger = logging.getLogger(__name__)

if not logger.handlers:
    logging.basicConfig(
        level=logging.INFO,
        format="%(levelname)s : %(message)s"
    )

PDF_COLUMNS = [
    "Sr", "Width", "Height", "ChargeWidth", "ChargeHeight",
    "Charge Size Status", "Qty", "Remark", "WindowCode", "GlassType"
]


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\n", " ").replace("\r", " ")
    return re.sub(r"\s+", " ", text).strip()


def safe_int(value: Any) -> Optional[int]:
    if value is None:
        return None

    cleaned = re.sub(r"[^\d.]", "", str(value).strip())

    try:
        val_float = float(cleaned)
        return int(round(val_float)) if val_float > 0 else None
    except (ValueError, TypeError):
        return None


def integer_token(value: Any) -> Optional[int]:
    """Strict integer parser: decimal area/rate values cannot become Qty."""
    text = clean_text(value).replace(",", "")

    if not re.fullmatch(r"\d+", text):
        return None

    try:
        return int(text)
    except (ValueError, TypeError):
        return None


def clean_window_code(value: Any) -> str:
    if not value:
        return ""

    val = str(value).upper().strip()

    if "BLOCK" in val:
        parts = val.split("BLOCK", 1)
        after_block = parts[1].strip()
        after_block = re.sub(r"^[\.\-\s]+", "", after_block)
        after_block = after_block.replace("-", " ")
        return re.sub(r"\s+", " ", after_block).strip()

    val = val.replace("-", " ")
    return re.sub(r"\s+", " ", val).strip()


def extract_pdf_raw(pdf_path: str) -> pd.DataFrame:
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF not found : {pdf_path}")

    raw_rows = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_no, page in enumerate(pdf.pages, start=1):
                text = page.extract_text() or ""
                tables = page.extract_tables() or []

                for table in tables:
                    for row in table:
                        if row:
                            raw_rows.append({
                                "Page": page_no,
                                "PageText": text,
                                "RawRow": row
                            })

    except Exception as e:
        logger.error(f"Raw Extraction Failed: {e}")

    return pd.DataFrame(raw_rows)


def is_data_row(row: list) -> bool:
    if not isinstance(row, list):
        return False

    cells = [clean_text(x) for x in row if x is not None]
    valid_cells = [x for x in cells if x != ""]

    if len(valid_cells) < 4:
        return False

    sr = integer_token(valid_cells[0])

    if sr is None:
        return False

    dimension_values = [
        integer_token(x)
        for x in valid_cells[1:]
    ]

    dimension_values = [
        x for x in dimension_values
        if x is not None and x >= 100
    ]

    return len(dimension_values) >= 6


def extract_charge_size_from_string(
    cell_text: str
) -> Optional[tuple[int, int]]:

    match = re.search(
        r"(\d{3,4})\s*[\u00d7\*X]\s*(\d{3,4})",
        cell_text,
        re.IGNORECASE
    )

    if match:
        return int(match.group(1)), int(match.group(2))

    return None


def extract_structured_row_values(
    row: list
) -> Optional[Dict[str, Any]]:
    """
    PDF table layouts change column count between pages.
    Do NOT remove None cells before interpreting the row.

    Stable numeric order:
      Sr
      Actual W
      Actual H
      Actual W
      Actual H
      Charge W
      Charge H
      Qty
      Area
      Rate
      ...
      Remark

    The first six integer dimensions are therefore used for
    Actual/Charge sizes, and Qty is the first strict integer
    1..99 after Charge Height.
    """

    if not isinstance(row, list):
        return None

    raw_cells = [clean_text(x) for x in row]

    first_index = None
    sr_no = None

    for idx, cell in enumerate(raw_cells):
        if not cell:
            continue

        candidate = integer_token(cell)

        if candidate is not None:
            first_index = idx
            sr_no = candidate
            break

        return None

    if first_index is None or sr_no is None:
        return None

    dimension_values = []

    for cell in raw_cells[first_index + 1:]:
        value = integer_token(cell)

        if value is not None and value >= 100:
            dimension_values.append(value)

    if len(dimension_values) < 6:
        return None

    width = dimension_values[0]
    height = dimension_values[1]
    charge_width = dimension_values[4]
    charge_height = dimension_values[5]

    qty = None
    charge_seen = 0

    for cell in raw_cells[first_index + 1:]:
        value = integer_token(cell)

        if value is not None and value >= 100:
            charge_seen += 1

            if charge_seen >= 6:
                continue

        elif value is not None and 1 <= value < 100 and charge_seen >= 6:
            qty = value
            break

    if qty is None:
        qty = 1

    remark = ""

    for cell in reversed(raw_cells):
        if "BLOCK" in cell.upper():
            remark = cell
            break

    return {
        "Sr": sr_no,
        "Width": width,
        "Height": height,
        "ChargeWidth": charge_width,
        "ChargeHeight": charge_height,
        "Qty": qty,
        "Remark": remark
    }


GLASS_TYPE_MAPPINGS = {
    r"\b(DGU|INSULAT(?:ING|ED)|DOUBLE\s*GLAZED|IGU)\b": "DGU",
    r"\b(SGU|MONOLITHIC|SINGLE\s*GLAZED)\b": "SGU",
    r"\b(TRIPLE\s*GLAZED|TGU)\b": "TGU",
    r"\b(LAMINATED|LAM)\b": "LAMINATED",
    r"\b(TOUGHENED|TEMPERED)\b": "TOUGHENED",
    r"\b(FROSTED|SATIN|OPAQUE)\b": "FROSTED",
    r"\b(LOW\s*E|LOWE)\b": "LOW-E",
}


# ============================================================
# DYNAMIC GLASS SPEC DETECTION ENGINE
# ============================================================

def extract_dynamic_glass_spec(
    full_line: str,
    current_glass_type: str
) -> str:

    line_upper = str(full_line or "").upper().strip()
    curr_upper = str(current_glass_type or "").upper().strip()

    dgu_patterns = [
        r'\bDGU\b',
        r'\bINSULAT(?:ING|ED)\b',
        r'\bIGU\b',
        r'\bDOUBLE\s*GLAZED\b'
    ]

    is_dgu = any(
        re.search(
            pattern,
            line_upper,
            re.IGNORECASE
        )
        for pattern in dgu_patterns
    )

    current_is_dgu = any(
        re.search(
            pattern,
            curr_upper,
            re.IGNORECASE
        )
        for pattern in dgu_patterns
    )

    # ========================================================
    # DGU
    # ========================================================

    if is_dgu or current_is_dgu:

        # ----------------------------------------------------
        # 1. Current section already knows DGU thickness
        # ----------------------------------------------------

        current_dgu_match = re.search(
            r'(\d+(?:\.\d+)?)\s*MM\s*DGU\b',
            curr_upper,
            re.IGNORECASE
        )

        if current_dgu_match:

            mm_value = float(
                current_dgu_match.group(1)
            )

            if mm_value.is_integer():
                mm_text = str(int(mm_value))
            else:
                mm_text = str(mm_value)

            return f"{mm_text} MM DGU"

        # ----------------------------------------------------
        # 2. Explicit DGU thickness in current row
        # ----------------------------------------------------

        dgu_match = re.search(
            r'(\d+(?:\.\d+)?)\s*MM\s*'
            r'(?:DGU|IGU|INSULATING\s+GLASS|'
            r'DOUBLE\s+GLAZED)',
            line_upper,
            re.IGNORECASE
        )

        if dgu_match:

            mm_value = float(
                dgu_match.group(1)
            )

            if mm_value.is_integer():
                mm_text = str(int(mm_value))
            else:
                mm_text = str(mm_value)

            return f"{mm_text} MM DGU"

        # ----------------------------------------------------
        # 3. Reverse format:
        #
        # DGU 20 MM
        # IGU 24 MM
        # ----------------------------------------------------

        reverse_dgu_match = re.search(
            r'(?:DGU|IGU|INSULATING\s+GLASS|'
            r'DOUBLE\s+GLAZED)'
            r'.{0,20}?'
            r'(\d+(?:\.\d+)?)\s*MM',
            line_upper,
            re.IGNORECASE
        )

        if reverse_dgu_match:

            mm_value = float(
                reverse_dgu_match.group(1)
            )

            if mm_value.is_integer():
                mm_text = str(int(mm_value))
            else:
                mm_text = str(mm_value)

            return f"{mm_text} MM DGU"

        # ----------------------------------------------------
        # 4. Thickness unavailable
        #
        # NEVER assume 20 MM
        # ----------------------------------------------------

        return "DGU - THICKNESS NOT SPECIFIED"

    # ========================================================
    # NORMAL SINGLE GLASS
    # ========================================================

    line_mm_match = re.search(
        r'(\d+(?:\.\d+)?\s*MM)',
        line_upper
    )

    if line_mm_match:

        mm_str = re.sub(
            r'(\d+)\s*MM',
            r'\1 MM',
            line_mm_match.group(1)
        )

        return (
            f"{mm_str} "
            "CLEAR GLASS SAINT GOBAIN WITH RG"
        )

    # ========================================================
    # CURRENT SECTION FALLBACK
    # ========================================================

    if curr_upper:
        return current_glass_type

    # ========================================================
    # FINAL FALLBACK
    # ========================================================

    return "5MM CLEAR GLASS SAINT GOBAIN WITH RG"
    
def extract_glass_header(row_text: str) -> Optional[str]:
    """
    Supports both:
      TOUGHENED GLASS.: [...]
    and pdfplumber's split:
      TOUGHENED G LASS.: [...]
    """

    match = re.search(
        r"TOUGHENED\s+G?\s*LASS\s*\.\s*:\s*\[\s*([^\]]+?)\s*\]",
        row_text,
        re.IGNORECASE
    )

    if match:
        return clean_text(match.group(1))

    return None


def parse_glass_row(
    row: list,
    current_glass_type: str,
    default_window_code: str
) -> Optional[Dict[str, Any]]:

    if not isinstance(row, list):
        return None

    structured = extract_structured_row_values(row)

    if not structured:
        return None

    remark = structured["Remark"]

    window_code = (
        clean_window_code(remark)
        if remark
        else default_window_code
    )

    row_str = " ".join(
        clean_text(x) for x in row if x is not None
    ).upper()

    actual_glass_type = extract_dynamic_glass_spec(
        row_str,
        current_glass_type
    )

    charge_status = (
        "OK"
        if (
            structured["ChargeWidth"] == structured["Width"] + 20
            and
            structured["ChargeHeight"] == structured["Height"] + 20
        )
        else "CHARGE SIZE INCORRECT"
    )

    return {
        "Sr": structured["Sr"],
        "Width": structured["Width"],
        "Height": structured["Height"],
        "ChargeWidth": structured["ChargeWidth"],
        "ChargeHeight": structured["ChargeHeight"],
        "Charge Size Status": charge_status,
        "Qty": structured["Qty"],
        "Remark": remark,
        "WindowCode": window_code,
        "GlassType": actual_glass_type
    }


def parse_pdf_records(raw_df: pd.DataFrame) -> pd.DataFrame:

    records = []

    current_glass_type = ""
    default_window_code = ""

    if raw_df.empty:
        return pd.DataFrame(columns=PDF_COLUMNS)

    for _, item in raw_df.iterrows():

        row = item.get("RawRow", None)
        page_text = item.get("PageText", "")

        if not isinstance(row, list):
            continue

        row_text = " ".join(
            clean_text(x)
            for x in row
            if x is not None
        )

        row_upper = row_text.upper()

        # =====================================================
        # GLASS SECTION HEADER
        #
        # IMPORTANT:
        # ONLY ROW TEXT.
        #
        # DO NOT use PageText here.
        # Otherwise one DGU header on a page can make
        # every row on that page DGU.
        # =====================================================

        dgu_header_match = re.search(
            r'(\d+(?:\.\d+)?)\s*MM\s+'
            r'(INSULATING\s+GLASS|DGU|IGU|DOUBLE\s+GLAZED)',
            row_upper,
            re.IGNORECASE
        )

        if dgu_header_match:

            mm_value = float(
                dgu_header_match.group(1)
            )

            if mm_value.is_integer():
                mm_text = str(int(mm_value))
            else:
                mm_text = str(mm_value)

            current_glass_type = (
                f"{mm_text} MM DGU"
            )

        else:

            # -------------------------------------------------
            # Normal glass section header
            # -------------------------------------------------

            glass_header = extract_glass_header(row_text)

            if glass_header:
                current_glass_type = glass_header

        # =====================================================
        # WINDOW CODE FROM PAGE TEXT
        #
        # PageText is OK here.
        # It must NOT be used for glass section detection.
        # =====================================================

        gh_match = re.search(
            r"\b(GH\d+|MARK\s*\w+|ITEM\s*\w+)\b",
            page_text,
            re.IGNORECASE
        )

        if gh_match:
            default_window_code = (
                gh_match.group(1).upper()
            )

        # =====================================================
        # IGNORE NON-DATA ROWS
        # =====================================================

        if not is_data_row(row):
            continue

        # =====================================================
        # IGNORE TABLE HEADER / TOTAL ROWS
        # =====================================================

        if any(
            k in row_upper
            for k in [
                "WIDTH",
                "HEIGHT",
                "HSN #",
                "CS:",
                "AMOUNT INR"
            ]
        ):
            continue

        # =====================================================
        # PARSE ACTUAL GLASS ROW
        # =====================================================

        record = parse_glass_row(
            row,
            current_glass_type,
            default_window_code
        )

        if record:
            records.append(record)

    if not records:
        return pd.DataFrame(columns=PDF_COLUMNS)

    return pd.DataFrame(
        records,
        columns=PDF_COLUMNS
    )


def process_pdf_dataframe(
    pdf_df: pd.DataFrame
) -> pd.DataFrame:

    if pdf_df.empty:
        return pdf_df

    pdf_df = pdf_df.drop_duplicates().reset_index(drop=True)

    return pdf_df


def read_pdf(pdf_path: str) -> pd.DataFrame:

    try:
        raw_df = extract_pdf_raw(pdf_path)

        if raw_df.empty:
            return pd.DataFrame(columns=PDF_COLUMNS)

        pdf_df = parse_pdf_records(raw_df)
        pdf_df = process_pdf_dataframe(pdf_df)

        return pdf_df

    except Exception as e:
        logger.exception(f"PDF Reader Failed : {e}")
        raise PDFReadError(
            f"Error parsing PDF: {str(e)}"
        )