import re
from typing import Any, Dict, List, Optional, Set, Tuple
import pandas as pd


def normalize_text(val: Any) -> str:
    """Clean and normalize text for uniform matching."""
    if pd.isna(val) or val is None:
        return ""
    text = str(val).upper().strip()
    return re.sub(r"[^A-Z0-9]", "", text)


def standardize_glass_spec(val: Any) -> str:
    """Standardize glass specification string while preserving original text structure."""
    if pd.isna(val) or val is None:
        return "NOT SPECIFIED"

    text = str(val).strip()

    if not text or text.lower() == "nan":
        return "NOT SPECIFIED"

    return re.sub(r"\s+", " ", text).strip()


def extract_glass_thickness(spec_str: Any) -> Optional[float]:
    """Extract glass thickness as a numerical float
    (e.g. '6.0 MM' -> 6.0, '12MM' -> 12.0, '5MM' -> 5.0).
    """

    if pd.isna(spec_str) or not str(spec_str).strip():
        return None

    text = str(spec_str).upper().strip()

    # Matches numbers followed by optional space and MM, M.M., or MM.
    match = re.search(
        r"(\d+(?:\.\d+)?)\s*(?:MM|M\.M\.?|MM\.)",
        text
    )

    if match:
        try:
            return float(match.group(1))
        except ValueError:
            return None

    # Fallback to standalone numbers in realistic glass thickness range
    # (2mm to 35mm)
    match_fallback = re.search(
        r"\b(\d+(?:\.\d+)?)\b",
        text
    )

    if match_fallback:
        try:
            val = float(match_fallback.group(1))

            if 2.0 <= val <= 35.0:
                return val

        except ValueError:
            return None

    return None


# ==========================================================
# GLASS SPECIFICATION MATCH
# ==========================================================

def check_glass_spec_match(
    excel_glass: str,
    pdf_glass: str
) -> bool:

    if not excel_glass or not pdf_glass:
        return True

    ex = str(excel_glass).upper().strip()
    pdf = str(pdf_glass).upper().strip()

    if pdf in ["-", "", "NONE", "NOT SPECIFIED", "NAN"]:
        return True

    # =====================================================
    # 1. THICKNESS CHECK
    # =====================================================

    ex_mm = re.search(
        r'(\d+(?:\.\d+)?)\s*MM',
        ex
    )

    pdf_mm = re.search(
        r'(\d+(?:\.\d+)?)\s*MM',
        pdf
    )

    if ex_mm and pdf_mm:

        if float(ex_mm.group(1)) != float(pdf_mm.group(1)):
            return False

    # =====================================================
    # 2. SPECIAL RULE
    #
    # Excel:
    #   5 MM REFLECTIVE TOUGHENED
    #
    # PDF:
    #   5MM REF IMPERIAL BLUE ST 720 SAINT GOBAIN WITH RG
    #
    # IMPORTANT:
    # PDF colour can be ANYTHING.
    #
    # We ignore:
    #   IMPERIAL
    #   BLUE / GREEN / GREY / ANY COLOR
    #   ST 720
    #   SAINT GOBAIN
    #   WITH RG
    #
    # Only these keywords are required:
    #
    # Excel:
    #   5MM + REFLECTIVE + TOUGHENED
    #
    # PDF:
    #   5MM + REF
    #
    # =====================================================

    ex_compact = re.sub(
        r"[^A-Z0-9]",
        "",
        ex
    )

    pdf_compact = re.sub(
        r"[^A-Z0-9]",
        "",
        pdf
    )

    excel_reflective_5mm = (
        "5MM" in ex_compact
        and
        "REFLECTIVE" in ex_compact
        and
        "TOUGHENED" in ex_compact
    )

    pdf_reflective_5mm = (
        "5MM" in pdf_compact
        and
        "REF" in pdf_compact
    )

    if excel_reflective_5mm and pdf_reflective_5mm:
        return True

    # =====================================================
    # 3. NORMAL SPECIAL GLASS TYPE CHECK
    #
    # Existing logic preserved.
    # =====================================================

    keywords = [
        "REFLECTIVE",
        "EXTRA CLEAR",
        "LOW IRON",
        "LOW-IRON",
        "OPTIWHITE",
        "FROSTED",
        "LAMINATED",
        "DGU",
        "TINTED"
    ]

    for kw in keywords:

        if (kw in ex) != (kw in pdf):
            return False

    return True


def compare_glass_specs(
    excel_glass: Any,
    pdf_glass: Any
) -> Tuple[bool, str]:

    ex_str = str(excel_glass or "").strip()
    pdf_str = str(pdf_glass or "").strip()

    is_matched = check_glass_spec_match(
        ex_str,
        pdf_str
    )

    if is_matched:
        return True, "Glass spec matched"

    else:
        return (
            False,
            f"Glass Spec Mismatch "
            f"(Excel: '{ex_str}' | PDF: '{pdf_str}')"
        )


# ==========================================================
# MULTI-PASS BEST-FIT VERIFICATION ENGINE
# ==========================================================

def verify_pi_against_excel(
    excel_df: pd.DataFrame,
    pdf_df: pd.DataFrame,
    tolerance_mm: int = 2
) -> pd.DataFrame:

    """
    Smart Multi-Pass Engine:

    Matches rows by Exact Window Code and Dimensions
    regardless of row order!
    """

    if excel_df is None or excel_df.empty:
        return pd.DataFrame()

    excel_df = excel_df.copy().reset_index(drop=True)

    pdf_df = (
        pdf_df.copy().reset_index(drop=True)
        if pdf_df is not None and not pdf_df.empty
        else pd.DataFrame()
    )

    # =====================================================
    # TRACK MATCHES
    # =====================================================

    matched_pdf_for_excel = {
        i: None
        for i in range(len(excel_df))
    }

    used_pdf_indices: Set[int] = set()

    # =====================================================
    # HELPER FUNCTION
    # =====================================================

    def parse_row(
        df,
        idx,
        is_pdf=False
    ):

        row = df.iloc[idx]

        code_col = (
            "WindowCode"
            if "WindowCode" in df.columns
            else (
                "Window Code (PDF)"
                if is_pdf
                else "Window Code (Excel)"
            )
        )

        w_col = (
            "Width"
            if "Width" in df.columns
            else "WIDTH"
        )

        h_col = (
            "Height"
            if "Height" in df.columns
            else "HEIGHT"
        )

        qty_col = (
            "Qty"
            if "Qty" in df.columns
            else (
                "QTY"
                if "QTY" in df.columns
                else (
                    "PDF Qty"
                    if is_pdf
                    else "Excel Qty"
                )
            )
        )

        glass_col = (
            "GlassType"
            if "GlassType" in df.columns
            else (
                "Glass Spec (PDF)"
                if is_pdf
                else "Glass Spec (Excel)"
            )
        )

        code = str(
            row.get(code_col, "")
        ).strip()

        norm_code = normalize_text(code)

        try:

            w = (
                int(float(row.get(w_col, 0)))
                if pd.notna(row.get(w_col))
                else 0
            )

        except ValueError:

            w = 0

        try:

            h = (
                int(float(row.get(h_col, 0)))
                if pd.notna(row.get(h_col))
                else 0
            )

        except ValueError:

            h = 0

        try:

            qty = (
                int(float(row.get(qty_col, 1)))
                if pd.notna(row.get(qty_col))
                else 1
            )

        except ValueError:

            qty = 1

        glass = str(
            row.get(glass_col, "")
        ).strip()

        return (
            norm_code,
            code,
            w,
            h,
            qty,
            glass
        )

    # =====================================================
    # PARSE ALL ROWS UPFRONT
    # =====================================================

    ex_parsed = [
        parse_row(
            excel_df,
            i,
            False
        )
        for i in range(len(excel_df))
    ]

    pdf_parsed = (
        [
            parse_row(
                pdf_df,
                j,
                True
            )
            for j in range(len(pdf_df))
        ]
        if not pdf_df.empty
        else []
    )

    # =====================================================
    # PASS 1
    # Perfect Match
    # Code + Size + Qty + Glass Spec
    # =====================================================

    for i, ex in enumerate(ex_parsed):

        if matched_pdf_for_excel[i] is not None:
            continue

        (
            ex_norm,
            _,
            ex_w,
            ex_h,
            ex_qty,
            ex_glass
        ) = ex

        for j, pdf in enumerate(pdf_parsed):

            if j in used_pdf_indices:
                continue

            (
                pdf_norm,
                _,
                pdf_w,
                pdf_h,
                pdf_qty,
                pdf_glass
            ) = pdf

            if ex_norm == pdf_norm:

                dim_ok = (
                    abs(ex_w - pdf_w) <= tolerance_mm
                    and
                    abs(ex_h - pdf_h) <= tolerance_mm
                )

                qty_ok = (
                    ex_qty == pdf_qty
                )

                glass_ok, _ = compare_glass_specs(
                    ex_glass,
                    pdf_glass
                )

                if (
                    dim_ok
                    and qty_ok
                    and glass_ok
                ):

                    matched_pdf_for_excel[i] = j
                    used_pdf_indices.add(j)

                    break

    # =====================================================
    # PASS 2
    # Size & Glass Match
    # Code + Size + Glass Spec
    # =====================================================

    for i, ex in enumerate(ex_parsed):

        if matched_pdf_for_excel[i] is not None:
            continue

        (
            ex_norm,
            _,
            ex_w,
            ex_h,
            _,
            ex_glass
        ) = ex

        for j, pdf in enumerate(pdf_parsed):

            if j in used_pdf_indices:
                continue

            (
                pdf_norm,
                _,
                pdf_w,
                pdf_h,
                _,
                pdf_glass
            ) = pdf

            if ex_norm == pdf_norm:

                dim_ok = (
                    abs(ex_w - pdf_w) <= tolerance_mm
                    and
                    abs(ex_h - pdf_h) <= tolerance_mm
                )

                glass_ok, _ = compare_glass_specs(
                    ex_glass,
                    pdf_glass
                )

                if (
                    dim_ok
                    and glass_ok
                ):

                    matched_pdf_for_excel[i] = j
                    used_pdf_indices.add(j)

                    break

    # =====================================================
    # PASS 3
    # Dimension Match
    # Code + Size Match
    # =====================================================

    for i, ex in enumerate(ex_parsed):

        if matched_pdf_for_excel[i] is not None:
            continue

        (
            ex_norm,
            _,
            ex_w,
            ex_h,
            _,
            _
        ) = ex

        for j, pdf in enumerate(pdf_parsed):

            if j in used_pdf_indices:
                continue

            (
                pdf_norm,
                _,
                pdf_w,
                pdf_h,
                _,
                _
            ) = pdf

            if ex_norm == pdf_norm:

                dim_ok = (
                    abs(ex_w - pdf_w) <= tolerance_mm
                    and
                    abs(ex_h - pdf_h) <= tolerance_mm
                )

                if dim_ok:

                    matched_pdf_for_excel[i] = j
                    used_pdf_indices.add(j)

                    break

    # =====================================================
    # PASS 4
    # Rotated Size Match
    # Width <-> Height flipped
    # =====================================================

    for i, ex in enumerate(ex_parsed):

        if matched_pdf_for_excel[i] is not None:
            continue

        (
            ex_norm,
            _,
            ex_w,
            ex_h,
            _,
            _
        ) = ex

        for j, pdf in enumerate(pdf_parsed):

            if j in used_pdf_indices:
                continue

            (
                pdf_norm,
                _,
                pdf_w,
                pdf_h,
                _,
                _
            ) = pdf

            if ex_norm == pdf_norm:

                rotated_dim_ok = (
                    abs(ex_w - pdf_h) <= tolerance_mm
                    and
                    abs(ex_h - pdf_w) <= tolerance_mm
                )

                if rotated_dim_ok:

                    matched_pdf_for_excel[i] = j
                    used_pdf_indices.add(j)

                    break

    # =====================================================
    # PASS 5
    # Window Code Match
    # Fallback for dimension mismatch
    # =====================================================

    for i, ex in enumerate(ex_parsed):

        if matched_pdf_for_excel[i] is not None:
            continue

        ex_norm = ex[0]

        for j, pdf in enumerate(pdf_parsed):

            if j in used_pdf_indices:
                continue

            pdf_norm = pdf[0]

            if ex_norm == pdf_norm:

                matched_pdf_for_excel[i] = j
                used_pdf_indices.add(j)

                break

    # =====================================================
    # CONSTRUCT VERIFICATION REPORT
    # =====================================================

    results = []

    for i, ex in enumerate(ex_parsed):

        (
            _,
            ex_code,
            ex_w,
            ex_h,
            ex_qty,
            ex_glass
        ) = ex

        j = matched_pdf_for_excel[i]

        # -------------------------------------------------
        # NOT FOUND
        # -------------------------------------------------

        if j is None:

            status = "❌ NOT FOUND IN PDF"

            remark = (
                "Window Code not present in PDF"
            )

            (
                pdf_code,
                pdf_w,
                pdf_h,
                pdf_qty,
                pdf_glass
            ) = (
                "-",
                "-",
                "-",
                "-",
                "-"
            )

        # -------------------------------------------------
        # FOUND
        # -------------------------------------------------

        else:

            (
                _,
                pdf_code,
                pdf_w,
                pdf_h,
                pdf_qty,
                pdf_glass
            ) = pdf_parsed[j]

            # ---------------------------------------------
            # DIMENSION CHECK
            # ---------------------------------------------

            dim_match = (
                (
                    abs(ex_w - pdf_w)
                    <= tolerance_mm
                    and
                    abs(ex_h - pdf_h)
                    <= tolerance_mm
                )
                or
                (
                    abs(ex_w - pdf_h)
                    <= tolerance_mm
                    and
                    abs(ex_h - pdf_w)
                    <= tolerance_mm
                )
            )

            # ---------------------------------------------
            # QTY CHECK
            # ---------------------------------------------

            qty_match = (
                ex_qty == pdf_qty
            )

            # ---------------------------------------------
            # GLASS CHECK
            # ---------------------------------------------

            glass_match, glass_remark = (
                compare_glass_specs(
                    ex_glass,
                    pdf_glass
                )
            )

            # ---------------------------------------------
            # FINAL STATUS
            # ---------------------------------------------

            if (
                dim_match
                and qty_match
                and glass_match
            ):

                status = "✅ MATCHED"

                remark = (
                    "All parameters matched"
                )

            elif not dim_match:

                status = (
                    "⚠️ DIMENSION MISMATCH"
                )

                remark = (
                    f"Excel: {ex_w}x{ex_h} | "
                    f"PDF: {pdf_w}x{pdf_h}"
                )

            elif not qty_match:

                status = (
                    "⚠️ QTY MISMATCH"
                )

                remark = (
                    f"Excel Qty: {ex_qty} | "
                    f"PDF Qty: {pdf_qty}"
                )

            elif not glass_match:

                status = (
                    "⚠️ GLASS SPEC MISMATCH"
                )

                remark = glass_remark

        # -------------------------------------------------
        # RESULT ROW
        # -------------------------------------------------

        results.append(
            {
                "Window Code (Excel)": ex_code,
                "Window Code (PDF)": pdf_code,
                "Excel Size": f"{ex_w} x {ex_h}",
                "PDF Size": (
                    f"{pdf_w} x {pdf_h}"
                    if pdf_w != "-"
                    else "-"
                ),
                "Excel Qty": ex_qty,
                "PDF Qty": pdf_qty,
                "Glass Spec (Excel)": ex_glass,
                "Glass Spec (PDF)": pdf_glass,
                "Verification Status": status,
                "Remarks": remark
            }
        )

    return pd.DataFrame(results)