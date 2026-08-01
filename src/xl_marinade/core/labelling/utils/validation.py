# ABOUTME: Validation utilities for documentation agent components
# This module provides shared validation logic for labels, variables, and other entities.

import re


def is_valid_label_candidate(text: str) -> bool:
    """
    Check if a text string is a valid candidate for a variable label.

    Rejects:
    - Strings starting with '=' (formulas)
    - Strings that look like cell ranges (e.g. "A1:B10", "$A$1")
    - Strings containing math operators that suggest a formula snippet (e.g. "A+B")
    - Empty strings

    Args:
        text: The text to validate

    Returns:
        True if valid, False otherwise
    """
    # H6: defensive — non-str inputs (e.g. a raw numeric JSON literal) are not labels.
    if not isinstance(text, str):
        return False

    if not text:
        return False

    text = text.strip()
    if not text:
        return False

    # Reject formulas
    if text.startswith("="):
        return False

    # Reject labels containing formula patterns (e.g. "Label =SUM(...)")
    # Matches " =SUM(", " =AVERAGE(", etc.
    if re.search(r"\s=[A-Z]+\(", text):
        return False

    # Reject labels that end with a formula-like structure
    # e.g. "discounted profit cashflow =SUM(U7:U607)"
    if re.search(r"=[A-Z]+\([A-Z0-9:]+\)$", text):
        return False

    # Reject common error values
    if text.startswith("#"):
        return False

    # H6: reject footnote / schedule markers (e.g. "(2)", "(13)") — these are
    # legal-schedule cross-references that sit directly above data columns, not names.
    if re.fullmatch(r"\(\d+\)", text):
        return False

    # H6: reject GUIDs. Issue #3: use search (not fullmatch) so a GUID with a
    # suffix/prefix (e.g. "<guid>_5", a split sub-binding's leaked id) is caught
    # too — no legitimate label embeds a full GUID.
    if re.search(
        r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}", text
    ):
        return False

    # Issue #3: reject arithmetic formula-tokens like "(1) * (4)" or
    # "(1) * (3) + (2) * (4)" — parenthesised column-index refs joined by
    # operators. These carry no letters and are a formula's structure, not a
    # name. Gated on "(" + an operator so legit parenthetical headers
    # ("(draft)", "('000s)", "(A)") and bare footnotes ("(2)", handled above)
    # are untouched.
    if not any(ch.isalpha() for ch in text) and "(" in text and re.search(r"[+\-*/]", text):
        return False

    # Issue #3: reject computed-VALUE numerics that are decidably not headers —
    # scientific notation (8.48E-2) and long-precision decimals
    # (-20.995834250000001, >=6 fractional digits). Both are computed cell
    # values, never headers. (This refines H6's earlier "stays permissive" stance
    # for these two unambiguous value forms; SHORT numeric headers — years like
    # 2024, integers like 42, few-dp ratios like 1.5 — remain valid.)
    if re.fullmatch(r"[+-]?\d+(?:\.\d+)?[eE][+-]?\d+", text):
        return False
    _dec = re.fullmatch(r"[+-]?\d+\.(\d+)", text)
    if _dec and len(_dec.group(1)) >= 6:
        return False

    # H6: require at least one alphanumeric char — rejects pure punctuation/symbol
    # noise ("()", "@@", "+-*/") while KEEPING numeric headers (years/quarters like
    # "2024") and short domain codes ("Q1", "FY24", "PV01", "S1"). Distinguishing a
    # value-as-header from a real numeric/coded header is not decidable here, so the
    # validator stays permissive; preference for a *descriptive* sibling header is
    # handled by the id-code/junk skip-walk in simple_labeller.select_best_literal.
    if not any(ch.isalnum() for ch in text):
        return False

    # Reject pure ranges (e.g. A1:B10, $A$1:$B$10)
    # Case insensitive check
    t_upper = text.upper()

    # Range pattern: optional $ then letters then optional $ then digits, colon, repeat
    range_pattern = r"^\$?[A-Z]+\$?[0-9]+:\$?[A-Z]+\$?[0-9]+$"
    if re.match(range_pattern, t_upper):
        return False

    # Single cell address with $ (e.g. $A$1, A$1, $A1) - definitely a reference
    if "$" in text:
        # If it looks like a reference with $
        # Check for basic cell reference structure
        ref_pattern = r"\$?[A-Z]+\$?[0-9]+"
        if re.search(ref_pattern, t_upper):
            return False

    # Sheet qualified reference (Sheet!A1, 'Sheet'!A1)
    # Heuristic: ends with a cell reference or range
    # Matches: ...!A1 or ...!A1:B2
    return not ("!" in text and re.search(r"![^!]*[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$", t_upper))
