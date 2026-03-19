"""
Extracts budget variables from a filled RV Form (.docx).
Returns a dict of values to fill into the budget template.
"""
import re
import subprocess
import tempfile
import os


def extract_rv_form(docx_path: str) -> dict:
    text = _get_text(docx_path)
    return {
        "client_name":        _org(text),
        "currency":           _currency(text),
        "study_type":         _study_type(text),
        "components":         _components(text),
        "sample_size":        _sample_size(text),
        "states":             _states(text),
        "num_blocks":         _num_blocks(text),
        "survey_duration":    _survey_duration(text),
        "num_fgds":           _qual_count(text, "FGD"),
        "num_idis":           _qual_count(text, "IDI"),
        "languages":          _languages(text),
        "oi_codes":           _checked(text, "Coding of survey tool"),
        "oi_devices":         _checked(text, "Survey Devices"),
        "revisits":           _revisits(text),
        "num_revisits":       _num_revisits(text),
        "back_check":         _back_check(text),
        "timeline_months":    _timeline(text),
        "budget_constraint":  _budget_constraint(text),
        "dc_mode":            _dc_mode(text),
    }


def _get_text(docx_path):
    with tempfile.NamedTemporaryFile(suffix=".md", delete=False) as f:
        tmp = f.name
    try:
        subprocess.run(["pandoc", "--track-changes=all", docx_path, "-o", tmp],
                       capture_output=True, check=True)
        with open(tmp, encoding="utf-8") as f:
            return f.read()
    finally:
        os.unlink(tmp)


def _org(text):
    m = re.search(r"Name of the organization[^:]*:\*?\s*(.+)", text, re.IGNORECASE)
    if m:
        v = re.sub(r"^[\s>\\*]+", "", m.group(1)).strip()
        return v[:80]
    return ""


def _currency(text):
    # Q28 mentions currency
    q28 = re.search(r"28\..{0,500}", text, re.DOTALL | re.IGNORECASE)
    snippet = q28.group(0) if q28 else text[:500]
    if re.search(r"\bUSD\b|\$|US dollar", snippet, re.IGNORECASE):
        return "USD"
    if re.search(r"\bGBP\b|£|sterling|pounds?", snippet, re.IGNORECASE):
        return "GBP"
    return "INR"


def _study_type(text):
    # Bold/underlined items indicate selection
    if re.search(r"\*\*\[?Mixed", text, re.IGNORECASE):
        return "mixed"
    if re.search(r"\*\*\[?Qualitative", text, re.IGNORECASE):
        return "qualitative"
    # fallback: if FGDs/IDIs mentioned
    has_q = bool(re.search(r"quantitative survey|household survey", text, re.IGNORECASE))
    has_ql = bool(re.search(r"\bFGD\b|\bIDI\b|\bKII\b|in-depth interview|focus group", text, re.IGNORECASE))
    if has_q and has_ql:
        return "mixed"
    if has_ql:
        return "qualitative"
    return "quantitative"


def _components(text):
    found = []
    mapping = {
        "Data Collection": r"\*\*\[?Data collection",
        "Analysis": r"\*\*\[?Analysis",
        "Report Writing": r"\*\*\[?Report Writing",
        "Translation": r"\*\*\[?Translation of Study",
        "Pretest": r"\*\*\[?Pretest",
        "Field Training": r"\*\*\[?Field training",
        "Study Tools": r"\*\*\[?Study Tools",
    }
    for name, pat in mapping.items():
        if re.search(pat, text, re.IGNORECASE):
            found.append(name)
    return found or ["Data Collection"]


def _sample_size(text):
    # Try "Ideally N households" or "N households" near Q12
    patterns = [
        r"Ideally\s+(\d[\d,]*)\s+household",
        r"(\d[\d,]*)\s+household",
        r"sample size.{0,50}?(\d[\d,]+)",
        r"(\d+)\s*[-–to]+\s*(\d+)\s*household",  # range → take upper
    ]
    for p in patterns:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            try:
                # If range pattern, group(2) is upper bound
                if "[-–to]+" in p and m.lastindex == 2:
                    return int(m.group(2).replace(",", ""))
                return int(m.group(1).replace(",", ""))
            except (ValueError, IndexError):
                pass
    return 500


def _states(text):
    state_list = [
        "Uttarakhand", "Uttar Pradesh", "Bihar", "Rajasthan", "Maharashtra",
        "Karnataka", "West Bengal", "Tamil Nadu", "Gujarat", "Delhi",
        "Madhya Pradesh", "Andhra Pradesh", "Kerala", "Odisha", "Jharkhand",
        "Chhattisgarh", "Himachal Pradesh", "Punjab", "Haryana", "Assam",
        "Telangana", "Goa",
    ]
    found = [s for s in state_list if re.search(rf"\b{s}\b", text, re.IGNORECASE)]
    return found if found else ["State 1"]


def _num_blocks(text):
    # Q14 table or mentions of blocks
    m = re.search(r"Block#?\s*[|:]?\s*(\d+)", text, re.IGNORECASE)
    if m:
        return int(m.group(1))
    m = re.search(r"(\d+)\s+blocks?", text, re.IGNORECASE)
    if m:
        v = int(m.group(1))
        if 1 <= v <= 50:
            return v
    return 1


def _survey_duration(text):
    if re.search(r"\*\*Less than 30", text, re.IGNORECASE):
        return 20
    if re.search(r"\*\*Between 30.60", text, re.IGNORECASE):
        return 45
    if re.search(r"\*\*More than 60", text, re.IGNORECASE):
        return 75
    m = re.search(r"(\d+)\s*[-–]\s*(\d+)\s*min", text, re.IGNORECASE)
    if m:
        return (int(m.group(1)) + int(m.group(2))) // 2
    m = re.search(r"~\s*(\d+)\s*min", text, re.IGNORECASE)
    if m:
        return int(m.group(1))
    return 45


def _qual_count(text, qual_type):
    best = 0
    patterns = [
        rf"No\.\s*of\s*{qual_type}\w*\s*[-:]+\s*(\d+)",
        rf"(\d+)\s*{qual_type}s?\b",
        rf"{qual_type}s?\s*[-:]\s*(\d+)",
    ]
    for p in patterns:
        for m in re.finditer(p, text, re.IGNORECASE):
            try:
                v = int(m.group(1))
                if 0 < v <= 50 and v > best:
                    best = v
            except (ValueError, IndexError):
                pass
    return best


def _languages(text):
    lang_list = ["Hindi", "Marathi", "Bengali", "Tamil", "Telugu",
                 "Kannada", "Malayalam", "Gujarati", "Punjabi", "Odia"]
    # Only count languages mentioned in translation/tool context
    ctx = re.findall(r"(?:translat|language|tool|material).{0,300}", text, re.IGNORECASE | re.DOTALL)
    search = " ".join(ctx) if ctx else ""
    return [l for l in lang_list if re.search(rf"\b{l}\b", search, re.IGNORECASE)]


def _checked(text, label):
    # Look for bold/underlined formatting next to the checkbox item
    m = re.search(rf"\*\*\[?{re.escape(label)}", text, re.IGNORECASE)
    return bool(m)


def _revisits(text):
    q21 = re.search(r"21\..{0,300}", text, re.DOTALL | re.IGNORECASE)
    if q21:
        return bool(re.search(r"\*\*Yes\*\*", q21.group(0), re.IGNORECASE))
    return False


def _num_revisits(text):
    q22 = re.search(r"22\..{0,200}", text, re.DOTALL | re.IGNORECASE)
    if q22:
        s = q22.group(0)
        if re.search(r"\*\*1\*\*", s): return 1
        if re.search(r"\*\*2\*\*", s): return 2
        if re.search(r"\*\*3\*\*", s): return 3
        if re.search(r"More than 3", s, re.IGNORECASE): return 4
    return 0


def _back_check(text):
    q23 = re.search(r"23\..{0,200}", text, re.DOTALL | re.IGNORECASE)
    if q23:
        return bool(re.search(r"\*\*Yes\*\*", q23.group(0)))
    return False


def _timeline(text):
    m = re.search(r"(\d+)\s*months?", text, re.IGNORECASE)
    if m:
        v = int(m.group(1))
        if 1 <= v <= 48:
            return v
    return 3


def _budget_constraint(text):
    m = re.search(r"(?:INR|Rs\.?|₹)?\s*([\d,]+(?:\.\d+)?)\s*(?:lakh|crore)?",
                  text[text.lower().find("28."):text.lower().find("28.")+300]
                  if "28." in text.lower() else "", re.IGNORECASE)
    if m:
        v = float(m.group(1).replace(",", ""))
        ctx = text[text.lower().find("28."):text.lower().find("28.")+300].lower()
        if "lakh" in ctx:
            v *= 100000
        elif "crore" in ctx:
            v *= 10000000
        return v
    return 0.0


def _dc_mode(text):
    if re.search(r"\*\*Field/in.?person\*\*", text, re.IGNORECASE):
        return "field"
    if re.search(r"\*\*Telephonic\*\*", text, re.IGNORECASE):
        return "telephonic"
    if re.search(r"\*\*Online\*\*", text, re.IGNORECASE):
        return "online"
    return "field"


def extract_query_doc(docx_path: str) -> dict:
    """
    Extract clarification data from a Query Document.
    Returns a partial schema dict — only keys that are found.
    """
    text = _get_text(docx_path)
    result = {}

    fgds = _qual_count(text, "FGD")
    idis = _qual_count(text, "IDI")
    if fgds > 0: result["num_fgds"] = fgds
    if idis > 0: result["num_idis"] = idis

    sample = _sample_size(text)
    if sample > 0: result["sample_size"] = sample

    langs = _languages(text)
    if langs: result["languages"] = langs

    blocks = _num_blocks(text)
    if blocks > 1: result["num_blocks"] = blocks

    return result
