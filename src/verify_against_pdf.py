"""
Checks every extracted number against the text of the source PDF.

For each field it builds the ways the number can be printed in a filing
(85,777 · (2,985) · 1.40 · 15.9% · $26.0 billion) and looks for one of them
on the cited page. Outcomes:

  ON PAGE     found on the page Reducto cited (or the expected page)
  OTHER PAGE  found in the document, on a different page (check the citation)
  NOT FOUND   not printed anywhere in the document (investigate)
  NULL        the field returned no value (listed in Notes as omitted)

Usage (Colab or local):
    from verify_against_pdf import verify_payload, verify_values
    rows = verify_payload(apple_raw, "apple_10q.pdf")          # Reducto output + citations
    rows = verify_values({"q3_fy2024_total_net_sales": (85777, 4)}, "apple_10q.pdf")
"""

import re
import sys
import json

import pdfplumber


def page_texts(pdf_path):
    """PDF page number (1-based) -> text with whitespace collapsed."""
    with pdfplumber.open(pdf_path) as pdf:
        return {i: re.sub(r"\s+", " ", p.extract_text() or "") for i, p in enumerate(pdf.pages, 1)}


def printed_forms(value):
    """Ways a number can appear in an SEC filing or earnings release."""
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return []
    v = abs(value)
    forms = set()
    if float(v).is_integer():
        n = int(v)
        forms.add(f"{n:,}")
        if n >= 1000:  # USD millions stated in billions in the narrative
            b = n / 1000
            forms.update({f"{b:.1f} billion", f"{b:g} billion"})
    else:
        forms.update({f"{v:.2f}", f"{v:.1f}", f"{v:g}"})
    forms.update({f + "%" for f in list(forms) if "billion" not in f})
    if value < 0:
        forms.update({f"({f})" for f in list(forms)})
    return sorted(forms, key=len, reverse=True)


def _found(text, forms):
    for f in forms:
        # number must not be part of a longer number (e.g. 85,777 inside 185,777)
        if re.search(r"(?<![\d,.])" + re.escape(f) + r"(?![\d])", text):
            return f
    return None


def check(field, value, page, texts):
    if value is None:
        return {"field": field, "value": None, "page": page, "status": "NULL", "matched": ""}
    forms = printed_forms(value)
    if not forms:
        return {"field": field, "value": value, "page": page, "status": "TEXT", "matched": ""}
    if page and (hit := _found(texts.get(page, ""), forms)):
        return {"field": field, "value": value, "page": page, "status": "ON PAGE", "matched": hit}
    for p, text in texts.items():
        if hit := _found(text, forms):
            return {"field": field, "value": value, "page": page, "status": "OTHER PAGE", "matched": f"{hit} (p.{p})"}
    return {"field": field, "value": value, "page": page, "status": "NOT FOUND", "matched": ""}


def verify_values(values, pdf_path):
    """values: field -> (value, expected_page)."""
    texts = page_texts(pdf_path)
    return [check(f, v, p, texts) for f, (v, p) in values.items()]


def verify_payload(raw, pdf_path):
    """raw: Reducto extract payload (each field = {'value': ..., 'citations': [...]})."""
    texts = page_texts(pdf_path)
    rows = []
    for field, item in raw.items():
        value = item.get("value") if isinstance(item, dict) else item
        if isinstance(value, (list, dict, str)):
            continue  # qualitative fields are reviewed in the Disclosures tab
        cites = item.get("citations") if isinstance(item, dict) else None
        page = (cites[0].get("bbox") or {}).get("page") if cites else None
        rows.append(check(field, value, page, texts))
    return rows


def compare_to_expected(values, expected):
    """Reducto values vs. ground truth: field -> (extracted, expected, match)."""
    out = []
    for field, (exp, _) in expected.items():
        got = values.get(field)
        match = (got is None and exp is None) or (got is not None and exp is not None and abs(got - exp) < 0.005)
        out.append((field, got, exp, match))
    return out


def summary(rows):
    counts = {}
    for r in rows:
        counts[r["status"]] = counts.get(r["status"], 0) + 1
    return counts


if __name__ == "__main__":
    # python src/verify_against_pdf.py tests/apple_q3_fy2024_expected.json path/to/10-Q.pdf
    expected = json.load(open(sys.argv[1]))["values"]
    rows = verify_values(expected, sys.argv[2])
    for r in rows:
        if r["status"] != "ON PAGE":
            print(f'{r["status"]:<10} {r["field"]:<52} {r["value"]!s:<12} p.{r["page"]} {r["matched"]}')
    print(summary(rows))
