"""
surcharge_extract.py
────────────────────
從 COSCO PDF 萃取 surcharge 數據。
"""

import re
import pdfplumber


def _num(text):
    m = re.search(r'[\d]+\.?\d*', str(text).replace(',', ''))
    return float(m.group()) if m else None


def extract_surcharges(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ""
            if "BUC" in t and "ETS" in t:
                text = t
                break

    if not text:
        return {}

    result = {}
    lines = text.split("\n")

    # ── BUC ──────────────────────────────────────────────────
    m = re.search(r'BUC.*?\$\s*([\d.,]+)\s*/TEU', text, re.IGNORECASE)
    if m:
        v = _num(m.group(1))
        result['buc_20'] = v
        result['buc_40'] = v * 2 if v else None

    # ── ETS ──────────────────────────────────────────────────
    m = re.search(r'ETS.*?EUR\s*([\d.,]+)\s*/TEU', text, re.IGNORECASE)
    if m:
        result['ets_eur'] = _num(m.group(1))

    # ── PSS ──────────────────────────────────────────────────
    m = re.search(r'PSS.*?USD\s*([\d.,]+)\s*/teu', text, re.IGNORECASE)
    if m:
        v = _num(m.group(1))
        result['pss_20'] = v
        result['pss_40'] = v * 2 if v else None
    m = re.search(r'PSS.*?effective\s+Date[:\s]+(.+?)(?:USD|\n)', text, re.IGNORECASE)
    result['pss_note'] = m.group(1).strip(' -') if m else ''

    # ── EBS（動態：收集含 "Emergency Bunker" 的行 + 緊接的下一行）──
    ebs_list = []
    for idx, line in enumerate(lines):
        if "Emergency Bunker" in line:
            # 這行加上下一行（下一行可能有第二個日期）
            candidate_lines = [line]
            if idx + 1 < len(lines):
                next_line = lines[idx + 1]
                # 下一行不是另一個 surcharge 的開頭才加入
                if not re.match(r'^[A-Z]{2,}[^a-z]', next_line.strip()):
                    candidate_lines.append(next_line)
            combined = " ".join(candidate_lines)
            # 找所有 from [日期] USD [數字] 的模式
            for m in re.finditer(
                r'from\s+([\w\s]+?\d{4})\s*:?\s*USD\s*([\d.,]+)',
                combined, re.IGNORECASE
            ):
                v = _num(m.group(2))
                ebs_list.append({
                    'note': "from " + m.group(1).strip(),
                    '20':   v,
                    '40':   v * 2 if v else None,
                })
    result['ebs'] = ebs_list

    # ── WHF ──────────────────────────────────────────────────
    m = re.search(r'Miami.*?\$\s*([\d.,]+)\s*/20.*?\$\s*([\d.,]+)\s*/40',
                  text, re.IGNORECASE)
    if m:
        result['whf_miami_20'] = _num(m.group(1))
        result['whf_miami_40'] = _num(m.group(2))

    m = re.search(r'Houston.*?\$\s*([\d.,]+)\s*/unit', text, re.IGNORECASE)
    if m:
        v = _num(m.group(1))
        result['whf_houston_20'] = v
        result['whf_houston_40'] = v

    m = re.search(r'(?:New Orleans|Orleans).*?\$\s*([\d.,]+)\s*/unit',
                  text, re.IGNORECASE)
    if m:
        v = _num(m.group(1))
        result['whf_nola_20'] = v
        result['whf_nola_40'] = v

    return result


if __name__ == "__main__":
    import sys, json
    path = sys.argv[1] if len(sys.argv) > 1 else "cosco_pdf.pdf"
    data = extract_surcharges(path)
    print(json.dumps(data, indent=2))
