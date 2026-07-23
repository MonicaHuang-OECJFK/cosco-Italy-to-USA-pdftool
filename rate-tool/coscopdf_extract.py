"""
cosco_extract.py
────────────────
從 COSCO PDF 萃取 Direct Ports 與 Outports 的：
  POL / POD / rate_20 / rate_40

設計原則：從 header 列讀取欄位語意，不 hardcode col 號（除了 POL）。
  - pod_col / rate20_col / rate40_col：從 header 關鍵字動態找
  - 每個 col 讀取時同時試 col 和 col-1（±1 容錯漂移）
  - POL col：Direct Ports 固定 [2]，Outports 固定 [4, 5]
    （這兩個是 Service/Via 之後、POD 之前的固定欄，跨版本一致）

支援任意欄數的 PDF 版本（已驗證 8 / 12 / 22 cols）。
偵測兩表是否合併：table[0] 裡有沒有第二個 'POD (Terminal)' header。

依賴：pip install pdfplumber
用法：python cosco_extract.py <path_to_pdf>
"""

import re
import sys
import pdfplumber


# ── 工具函式 ──────────────────────────────────────────────────

def clean(val):
    if val is None:
        return ""
    return str(val).replace("\n", " ").strip()

def parse_usd(val):
    """
    '1.475 USD' → 1475（歐式千位分隔）
    必須包含 USD 或 $ 才算是 rate，純數字（如 Transit '57'）回傳 None。
    """
    s = clean(val)
    if not re.search(r'USD|\$', s, re.IGNORECASE):
        return None
    s = re.sub(r"[^\d.,]", "", s)
    if not s:
        return None
    if re.match(r"^\d{1,3}\.\d{3}$", s):
        return int(s.replace(".", ""))
    try:
        return int(float(s))
    except ValueError:
        return None

def fill_down(values):
    last = ""
    result = []
    for v in values:
        c = clean(v)
        if c:
            last = c
        result.append(last)
    return result

def read_col(row, candidates):
    """從 candidates 依序取第一個有值的欄"""
    for c in candidates:
        if 0 <= c < len(row) and clean(row[c]):
            return clean(row[c])
    return ""

def parse_col(row, candidates):
    """從 candidates 依序取第一個能 parse 的 USD 值"""
    for c in candidates:
        if 0 <= c < len(row):
            val = parse_usd(row[c])
            if val is not None:
                return val
    return None

def is_header_row(row):
    return "pod (terminal)" in " ".join(clean(v) for v in row).lower()

def detect_split(tables):
    """回傳 ('B', split_idx) 若兩表合一，否則 ('A', None)"""
    for i in range(1, len(tables[0])):
        if is_header_row(tables[0][i]):
            return "B", i
    return "A", None


# ── Header 解析 ───────────────────────────────────────────────

def build_col_map(header_rows):
    """
    從 header 找 pod_col、rate20_col、rate40_col。
    回傳 dict，每個 key 對應 [header_col, header_col-1] 兩個候選。
    """
    keywords = {
        "pod":      "pod (terminal)",
        "rate_20":  "20'",
        "rate_40":  "40'",
        "rate_ref": "rate ref",
    }
    mapping = {}
    for row in header_rows:
        for col, val in enumerate(row):
            v = clean(val).lower()
            for field, kw in keywords.items():
                if field not in mapping and kw in v:
                    candidates = [col, col - 1] if col > 0 else [col]
                    mapping[field] = candidates
    return mapping


# ── 通用萃取 ──────────────────────────────────────────────────

def extract_rows(data_rows, col_map, pol_cols):
    """
    根據 col_map 讀取每列資料。

    col_map：{'pod': [col, col-1], 'rate_20': [...], 'rate_40': [...]}
    pol_cols：POL 的候選 col 列表（Direct=[2], Outports=[4,5]）

    接續列：rate_20 的所有候選欄都沒有值 → 接續列，把 POD 串到上一筆。

    Outports 群組內 fill_up（Ancona New York 問題）：
      Service 欄（col 0）有值 → 群組起點，群組內先 fill_down 再 fill_up。
    """
    pod_cands    = col_map.get("pod",     [3, 2])
    rate20_cands = col_map.get("rate_20", [7, 6])
    rate40_cands = col_map.get("rate_40", [8, 7])

    # POL fill_down + 群組內 fill_up
    pols_raw     = [read_col(r, pol_cols) for r in data_rows]
    services_raw = [clean(r[0]) for r in data_rows]
    group_starts = [i for i, s in enumerate(services_raw) if s] + [len(data_rows)]

    pols = fill_down(pols_raw)
    if len(group_starts) > 2:   # 多群組 → 群組內 fill_up
        pols = [""] * len(pols_raw)
        for g in range(len(group_starts) - 1):
            start, end = group_starts[g], group_starts[g + 1]
            gv = pols_raw[start:end]
            filled = fill_down(gv)
            next_val = ""
            for i in range(len(gv) - 1, -1, -1):
                if gv[i]:
                    next_val = gv[i]
                elif next_val:
                    filled[i] = next_val
            for i, v in enumerate(filled):
                pols[start + i] = v

    results = []
    for i, row in enumerate(data_rows):
        r20 = parse_col(row, rate20_cands)
        r40 = parse_col(row, rate40_cands)

        if r20 is None:
            # 接續列：POD 碎片串到上一筆
            pod_frag = read_col(row, pod_cands)
            if results and pod_frag:
                results[-1]["pod"] += " " + pod_frag
            continue

        # POD：試所有候選欄（含 col+1 應對 Miami 右漂）
        pod_cands_extended = pod_cands + [pod_cands[0] + 1]
        pod = read_col(row, pod_cands_extended)

        if not pod:
            continue

        results.append({
            "pol":     pols[i],
            "pod":     pod,
            "rate_20": r20,
            "rate_40": r40,
        })

    return results


# ── Direct Ports / Outports ───────────────────────────────────

def parse_direct(raw):
    """
    header 1~2 列。
    POL 固定在 col 2（跨所有版本一致：Service=0, Via=1, POL=2）。
    """
    n_header = 1
    for i in range(1, min(3, len(raw))):
        if is_header_row(raw[i]) or not any(clean(v) for v in raw[i]):
            n_header = i + 1
        else:
            break

    col_map = build_col_map(raw[:n_header])
    pod_c   = col_map["pod"][0] if "pod" in col_map else 3
    pol_cols = [2] if pod_c <= 4 else [4]
    return extract_rows(raw[n_header:], col_map, pol_cols=pol_cols)


def parse_outports(raw):
    """
    header 1 列。
    POL col 根據 pod_col 動態決定：
      pod_col <= 4（窄版如 cosco5 8 cols）→ pol_cols = [2]
      pod_col > 4 （寬版如 cosco_pdf 22 cols）→ pol_cols = [4, 5]
    """
    col_map = build_col_map(raw[:1])
    pod_c   = col_map["pod"][0] if "pod" in col_map else 3
    pol_cols = [2] if pod_c <= 4 else [4, 5]
    return extract_rows(raw[1:], col_map, pol_cols=pol_cols)


# ── Rate Reference（例如 'TLI GL JULY'）────────────────────────

def extract_rate_ref(pdf_path):
    """
    萃取 Rate Reference 代碼（例如 'TLI GL JULY'），
    這個值整份 PDF 通常只有一種，寫在 Direct Ports 表格最右邊那欄。

    做法：
      1. 優先用 header 關鍵字 'rate ref' 找欄位（跟 pod/rate_20 一樣 ±1 容錯）
      2. 找不到 header 就 fallback：掃描第一頁所有表格，
         抓第一個符合 'TLI ...' 開頭 pattern 的儲存格
    """
    with pdfplumber.open(pdf_path) as pdf:
        tables = pdf.pages[0].extract_tables()

    if tables:
        col_map = build_col_map(tables[0][:2])
        if "rate_ref" in col_map:
            for row in tables[0][2:]:
                val = read_col(row, col_map["rate_ref"])
                if val:
                    return val

    for table in tables:
        for row in table:
            for cell in row:
                v = clean(cell)
                if re.match(r"^TLI\b", v, re.IGNORECASE):
                    return v

    return None


# ── US Inland（Rail Ramp）表格 ──────────────────────────────────

def build_ramp_col_map(header_rows):
    """
    從 Rail Ramp 表格的 header 找 location / via_pod / rate_20 / rate_40 欄。
    每個欄位漂移方向不固定，所以候選欄同時含 col-1 / col / col+1。
    """
    keywords = {
        "location": "rail ramp",
        "via_pod":  "via pod",
        "rate_20":  "20'box",
        "rate_40":  "40'bx",
    }
    mapping = {}
    for row in header_rows:
        for col, val in enumerate(row):
            v = clean(val).lower()
            for field, kw in keywords.items():
                if field not in mapping and kw in v:
                    lo = col - 1 if col > 0 else col
                    mapping[field] = [col, col + 1, lo]
    return mapping


def _detect_ramp_periods(header_rows):
    """
    有些 PDF 的 Rail Ramp 表格會同時列出多組期別（例如 APRIL / MAY 兩組
    20'box / 40'bx/HC 費率並排）。這個函式偵測是否有這種狀況。

    做法：掃描第二列 header（期別列，例如 'APRIL' / 'MAY'），
    每個有值的欄位，往同一欄（含 ±1 容錯）找第一列 header 是
    "20'box" 還是 "40'bx/HC"，藉此歸類。

    回傳 {period_label: {"rate_20": [col...], "rate_40": [col...]}}
    只有一組期別（沒有這種期別列）就回傳 {}
    """
    if len(header_rows) < 2:
        return {}

    row0, row1 = header_rows[0], header_rows[1]
    periods = {}
    for col, val in enumerate(row1):
        label = clean(val).upper()
        if not label:
            continue

        kw = ""
        for c in (col, col - 1, col + 1):
            if 0 <= c < len(row0):
                v = clean(row0[c]).lower()
                if "20'box" in v or "40'bx" in v:
                    kw = v
                    break

        if "20'box" in kw:
            periods.setdefault(label, {}).setdefault("rate_20", []).append(col)
        elif "40'bx" in kw:
            periods.setdefault(label, {}).setdefault("rate_40", []).append(col)

    return periods


def list_us_inland_periods(pdf_path):
    """
    檢查 PDF 的 Rail Ramp 表格是否同時列出多組期別（例如 ['APRIL', 'MAY']）。
    沒有這種狀況（只有一組費率）就回傳空 list。
    """
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                header_blob = " ".join(
                    clean(v) for row in table[:2] for v in row
                ).lower()
                if "rail ramp" in header_blob and "via pod" in header_blob:
                    return list(_detect_ramp_periods(table[:2]).keys())
    return []


def _parse_ramp_rows(table, period=None):
    col_map = build_ramp_col_map(table[:2])

    loc_cands = col_map.get("location", [1, 2, 0])
    pod_cands = col_map.get("via_pod",  [4, 3, 5])

    periods = _detect_ramp_periods(table[:2])
    if periods:
        if period is None:
            raise ValueError(
                f"此 PDF 的 Rail Ramp 表格有多組期別：{', '.join(periods)}，"
                f"請指定 period 參數（例如 period='MAY'）"
            )
        period_key = period.strip().upper()
        if period_key not in periods:
            raise ValueError(
                f"找不到期別「{period}」，PDF 裡的期別有：{', '.join(periods)}"
            )
        r20_cands = periods[period_key].get("rate_20", col_map.get("rate_20", [7, 6, 8]))
        r40_cands = periods[period_key].get("rate_40", col_map.get("rate_40", [9, 8, 10]))
    else:
        r20_cands = col_map.get("rate_20", [7, 6, 8])
        r40_cands = col_map.get("rate_40", [9, 8, 10])

    results = []
    for row in table[2:]:
        location = read_col(row, loc_cands)
        via_pod  = read_col(row, pod_cands)
        if not location or not via_pod:
            continue

        results.append({
            "location": location,
            "via_pod":  via_pod,
            "rate_20":  parse_col(row, r20_cands),   # 'No 20box' 沒有 USD → None
            "rate_40":  parse_col(row, r40_cands),
        })

    return results


def extract_us_inland(pdf_path, period=None):
    """
    萃取 'CY US RAMP' 表格（Rail Ramp Location / VIA POD / 20'box / 40'bx/HC）。
    掃描每一頁的每個表格，找到含 'rail ramp' 與 'via pod' header 的那個。

    參數：
      period : 如果該 PDF 的 Rail Ramp 表格同時列出多組期別（例如 'APRIL' / 'MAY'），
               需要指定要抓哪一組；只有一組期別時可省略。
               可用 list_us_inland_periods() 先查詢有哪些期別可選。

    回傳 list of dict：{"location", "via_pod", "rate_20", "rate_40"}
    找不到 Rail Ramp 表格就回傳空 list。
    """
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                header_blob = " ".join(
                    clean(v) for row in table[:2] for v in row
                ).lower()
                if "rail ramp" in header_blob and "via pod" in header_blob:
                    return _parse_ramp_rows(table, period=period)
    return []


# ── 對外介面 ──────────────────────────────────────────────────

def extract(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        tables = pdf.pages[0].extract_tables()

    version, split_row = detect_split(tables)

    if version == "A":
        direct   = parse_direct(tables[0])
        outports = parse_outports(tables[1])
    else:
        raw      = tables[0]
        direct   = parse_direct(raw[:split_row])
        outports = parse_outports(raw[split_row:])

    return direct, outports


if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else "cosco_pdf.pdf"
    direct, outports = extract(path)
    print(f"\n── Direct Ports ({len(direct)} rows) ──")
    for r in direct:
        print(r)
    print(f"\n── Outports ({len(outports)} rows) ──")
    for r in outports:
        print(r)


def split_pol(rows):
    """
    把 POL 用 '/' 拆開成獨立列，其餘欄位複製。

    例如：
    {'pol': 'La Spezia (LSCT) / Genova (VTE)', 'pod': 'New York', ...}
    →
    {'pol': 'La Spezia (LSCT)', 'pod': 'New York', ...}
    {'pol': 'Genova (VTE)',     'pod': 'New York', ...}
    """
    result = []
    for row in rows:
        pol_parts = [p.strip() for p in row["pol"].split("/") if p.strip()]
        if not pol_parts:
            pol_parts = [row["pol"]]
        for pol in pol_parts:
            result.append({
                "pol":     pol,
                "pod":     row["pod"],
                "rate_20": row["rate_20"],
                "rate_40": row["rate_40"],
            })
    return result
