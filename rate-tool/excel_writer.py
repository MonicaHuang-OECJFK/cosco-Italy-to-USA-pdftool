"""
excel_writer.py
───────────────
讀取 cheatsheet 的 Mapping tab，
把從 PDF 萃取的 OFT rate 寫入 OFT sheet 的黃色欄位。

流程：
1. 讀取 Mapping tab → 建立 (PDF_POL, PDF_POD) → (POR, POD) 對照字典
2. 掃描任意 row 找到 POD header → 得到 pod_col 和 header_row
3. POD 右邊 +2 格是 20'，往上驗證 header 是 "20'" → 確認欄位正確
4. +3 = 40'，+4 = 40HC
5. 掃描資料列，用 POR + POD 匹配，寫入 rate
"""

import openpyxl


def _find_header(ws, keyword):
    """
    掃描整張 sheet，找第一個 value == keyword 的 cell。
    回傳 (row, col) 或拋出 ValueError。
    """
    for row in ws.iter_rows():
        for cell in row:
            if str(cell.value or "").strip().upper() == keyword.upper():
                return cell.row, cell.column
    raise ValueError(f"找不到 header「{keyword}」")


def _find_oft_cols(ws, pod_col, header_row):
    """
    從 POD 欄往右 +2 格找 OFT 的 20' 欄，
    往上掃 header 驗證那格確實是 '20''。
    確認後 +3 = 40'，+4 = 40HC。

    pod_col    : POD 的 col index
    header_row : POD header 所在的 row

    回傳 (col_20, col_40, col_40hc) 或拋出 ValueError。
    """
    col_20_candidate = pod_col + 2

    # 往上掃這欄，看有沒有任何 row 的 header 是 "20'"
    confirmed = False
    for r in range(1, header_row + 1):
        val = str(ws.cell(r, col_20_candidate).value or "").strip()
        if val == "20'":
            confirmed = True
            break

    if not confirmed:
        raise ValueError(
            f"POD 右邊 +2 格（col {col_20_candidate}）往上找不到 \"20'\" header，"
            f"請確認 cheatsheet 欄位結構"
        )

    col_20   = col_20_candidate
    col_40   = col_20 + 1
    col_40hc = col_20 + 2

    return col_20, col_40, col_40hc


def _load_mapping(wb):
    """
    讀取 Mapping tab，回傳：
    {(pdf_pol_upper, pdf_pod_upper): (cheatsheet_por_upper, cheatsheet_pod_upper)}
    """
    if "Mapping" not in wb.sheetnames:
        raise ValueError("找不到 Mapping tab，請確認 cheatsheet 含有 Mapping 分頁")

    ws = wb["Mapping"]
    mapping = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        pdf_pol, pdf_pod, cs_por, cs_pod = [str(v or "").strip() for v in row[:4]]
        mapping[(pdf_pol.upper(), pdf_pod.upper())] = (cs_por.upper(), cs_pod.upper())
    return mapping


def _build_rate_lookup(all_rates, mapping):
    """
    把 PDF 萃取的 all_rates（list of dict）
    透過 mapping 轉成 {(POR_upper, POD_upper): (rate_20, rate_40)} 字典。
    """
    lookup = {}
    for row in all_rates:
        pdf_pol = row["pol"].strip().upper()
        pdf_pod = row["pod"].strip().upper()
        key = (pdf_pol, pdf_pod)
        if key in mapping:
            por, pod = mapping[key]
            lookup[(por, pod)] = (row["rate_20"], row["rate_40"])
    return lookup


def update_oft_rates(excel_path, all_rates, output_path=None, max_scan_rows=38):
    """
    主函式：把 all_rates 寫入 cheatsheet 的 OFT sheet。

    參數：
      excel_path  : cheatsheet 的路徑（含 Mapping tab）
      all_rates   : coscopdf_extract 產出的 list of dict
                    每筆含 pol / pod / rate_20 / rate_40
      output_path : 輸出路徑，None 則直接覆蓋原檔

    回傳：
      (updated_count, skipped_rows)
      updated_count : 成功寫入的列數
      skipped_rows  : 找不到對應 rate 的 (row_num, por, pod) list
    """
    wb = openpyxl.load_workbook(excel_path)
    ws = wb["OFT"]

    # 1. 讀取 mapping
    mapping = _load_mapping(wb)

    # 2. 建立 rate lookup
    rate_lookup = _build_rate_lookup(all_rates, mapping)

    # 3. 找 POR / POD header（掃描任意 row，不假設固定在哪行）
    por_header_row, por_col = _find_header(ws, "POR")
    pod_header_row, pod_col = _find_header(ws, "POD")

    # 4. 從 POD +2 找 OFT 的 20'/40'/40HC 欄，往上驗證 "20'" header
    col_20, col_40, col_40hc = _find_oft_cols(ws, pod_col, pod_header_row)

    # 5. 先清空所有 OFT 欄位的值（20' / 40' / 40HC）
    #    只清空純數值的格子，有公式的格子不動
    data_start_row = max(por_header_row, pod_header_row) + 1
    scan_end_row = data_start_row + max_scan_rows - 1
    for row_num in range(data_start_row, scan_end_row + 1):
        for col in (col_20, col_40, col_40hc):
            cell = ws.cell(row_num, col)
            if cell.data_type != 'f':   # 'f' = formula，公式格不動
                cell.value = None

    # 6. 寫入新 rate
    updated_count = 0
    skipped_rows = []

    for row_num in range(data_start_row, scan_end_row + 1):
        por = str(ws.cell(row_num, por_col).value or "").strip().upper()
        pod = str(ws.cell(row_num, pod_col).value or "").strip().upper()

        if not por or not pod:
            continue

        key = (por, pod)
        if key not in rate_lookup:
            skipped_rows.append((row_num, por, pod))
            continue

        rate_20, rate_40 = rate_lookup[key]
        ws.cell(row_num, col_20).value   = rate_20
        ws.cell(row_num, col_40).value   = rate_40
        ws.cell(row_num, col_40hc).value = rate_40   # 40HC = 40'

        updated_count += 1

    out = output_path or excel_path
    wb.save(out)
    return updated_count, skipped_rows
