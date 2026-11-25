# -*- coding: utf-8 -*-
import os, tempfile, streamlit as st
from openpyxl import load_workbook

PROJECT_KEYS = {"ANTA": ("商品料号", "STYLE"), "VF": ("商品料号", "STYLE"), "Carters": ("商品料号", "ITEM/SKU")}
HDR_SCAN_ROWS = 30
INSERT_AT = 14
QTY_COL, BOM_COL, REL_COL = 14, 15, 16
INSERT_HEADERS = ("QTY", "bom", "套装关系")
QTY_HEADERS = ("数量", "QTY", "QUANTITY", "件数")
REL_A = "A"

def _text(v): return "" if v is None else str(v).strip()
def _hkey(v): return "".join(_text(v).split()).upper()
def _key_text(v):
    if v is None: return ""
    if isinstance(v, float) and v.is_integer(): return str(int(v))
    return str(v).strip()
def _to_number(v):
    if v is None or isinstance(v, bool): return None
    if isinstance(v, (int, float)): return float(v)
    t = str(v).strip().replace(",", "")
    if not t: return None
    try: return float(t)
    except ValueError: return None
def _is_integer_bom(v):
    n = _to_number(v)
    return n is not None and float(n).is_integer()

def find_header_row(ws, names):
    targets = {_hkey(n) for n in names}
    limit = min(ws.max_row or 1, HDR_SCAN_ROWS)
    for row in range(1, limit + 1):
        for col in range(1, (ws.max_column or 1) + 1):
            if _hkey(ws.cell(row, col).value) in targets:
                return row
    raise ValueError("未能在表格前 %d 行中找到表头：%s" % (HDR_SCAN_ROWS, " / ".join(names)))

def header_map(ws, hrow):
    m = {}
    for col in range(1, (ws.max_column or 1) + 1):
        k = _hkey(ws.cell(hrow, col).value)
        if k and k not in m: m[k] = col
    return m

def col_exact(m, names):
    for n in names:
        idx = m.get(_hkey(n))
        if idx: return idx
    return None

def col_contains(m, kw):
    t = _hkey(kw)
    for k, idx in m.items():
        if t in k: return idx
    return None

def _calc_group(group, suit_rows):
    result = {rec["row"]: ("NA", "NA", rec["qty_raw"]) for rec in group}
    if not suit_rows: return result

    for i, rec in enumerate(group):
        bom_raw, rel_raw = suit_rows[min(i, len(suit_rows) - 1)]
        rec["suit_bom"] = bom_raw
        rec["suit_rel"] = _text(rel_raw)

    if len(suit_rows) == 1:
        for rec in group:
            bom = _to_number(rec["suit_bom"])
            rel = rec["suit_rel"]
            qty = None
            rel_blank = rel == ""
            if rec["num"] is not None and bom and (rel.upper() == REL_A or (rel_blank and _is_integer_bom(rec["suit_bom"]))):
                qty = round(rec["num"] / bom, 4)
            result[rec["row"]] = (rec["suit_bom"], rel, qty)
        return result

    if len(group) != len(suit_rows):
        first_bom, first_rel_raw = suit_rows[0]
        first_rel = _text(first_rel_raw)
        first_bom_num = _to_number(first_bom)
        first_rel_blank = first_rel == ""
        fill_qty = ((first_rel.upper() == REL_A or first_rel_blank) and _is_integer_bom(first_bom) and first_bom_num)
        for rec in group:
            qty = None
            if fill_qty and rec["num"] is not None: qty = round(rec["num"] / first_bom_num, 4)
            result[rec["row"]] = (first_bom, first_rel, qty)
        return result

    a_rec = next((r for r in group if r["suit_rel"].upper() == REL_A), None)
    all_bom_int = all(_is_integer_bom(bom) for bom, _ in suit_rows)

    if not all_bom_int:
        qty_all = a_rec["qty_raw"] if a_rec is not None else None
        for rec in group: result[rec["row"]] = ("NA", rec["suit_rel"], qty_all)
        return result

    all_rel_a = all(_text(rel).upper() == REL_A for _, rel in suit_rows)
    for rec in group:
        rel = rec["suit_rel"]
        bom = _to_number(rec["suit_bom"])
        qty = None
        should_divide = rel == "" or (rel.upper() == REL_A and (all_rel_a or rec is a_rec))
        if should_divide and rec["num"] is not None and bom: qty = round(rec["num"] / bom, 4)
        result[rec["row"]] = (rec["suit_bom"], rel, qty)
    return result

def build_result_bytes(packing_path, suit_path, project):
    key_names = PROJECT_KEYS[project]

    wb_suit = load_workbook(suit_path, data_only=True)
    ws_suit = wb_suit.worksheets[0]
    suit_hrow = find_header_row(ws_suit, key_names)
    suit_map = header_map(ws_suit, suit_hrow)
    suit_key_cols = [col_exact(suit_map, [n]) for n in key_names]
    if None in suit_key_cols: raise ValueError("套装表中未找到主键表头：%s（当前项目：%s）" % (" / ".join(key_names), project))

    suit_bom_col = col_contains(suit_map, "bom")
    suit_rel_col = col_contains(suit_map, "套装关系")
    if not suit_bom_col or not suit_rel_col: raise ValueError("套装表中未找到 bom 或 套装关系 列")

    entries = {}
    for row in range(suit_hrow + 1, (ws_suit.max_row or suit_hrow) + 1):
        key = tuple(_key_text(ws_suit.cell(row, c).value) for c in suit_key_cols)
        if all(k == "" for k in key): continue
        entries.setdefault(key, []).append((ws_suit.cell(row, suit_bom_col).value, ws_suit.cell(row, suit_rel_col).value))

    wb_value = load_workbook(packing_path, data_only=True)
    ws_value = wb_value.worksheets[0]
    pack_hrow = find_header_row(ws_value, key_names)
    pack_map = header_map(ws_value, pack_hrow)
    pack_key_cols = [col_exact(pack_map, [n]) for n in key_names]
    if None in pack_key_cols: raise ValueError("Packinglist 中未找到主键表头：%s（当前项目：%s）" % (" / ".join(key_names), project))

    qty_col = col_exact(pack_map, QTY_HEADERS) or col_contains(pack_map, "数量")
    if not qty_col: raise ValueError("Packinglist 中未找到数量列")

    records = []
    for row in range(pack_hrow + 1, (ws_value.max_row or pack_hrow) + 1):
        key = tuple(_key_text(ws_value.cell(row, c).value) for c in pack_key_cols)
        if all(k == "" for k in key): continue
        qty_raw = ws_value.cell(row, qty_col).value
        records.append({"row": row, "key": key, "num": _to_number(qty_raw), "qty_raw": qty_raw})

    groups = {}
    for rec in records: groups.setdefault(rec["key"], []).append(rec)

    filled = {}
    for key, group in groups.items(): filled.update(_calc_group(group, entries.get(key, [])))

    wb_out = load_workbook(packing_path)
    ws_out = wb_out[ws_value.title]
    ws_out.insert_cols(INSERT_AT, 3)
    for off, name in enumerate(INSERT_HEADERS): ws_out.cell(pack_hrow, INSERT_AT + off).value = name
    for rec in records:
        row = rec["row"]
        bom_out, rel_out, qty_out = filled[row]
        ws_out.cell(row, QTY_COL).value = qty_out
        ws_out.cell(row, BOM_COL).value = bom_out
        ws_out.cell(row, REL_COL).value = rel_out

    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp_path = tmp.name
    tmp.close()
    wb_out.save(tmp_path)
    with open(tmp_path, "rb") as f: data = f.read()
    os.remove(tmp_path)
    return data, len(records)

st.set_page_config(page_title="套装计算", page_icon="📦", layout="centered")

st.markdown("<style>.stApp{background:#f2f2f2}.block-container{max-width:720px;padding-top:2rem}h1{color:#32c8cf!important;font-size:1.6rem!important}.stButton button{background:#32c8cf!important;color:#ffffff!important;border:none!important;font-weight:600!important;padding:0.6rem 2rem!important;border-radius:6px!important}</style>", unsafe_allow_html=True)

st.title("📦 套装计算")
st.caption("上传 Packinglist 和套装表，一键生成 QTY / bom / 套装关系")

with st.container():
    project = st.selectbox("请选择项目", list(PROJECT_KEYS.keys()), index=2)
    packing_file = st.file_uploader("请选择 Packinglist 表上传", type=["xlsx", "xlsm"], key="packing")
    suit_file = st.file_uploader("请选择套装表上传", type=["xlsx", "xlsm"], key="suit")

    if st.button("🚀 开始计算", type="primary", use_container_width=True):
        if not packing_file or not suit_file: st.error("请同时上传 Packinglist 和套装表")
        else:
            with st.spinner("正在计算，请稍候..."):
                try:
                    tmp_p = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
                    tmp_p.write(packing_file.read()); tmp_p.close()
                    tmp_s = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
                    tmp_s.write(suit_file.read()); tmp_s.close()
                    data, count = build_result_bytes(tmp_p.name, tmp_s.name, project)
                    os.remove(tmp_p.name); os.remove(tmp_s.name)
                    out_name = os.path.splitext(packing_file.name)[0] + "_套装计算.xlsx"
                    st.success(f"✅ 共处理 {count} 行数据")
                    st.download_button("📥 下载结果文件", data=data, file_name=out_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                except Exception as exc: st.error(f"计算失败：{exc}")

st.divider()
st.caption("🔒 所有处理均在浏览器会话中本地完成，数据不会上传到任何服务器")
