# -*- coding: utf-8 -*-
"""
套装计算 - Streamlit 网页版
==========================
部署方式：Streamlit Cloud，生成公开链接即可使用。
本地运行：streamlit run app.py
"""

import os
import tempfile

import streamlit as st
from openpyxl import load_workbook

# ---------------------------------------------------------------------------
# 配置与常量（与桌面版一致）
# ---------------------------------------------------------------------------
PROJECT_KEYS = {
    "ANTA": ("商品料号", "STYLE"),
    "VF": ("商品料号", "STYLE"),
    "Carters": ("商品料号", "ITEM/SKU"),
}

HDR_SCAN_ROWS = 30
INSERT_AT = 14       # M 列（第 13 列）之后插入，新列从第 14 列（N）开始
QTY_COL, BOM_COL, REL_COL = 14, 15, 16
INSERT_HEADERS = ("QTY", "bom", "套装关系")

QTY_HEADERS = ("数量", "QTY", "QUANTITY", "件数")
REL_A = "A"


# ---------------------------------------------------------------------------
# 通用工具（与桌面版一致）
# ---------------------------------------------------------------------------
def _text(value):
    return "" if value is None else str(value).strip()


def _hkey(value):
    return "".join(_text(value).split()).upper()


def _key_text(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _to_number(value):
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _is_integer_bom(value):
    num = _to_number(value)
    return num is not None and float(num).is_integer()


def find_header_row(ws, names):
    targets = {_hkey(n) for n in names}
    limit = min(ws.max_row or 1, HDR_SCAN_ROWS)
    for row in range(1, limit + 1):
        for col in range(1, (ws.max_column or 1) + 1):
            if _hkey(ws.cell(row, col).value) in targets:
                return row
    raise ValueError("未能在表格前 %d 行中找到表头：%s" % (HDR_SCAN_ROWS, " / ".join(names)))


def header_map(ws, header_row):
    mapping = {}
    for col in range(1, (ws.max_column or 1) + 1):
        key = _hkey(ws.cell(header_row, col).value)
        if key and key not in mapping:
            mapping[key] = col
    return mapping


def col_exact(mapping, names):
    for name in names:
        idx = mapping.get(_hkey(name))
        if idx:
            return idx
    return None


def col_contains(mapping, keyword):
    target = _hkey(keyword)
    for key, idx in mapping.items():
        if target in key:
            return idx
    return None


# ---------------------------------------------------------------------------
# 核心计算（与桌面版一致）
# ---------------------------------------------------------------------------
def _calc_group(group, suit_rows):
    result = {rec["row"]: ("NA", "NA", rec["qty_raw"]) for rec in group}
    if not suit_rows:
        return result

    for index, rec in enumerate(group):
        bom_raw, rel_raw = suit_rows[min(index, len(suit_rows) - 1)]
        rec["suit_bom"] = bom_raw
        rec["suit_rel"] = _text(rel_raw)

    # 规则 1：一一对应
    if len(suit_rows) == 1:
        for rec in group:
            bom = _to_number(rec["suit_bom"])
            rel = rec["suit_rel"]
            qty = None
            rel_is_blank = rel == ""
            if rec["num"] is not None and bom and (
                rel.upper() == REL_A or (rel_is_blank and _is_integer_bom(rec["suit_bom"]))
            ):
                qty = round(rec["num"] / bom, 4)
            result[rec["row"]] = (rec["suit_bom"], rel, qty)
        return result

    # 规则 4：多条不对等
    if len(group) != len(suit_rows):
        first_bom, first_rel_raw = suit_rows[0]
        first_rel = _text(first_rel_raw)
        first_bom_num = _to_number(first_bom)
        first_rel_blank = first_rel == ""
        fill_qty = (
            (first_rel.upper() == REL_A or first_rel_blank)
            and _is_integer_bom(first_bom)
            and first_bom_num
        )
        for rec in group:
            qty = None
            if fill_qty and rec["num"] is not None:
                qty = round(rec["num"] / first_bom_num, 4)
            result[rec["row"]] = (first_bom, first_rel, qty)
        return result

    # 规则 2：多条一致
    a_rec = next((r for r in group if r["suit_rel"].upper() == REL_A), None)
    all_bom_int = all(_is_integer_bom(bom) for bom, _ in suit_rows)

    if not all_bom_int:
        qty_all = a_rec["qty_raw"] if a_rec is not None else None
        for rec in group:
            result[rec["row"]] = ("NA", rec["suit_rel"], qty_all)
        return result

    all_rel_a = all(_text(rel).upper() == REL_A for _, rel in suit_rows)
    for rec in group:
        rel = rec["suit_rel"]
        bom = _to_number(rec["suit_bom"])
        qty = None
        should_divide = rel == "" or (
            rel.upper() == REL_A and (all_rel_a or rec is a_rec)
        )
        if should_divide and rec["num"] is not None and bom:
            qty = round(rec["num"] / bom, 4)
        result[rec["row"]] = (rec["suit_bom"], rel, qty)
    return result


def build_result_bytes(packing_path, suit_path, project):
    """返回 (输出字节, 记录数)，Streamlit 直接给用户下载。"""
    key_names = PROJECT_KEYS[project]

    # 套装表
    wb_suit = load_workbook(suit_path, data_only=True)
    ws_suit = wb_suit.worksheets[0]
    suit_header_row = find_header_row(ws_suit, key_names)
    suit_map = header_map(ws_suit, suit_header_row)
    suit_key_cols = [col_exact(suit_map, [n]) for n in key_names]
    if None in suit_key_cols:
        raise ValueError("套装表中未找到主键表头：%s（当前项目：%s）" % (" / ".join(key_names), project))

    suit_bom_col = col_contains(suit_map, "bom")
    suit_rel_col = col_contains(suit_map, "套装关系")
    if not suit_bom_col or not suit_rel_col:
        raise ValueError("套装表中未找到 “bom” 或 “套装关系” 列")

    entries = {}
    for row in range(suit_header_row + 1, (ws_suit.max_row or suit_header_row) + 1):
        key = tuple(_key_text(ws_suit.cell(row, c).value) for c in suit_key_cols)
        if all(k == "" for k in key):
            continue
        entries.setdefault(key, []).append(
            (ws_suit.cell(row, suit_bom_col).value, ws_suit.cell(row, suit_rel_col).value)
        )

    # Packinglist
    wb_value = load_workbook(packing_path, data_only=True)
    ws_value = wb_value.worksheets[0]
    packing_header_row = find_header_row(ws_value, key_names)
    packing_map = header_map(ws_value, packing_header_row)
    packing_key_cols = [col_exact(packing_map, [n]) for n in key_names]
    if None in packing_key_cols:
        raise ValueError("Packinglist 中未找到主键表头：%s（当前项目：%s）" % (" / ".join(key_names), project))

    qty_col = col_exact(packing_map, QTY_HEADERS) or col_contains(packing_map, "数量")
    if not qty_col:
        raise ValueError("Packinglist 中未找到数量列（如：数量 / QTY）")

    records = []
    for row in range(packing_header_row + 1, (ws_value.max_row or packing_header_row) + 1):
        key = tuple(_key_text(ws_value.cell(row, c).value) for c in packing_key_cols)
        if all(k == "" for k in key):
            continue
        qty_raw = ws_value.cell(row, qty_col).value
        records.append({
            "row": row,
            "key": key,
            "num": _to_number(qty_raw),
            "qty_raw": qty_raw,
        })

    groups = {}
    for rec in records:
        groups.setdefault(rec["key"], []).append(rec)

    filled = {}
    for key, group in groups.items():
        filled.update(_calc_group(group, entries.get(key, [])))

    # 写出到临时文件再读 bytes
    wb_out = load_workbook(packing_path)
    ws_out = wb_out[ws_value.title]
    ws_out.insert_cols(INSERT_AT, 3)
    for offset, name in enumerate(INSERT_HEADERS):
        ws_out.cell(packing_header_row, INSERT_AT + offset).value = name
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
    with open(tmp_path, "rb") as f:
        data = f.read()
    os.remove(tmp_path)
    return data, len(records)


# ---------------------------------------------------------------------------
# Streamlit 界面
# ---------------------------------------------------------------------------
st.set_page_config(page_title="套装计算", page_icon="📦", layout="centered")

st.markdown(
    """
    <style>
        .stApp { background: #f2f2f2; }
        .block-container { max-width: 720px; padding-top: 2rem; }
        .card {
            background: linear-gradient(135deg, #34f0f5 0%, #32c8cf 100%);
            border-radius: 20px;
            padding: 2.5rem 2rem;
            box-shadow: 0 4px 24px rgba(0,0,0,0.08);
        }
        h1 { color: #32c8cf !important; font-size: 1.6rem !important; }
        label, .stMarkdown { color: #ffffff !important; font-weight: 600 !important; }
        .stFileUploader, .stSelectbox { margin-bottom: 1.5rem; }
        .stButton button {
            background: #ffffff !important;
            color: #32c8cf !important;
            border: none !important;
            font-weight: 600 !important;
            padding: 0.6rem 2rem !important;
            border-radius: 6px !important;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📦 套装计算")
st.caption("上传 Packinglist 和套装表，一键生成 QTY / bom / 套装关系")

with st.container():
    project = st.selectbox(
        "请选择项目",
        list(PROJECT_KEYS.keys()),
        index=2,  # 默认 Carters
        help="主键组成：ANTA / VF = 商品料号 + STYLE；Carters = 商品料号 + ITEM/SKU",
    )

    packing_file = st.file_uploader(
        "请选择 Packinglist 表上传",
        type=["xlsx", "xlsm"],
        key="packing",
    )

    suit_file = st.file_uploader(
        "请选择套装表上传",
        type=["xlsx", "xlsm"],
        key="suit",
    )

    if st.button("🚀 开始计算", type="primary", use_container_width=True):
        if not packing_file or not suit_file:
            st.error("请同时上传 Packinglist 和套装表")
        else:
            with st.spinner("正在计算，请稍候..."):
                try:
                    # 把上传的临时文件落盘（openpyxl 需要真实路径）
                    tmp_packing = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
                    tmp_packing.write(packing_file.read())
                    tmp_packing.close()
                    tmp_suit = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
                    tmp_suit.write(suit_file.read())
                    tmp_suit.close()

                    data, count = build_result_bytes(tmp_packing.name, tmp_suit.name, project)
                    os.remove(tmp_packing.name)
                    os.remove(tmp_suit.name)

                    out_name = os.path.splitext(packing_file.name)[0] + "_套装计算.xlsx"
                    st.success(f"✅ 共处理 {count} 行数据")
                    st.download_button(
                        label="📥 下载结果文件",
                        data=data,
                        file_name=out_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                except Exception as exc:
                    st.error(f"计算失败：{exc}")

st.divider()
st.caption("🔒 所有处理均在浏览器会话中本地完成，数据不会上传到任何服务器")
