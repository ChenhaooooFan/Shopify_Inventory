import streamlit as st
import pandas as pd
import re
import fitz  # PyMuPDF
from collections import defaultdict
from io import BytesIO
from datetime import datetime
import os

# =========================
# 0) UI
# =========================
st.set_page_config(page_title="NailVesta 库存系统（独立站 Slip PDF）", layout="centered")
st.title("ColorFour Inventory 系统（独立站 Slip PDF → Sold → 更新库存 + 达人换货）")

st.caption("适用文件：独立站 Slip/Invoice PDF（逐件明细），库存 CSV（SKU编码 + 单一库存列），可选达人换货表（每行=1件）。")

pdf_files = st.file_uploader("上传独立站 Slip PDF（可多选）", type=["pdf"], accept_multiple_files=True)
stock_file = st.file_uploader("上传库存表 CSV", type=["csv"])

selected_pdfs = []
if pdf_files:
    selected_names = st.multiselect(
        "选择要参与统计的 Slip PDF（默认全选）",
        options=[f.name for f in pdf_files],
        default=[f.name for f in pdf_files]
    )
    selected_pdfs = [f for f in pdf_files if f.name in selected_names]

# —— 达人换货开关 ——（逐行一件）
if "show_exchange" not in st.session_state:
    st.session_state.show_exchange = False
if st.button("有达人换货吗？"):
    st.session_state.show_exchange = True

creator_swap_df = None
if st.session_state.show_exchange:
    st.info("上传【达人换货统计表】（CSV/XLSX）：每行代表发货了 1 件。选择“原SKU列”和“新SKU列”。系统会：原SKU Sold -1、库存 +1；新SKU Sold +1、库存 -1，并生成对账表。")
    creator_swap_file = st.file_uploader("上传达人换货统计表（每行=1件）", type=["csv", "xlsx"], key="creator_swap")
    if creator_swap_file:
        if creator_swap_file.name.lower().endswith(".csv"):
            creator_swap_df = pd.read_csv(creator_swap_file)
        else:
            creator_swap_df = pd.read_excel(creator_swap_file)
        creator_swap_df.columns = [str(c).strip() for c in creator_swap_df.columns]
        st.success("达人换货统计表已上传")

# =========================
# 1) 解析：独立站 Slip PDF（逐件明细）
# =========================

# SKU 形态：ABC123-S/M/L（你当前 slip 就是这种）
SKU_RE = re.compile(r"\b([A-Z]{3}\d{3}-[SML])\b")

# 订单号（用于粗略统计每个 PDF 有多少单）
ORDER_RE = re.compile(r"\bOrder\s*#\s*([A-Za-z0-9_-]+)\b", re.I)

# “qty of total” —— 预留更鲁棒：如果未来 slip 出现 2 of 2
OF_RE = re.compile(r"\b(\d+)\s+of\s+(\d+)\b", re.I)

def normalize_text(t: str) -> str:
    t = t.replace("\u00ad", "").replace("\u200b", "").replace("\u00a0", " ")
    t = t.replace("–", "-").replace("—", "-")
    return t

def parse_slip_pdf(file_bytes: bytes):
    """
    返回：
      sku_counts: defaultdict(int)  # SKU -> qty
      extracted_units: int
      order_ids: set[str]
    规则：
      - 匹配到每个 SKU（ABC123-S/M/L）就计入 1
      - 若 SKU 附近出现 '2 of 2' 这种格式，则取前面的数量作为 qty（默认 1）
    """
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    pages_text = [p.get_text("text") for p in doc]
    text_raw = "\n".join(pages_text)
    text = normalize_text(text_raw)

    order_ids = set(ORDER_RE.findall(text))

    sku_counts = defaultdict(int)
    for m in SKU_RE.finditer(text):
        sku = m.group(1).strip()

        # look-ahead 80 chars 找 "x of y"；找不到就默认 1
        after = text[m.end(): m.end() + 80]
        mm = OF_RE.search(after)
        qty = int(mm.group(1)) if mm else 1

        sku_counts[sku] += qty

    extracted_units = int(sum(sku_counts.values()))
    return sku_counts, extracted_units, order_ids

# =========================
# 2) 主流程
# =========================
if selected_pdfs and stock_file:
    st.success("文件上传成功，开始处理...")

    # ---------- 2.1 读取库存 CSV ----------
    stock_df = pd.read_csv(stock_file)
    stock_df.columns = [str(c).strip() for c in stock_df.columns]

    # 默认列名：SKU编码 + 单一库存列（如果库存列不止一个，允许你在 UI 选）
    if "SKU编码" not in stock_df.columns:
        st.error("库存表缺少列：SKU编码")
        st.stop()

    # 选择库存列：优先找常见名字，否则让你选第一个数值列
    common_stock_names = ["Stock", "库存", "Current Stock", "Old Stock", "现有库存"]
    default_stock_col = None
    for c in common_stock_names:
        if c in stock_df.columns:
            default_stock_col = c
            break

    # 候选库存列：除 SKU编码 外的列
    candidate_cols = [c for c in stock_df.columns if c != "SKU编码"]
    if not candidate_cols:
        st.error("库存表没有可用的库存列（除 SKU编码 外至少需要一列）")
        st.stop()

    stock_col = st.selectbox(
        "选择库存列（Old Stock）",
        options=candidate_cols,
        index=(candidate_cols.index(default_stock_col) if default_stock_col in candidate_cols else 0)
    )

    # 确保库存列可计算
    stock_df[stock_col] = pd.to_numeric(stock_df[stock_col], errors="coerce").fillna(0).astype(int)

    # ---------- 2.2 解析所有 PDF ----------
    pdf_audit_rows = []
    sku_counts_all = defaultdict(int)

    for pf in selected_pdfs:
        file_bytes = pf.read()
        sku_counts, units, order_ids = parse_slip_pdf(file_bytes)

        for sku, qty in sku_counts.items():
            sku_counts_all[sku] += qty

        pdf_audit_rows.append({
            "PDF文件": pf.name,
            "识别订单数（Order #）": len(order_ids),
            "提取SKU种类数": len(sku_counts),
            "提取件数（Sold Units）": units
        })

    pdf_audit_df = pd.DataFrame(pdf_audit_rows)
    if not pdf_audit_df.empty:
        total_row = {
            "PDF文件": "合计",
            "识别订单数（Order #）": int(pdf_audit_df["识别订单数（Order #）"].sum()),
            "提取SKU种类数": int(pdf_audit_df["提取SKU种类数"].sum()),
            "提取件数（Sold Units）": int(pdf_audit_df["提取件数（Sold Units）"].sum()),
        }
        pdf_audit_df = pd.concat([pdf_audit_df, pd.DataFrame([total_row])], ignore_index=True)

    st.subheader("PDF 提取对账（独立站 Slip：逐件明细口径）")
    st.dataframe(pdf_audit_df, use_container_width=True)

    # ---------- 2.3 可选：达人换货（逐行一件） ----------
    recon_df = None
    stock_delta_df = None
    swap_log_df = None

    if creator_swap_df is not None and not creator_swap_df.empty:
        st.subheader("达人换货（逐行一件）")

        colA, colB = st.columns(2)
        with colA:
            original_col = st.selectbox("选择原SKU列（原款式）", options=list(creator_swap_df.columns))
        with colB:
            new_col = st.selectbox("选择新SKU列（发货款式）", options=list(creator_swap_df.columns))

        if original_col and new_col:
            sold_before = dict(sku_counts_all)

            applied_rows = 0
            missing_in_sold = 0
            log_rows = []

            for _, row in creator_swap_df.iterrows():
                original_sku = str(row[original_col]).strip() if pd.notna(row[original_col]) else ""
                new_sku = str(row[new_col]).strip() if pd.notna(row[new_col]) else ""

                found_in_sold = False

                # Sold 修正：原 -1（仅当当日提取中存在可扣），新 +1（无条件）
                if original_sku:
                    if sku_counts_all.get(original_sku, 0) > 0:
                        sku_counts_all[original_sku] -= 1
                        if sku_counts_all[original_sku] == 0:
                            del sku_counts_all[original_sku]
                        found_in_sold = True
                    else:
                        missing_in_sold += 1

                if new_sku:
                    sku_counts_all[new_sku] += 1

                # 库存修正：原 +1，新 -1（只对库存表中存在的 SKU 行生效）
                if original_sku:
                    stock_df.loc[stock_df["SKU编码"] == original_sku, stock_col] += 1
                if new_sku:
                    stock_df.loc[stock_df["SKU编码"] == new_sku, stock_col] -= 1

                applied_rows += 1
                log_rows.append({
                    "原SKU": original_sku,
                    "新SKU": new_sku,
                    "原SKU是否在当日提取中找到": "是" if found_in_sold else "否"
                })

            swap_log_df = pd.DataFrame(log_rows)

            # 统计 delta
            dec_counts = swap_log_df["原SKU"].value_counts().rename("原SKU减少次数") if not swap_log_df.empty else pd.Series(dtype=int)
            inc_counts = swap_log_df["新SKU"].value_counts().rename("新SKU增加次数") if not swap_log_df.empty else pd.Series(dtype=int)

            delta_sold = pd.concat([
                -dec_counts.rename("Delta"),
                inc_counts.rename("Delta")
            ], axis=0).groupby(level=0).sum().sort_index()

            idx = sorted(delta_sold.index.tolist())
            recon_df = pd.DataFrame({
                "Before Sold": [sold_before.get(k, 0) for k in idx],
                "Delta from Swap": [int(delta_sold.get(k, 0)) for k in idx],
                "After Sold": [sku_counts_all.get(k, 0) for k in idx],
            }, index=idx)
            recon_df["OK?"] = recon_df["After Sold"] == (recon_df["Before Sold"] + recon_df["Delta from Swap"])

            stock_delta = pd.concat([
                dec_counts.rename("Stock Delta(原+1)"),
                (-inc_counts).rename("Stock Delta(新-1)")
            ], axis=0).groupby(level=0).sum().sort_values(ascending=False)
            stock_delta_df = stock_delta.to_frame(name="预期库存变动量（+原 / −新）")

            msg_tail = "" if missing_in_sold == 0 else f"；其中 {missing_in_sold} 行原SKU未在当日提取中找到（Sold 无法逐件 −1，但库存仍会按规则 +1，新SKU 仍会 +1）"
            st.success(f"达人换货处理完成：共应用 {applied_rows} 行{msg_tail}")

            c1, c2 = st.columns(2)
            with c1:
                st.caption("Sold 变动对账（理论 Delta vs 应用前/后）")
                st.dataframe(recon_df, use_container_width=True)
            with c2:
                st.caption("库存预期变动（按达人换货累计）")
                st.dataframe(stock_delta_df, use_container_width=True)

            st.caption("达人换货明细（前100行）")
            st.dataframe(swap_log_df.head(100), use_container_width=True)

    # ---------- 2.4 更新库存 ----------
    stock_df["Sold"] = stock_df["SKU编码"].map(sku_counts_all).fillna(0).astype(int)
    stock_df["New Stock"] = stock_df[stock_col] - stock_df["Sold"]

    summary_df = stock_df[["SKU编码", stock_col, "Sold", "New Stock"]].copy()
    summary_df.columns = ["SKU", "Old Stock", "Sold Qty", "New Stock"]
    summary_df.index += 1
    summary_df.loc["合计"] = ["—", int(summary_df["Old Stock"].sum()), int(summary_df["Sold Qty"].sum()), int(summary_df["New Stock"].sum())]

    st.subheader("库存更新结果")
    st.dataframe(summary_df, use_container_width=True)

    total_sold = int(summary_df.loc["合计", "Sold Qty"])
    st.info(f"本次提取 Sold 总件数：{total_sold}")

    # ---------- 2.5 一键复制 New Stock ----------
    st.subheader("一键复制 New Stock")
    new_stock_text = "\n".join(summary_df.iloc[:-1]["New Stock"].astype(int).astype(str).tolist())
    st.code(new_stock_text, language="text")

    # ---------- 2.6 下载 Excel（多 sheet） ----------
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pdf_audit_df.to_excel(writer, sheet_name="PDF_Audit", index=False)
        summary_df.to_excel(writer, sheet_name="Inventory_Update", index_label="序号")
        if recon_df is not None:
            recon_df.reset_index().rename(columns={"index": "SKU"}).to_excel(writer, sheet_name="Swap_Recon", index=False)
        if stock_delta_df is not None:
            stock_delta_df.reset_index().rename(columns={"index": "SKU"}).to_excel(writer, sheet_name="Swap_Stock_Delta", index=False)
        if swap_log_df is not None:
            swap_log_df.to_excel(writer, sheet_name="Swap_Log", index=False)

    st.download_button(
        label="下载结果 Excel",
        data=output.getvalue(),
        file_name="库存更新结果_独立站Slip.xlsx"
    )

    # ---------- 2.7 上传历史记录 ----------
    history_file = "upload_history.csv"
    new_record = {
        "时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "PDF文件": "; ".join([f.name for f in selected_pdfs]),
        "库存文件": stock_file.name,
        "库存列": stock_col,
        "提取出货数量（Sold Units）": total_sold
    }
    if creator_swap_df is not None:
        new_record["达人换货行数"] = int(len(creator_swap_df))
    else:
        new_record["达人换货行数"] = ""

    if os.path.exists(history_file):
        history_df = pd.read_csv(history_file)
        history_df = pd.concat([history_df, pd.DataFrame([new_record])], ignore_index=True)
    else:
        history_df = pd.DataFrame([new_record])
    history_df.to_csv(history_file, index=False)

    st.subheader("上传历史记录")
    st.dataframe(history_df, use_container_width=True)

else:
    st.info("请上传一个或多个独立站 Slip PDF + 库存 CSV 以开始处理。")
