import streamlit as st
import pandas as pd
import re
from collections import defaultdict
from io import BytesIO
from datetime import datetime
import os

# -----------------------------
# Optional dependency: PyMuPDF
# -----------------------------
try:
    import fitz  # PyMuPDF
except Exception:
    st.set_page_config(page_title="NailVesta 库存系统（独立站 Slip PDF）", layout="centered")
    st.error("缺少依赖 PyMuPDF（fitz）。请在 requirements.txt 中加入 PyMuPDF 并重新部署 / 重启应用。")
    st.stop()

# =========================
# 0) Page + Style (Pink theme)
# =========================
st.set_page_config(page_title="NailVesta 库存系统（独立站 Slip PDF）", layout="wide")

PINK_CSS = """
<style>
:root{
  --nv-pink-50:#fff1f6;
  --nv-pink-100:#ffe4ee;
  --nv-pink-200:#ffc7da;
  --nv-pink-300:#ff9dbf;
  --nv-pink-400:#ff73a4;
  --nv-pink-500:#ff4d8e;
  --nv-pink-600:#e63b7a;
  --nv-text:#2a1b22;
  --nv-muted:#6b5b63;
  --nv-card:#ffffff;
  --nv-border:rgba(255,77,142,.18);
  --nv-shadow:0 10px 30px rgba(255,77,142,.10);
}
.stApp{
  background:
    radial-gradient(1200px 600px at 15% 10%, rgba(255,77,142,.14), transparent 55%),
    radial-gradient(900px 500px at 85% 20%, rgba(255,157,191,.18), transparent 55%),
    linear-gradient(180deg, var(--nv-pink-50), #ffffff 55%);
  color: var(--nv-text);
}
.block-container{ padding-top: 1.2rem; padding-bottom: 2rem; }
h1, h2, h3{ color: var(--nv-text) !important; letter-spacing: .2px; }
h1{ font-weight: 800; } h2{ font-weight: 750; } h3{ font-weight: 700; }

.nv-card{
  background: var(--nv-card);
  border: 1px solid var(--nv-border);
  box-shadow: var(--nv-shadow);
  border-radius: 18px;
  padding: 16px 16px;
}
.nv-banner{
  display:flex; align-items:center; justify-content:space-between; gap:14px;
  padding: 18px 18px;
  border-radius: 22px;
  background:
    linear-gradient(135deg, rgba(255,77,142,.16), rgba(255,157,191,.16)),
    linear-gradient(180deg, #fff, #fff);
  border: 1px solid var(--nv-border);
  box-shadow: var(--nv-shadow);
}
.nv-banner-title{ font-size: 26px; font-weight: 850; margin: 0; }
.nv-banner-sub{ margin: 4px 0 0 0; color: var(--nv-muted); font-size: 13.5px; line-height: 1.45; }
.nv-badge{
  display:inline-flex; align-items:center; gap:10px;
  padding: 10px 12px;
  border-radius: 999px;
  background: rgba(255,77,142,.10);
  border: 1px solid rgba(255,77,142,.20);
  color: var(--nv-text);
  font-weight: 650;
  font-size: 13px;
  white-space: nowrap;
}
.nv-mini{ color: var(--nv-muted); font-size: 12.5px; }

.stButton button, .stDownloadButton button{
  background: linear-gradient(180deg, var(--nv-pink-500), var(--nv-pink-600)) !important;
  color: white !important;
  border: 0 !important;
  border-radius: 12px !important;
  padding: 0.65rem 1rem !important;
  font-weight: 700 !important;
  box-shadow: 0 8px 18px rgba(255,77,142,.25) !important;
}
.stButton button:hover, .stDownloadButton button:hover{ filter: brightness(1.03); transform: translateY(-1px); }
.stButton button:active, .stDownloadButton button:active{ transform: translateY(0px); }

div[data-testid="stFileUploader"] section{
  border: 1px dashed rgba(255,77,142,.35) !important;
  border-radius: 14px !important;
  background: rgba(255,241,246,.55) !important;
}
div[data-testid="stSelectbox"] > div{ border-radius: 12px !important; }
div[data-testid="stMultiSelect"] > div{ border-radius: 12px !important; }

div[data-testid="stDataFrame"]{
  border-radius: 14px;
  overflow: hidden;
  border: 1px solid rgba(255,77,142,.16);
}
div[data-testid="stAlert"]{
  border-radius: 14px;
  border: 1px solid rgba(255,77,142,.18);
  box-shadow: 0 8px 18px rgba(255,77,142,.08);
}
pre{
  border-radius: 14px !important;
  border: 1px solid rgba(255,77,142,.14) !important;
}
</style>
"""
st.markdown(PINK_CSS, unsafe_allow_html=True)

BOW_SVG = """
<svg width="56" height="36" viewBox="0 0 56 36" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <linearGradient id="g" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0" stop-color="#ff73a4"/>
      <stop offset="1" stop-color="#ff4d8e"/>
    </linearGradient>
  </defs>
  <path d="M18 18c-8-10-16-10-16-3 0 7 8 10 16 3z" fill="url(#g)" opacity="0.92"/>
  <path d="M38 18c8-10 16-10 16-3 0 7-8 10-16 3z" fill="url(#g)" opacity="0.92"/>
  <ellipse cx="28" cy="18" rx="7.5" ry="6.3" fill="#ffd1e3" stroke="#ff4d8e" stroke-width="1.2"/>
  <circle cx="28" cy="18" r="2.3" fill="#ff4d8e"/>
</svg>
"""

st.markdown(
    f"""
<div class="nv-banner">
  <div>
    <div class="nv-banner-title">ColorFour Inventory（Slip PDF 版本）</div>
    <div class="nv-banner-sub">
      流程：Slip PDF 提取 Sold → 映射库存（SKU编码）→ 计算 New Stock → 导出 Excel
    </div>
  </div>
  <div class="nv-badge">
    {BOW_SVG}
    <span>Pink Mode</span>
  </div>
</div>
""",
    unsafe_allow_html=True
)
st.write("")

# =========================
# 1) Uploaders
# =========================
colL, colR = st.columns([1.2, 1.0], gap="large")
with colL:
    st.markdown('<div class="nv-card">', unsafe_allow_html=True)
    st.subheader("上传文件")
    pdf_files = st.file_uploader("上传独立站 Slip/Invoice PDF（可多选）", type=["pdf"], accept_multiple_files=True)
    stock_file = st.file_uploader("上传库存表 CSV（包含 SKU编码 + 库存列）", type=["csv"])
    st.markdown('<div class="nv-mini">提示：Slip PDF 为逐件明细口径，不需要 Item quantity / bundle / NM001。</div>', unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

with colR:
    st.markdown('<div class="nv-card">', unsafe_allow_html=True)
    st.subheader("设置")
    st.markdown('<div class="nv-mini">库存列在处理阶段选择（保持原始 CSV 行顺序输出）。</div>', unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

selected_pdfs = []
if pdf_files:
    selected_names = st.multiselect(
        "选择要参与统计的 Slip PDF（默认全选）",
        options=[f.name for f in pdf_files],
        default=[f.name for f in pdf_files],
    )
    selected_pdfs = [f for f in pdf_files if f.name in selected_names]

# =========================
# 2) PDF parsing (Slip)
# =========================
SKU_RE = re.compile(r"\b([A-Z]{3}\d{3}-[SML])\b")
ORDER_RE = re.compile(r"\bOrder\s*#\s*([A-Za-z0-9_-]+)\b", re.I)
OF_RE = re.compile(r"\b(\d+)\s+of\s+(\d+)\b", re.I)

def normalize_text(t: str) -> str:
    t = t.replace("\u00ad", "").replace("\u200b", "").replace("\u00a0", " ")
    t = t.replace("–", "-").replace("—", "-")
    return t

def parse_slip_pdf(file_bytes: bytes):
    """
    Slip PDF（逐件明细）提取：
      - SKU 形态：ABC123-S/M/L
      - 默认 qty=1；若 SKU 附近出现 '2 of 2' 这种格式则取前面的数量
    """
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    pages_text = [p.get_text("text") for p in doc]
    text = normalize_text("\n".join(pages_text))

    order_ids = set(ORDER_RE.findall(text))
    sku_counts = defaultdict(int)

    for m in SKU_RE.finditer(text):
        sku = m.group(1).strip()
        after = text[m.end(): m.end() + 80]
        mm = OF_RE.search(after)
        qty = int(mm.group(1)) if mm else 1
        sku_counts[sku] += qty

    units = int(sum(sku_counts.values()))
    return sku_counts, units, order_ids

# =========================
# 3) Main
# =========================
if selected_pdfs and stock_file:
    st.markdown('<div class="nv-card">', unsafe_allow_html=True)
    st.success("文件已上传，开始处理。")
    st.markdown("</div>", unsafe_allow_html=True)

    # -------- 3.1 Load stock (KEEP ORIGINAL ORDER) --------
    stock_df = pd.read_csv(stock_file)
    stock_df.columns = [str(c).strip() for c in stock_df.columns]

    if "SKU编码" not in stock_df.columns:
        st.error("库存表缺少列：SKU编码")
        st.stop()

    # ✅ 记录原始顺序（最终展示按此顺序）
    stock_df["_row_order"] = range(len(stock_df))

    # 🚫 不允许重复 SKU：如果有，直接提示并停止（避免 Sold 被重复扣减）
    dup = stock_df["SKU编码"].duplicated(keep=False)
    if dup.any():
        dup_list = (
            stock_df.loc[dup, "SKU编码"]
            .astype(str)
            .value_counts()
            .head(50)
            .to_dict()
        )
        st.error("检测到库存表存在重复 SKU编码，已停止运行（为避免重复扣减库存）。")
        st.caption("请先修正库存表，使 SKU编码 唯一。下面是前 50 个重复 SKU（SKU: 重复次数）：")
        st.json(dup_list)
        st.stop()

    # pick stock column
    candidate_cols = [c for c in stock_df.columns if c not in ["SKU编码", "_row_order"]]
    if not candidate_cols:
        st.error("库存表没有可用的库存列（除 SKU编码 外至少需要一列）")
        st.stop()

    common_stock_names = ["Stock", "库存", "Current Stock", "Old Stock", "现有库存"]
    default_stock_col = None
    for c in common_stock_names:
        if c in stock_df.columns:
            default_stock_col = c
            break

    stock_col = st.selectbox(
        "选择库存列（Old Stock）",
        options=candidate_cols,
        index=(candidate_cols.index(default_stock_col) if default_stock_col in candidate_cols else 0),
    )

    stock_df[stock_col] = pd.to_numeric(stock_df[stock_col], errors="coerce").fillna(0).astype(int)

    # -------- 3.2 Parse PDFs --------
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
        pdf_audit_df = pd.concat(
            [pdf_audit_df, pd.DataFrame([{
                "PDF文件": "合计",
                "识别订单数（Order #）": int(pdf_audit_df["识别订单数（Order #）"].sum()),
                "提取SKU种类数": int(pdf_audit_df["提取SKU种类数"].sum()),
                "提取件数（Sold Units）": int(pdf_audit_df["提取件数（Sold Units）"].sum()),
            }])],
            ignore_index=True
        )

    st.subheader("PDF 提取对账（独立站 Slip：逐件明细口径）")
    st.dataframe(pdf_audit_df, use_container_width=True)

    # -------- 3.3 Update inventory (KEEP ORIGINAL ORDER) --------
    stock_df["Sold"] = stock_df["SKU编码"].map(sku_counts_all).fillna(0).astype(int)
    stock_df["New Stock"] = stock_df[stock_col] - stock_df["Sold"]

    # ✅ 按原始顺序输出（虽然没排序过，但强制更安心）
    stock_df = stock_df.sort_values("_row_order", kind="stable").reset_index(drop=True)

    summary_df = stock_df[["SKU编码", stock_col, "Sold", "New Stock"]].copy()
    summary_df.columns = ["SKU", "Old Stock", "Sold Qty", "New Stock"]
    summary_df.index += 1

    summary_df.loc["合计"] = [
        "—",
        int(summary_df["Old Stock"].sum()),
        int(summary_df["Sold Qty"].sum()),
        int(summary_df["New Stock"].sum()),
    ]

    st.subheader("库存更新结果（保持库存 CSV 原始顺序）")
    st.dataframe(summary_df, use_container_width=True)

    total_sold_from_pdf = int(sum(sku_counts_all.values()))
    st.markdown('<div class="nv-card">', unsafe_allow_html=True)
    st.markdown(
        f"<div style='font-size:16px;font-weight:750;color:var(--nv-text);'>本次提取 Sold 总件数：{total_sold_from_pdf}</div>",
        unsafe_allow_html=True
    )
    st.markdown(
        "<div class='nv-mini'>Sold 总件数以 PDF 提取明细为准；如库存表存在重复 SKU，本程序会直接停止并提示。</div>",
        unsafe_allow_html=True
    )
    st.markdown("</div>", unsafe_allow_html=True)

    # -------- 3.4 Copy New Stock --------
    st.subheader("一键复制 New Stock（按库存 CSV 顺序）")
    new_stock_text = "\n".join(summary_df.iloc[:-1]["New Stock"].astype(int).astype(str).tolist())
    st.code(new_stock_text, language="text")

    # -------- 3.5 Export Excel --------
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pdf_audit_df.to_excel(writer, sheet_name="PDF_Audit", index=False)
        summary_df.to_excel(writer, sheet_name="Inventory_Update", index_label="序号")

    st.download_button(
        label="下载结果 Excel",
        data=output.getvalue(),
        file_name="库存更新结果_独立站Slip_Pink.xlsx"
    )

    # -------- 3.6 Upload history --------
    history_file = "upload_history.csv"
    new_record = {
        "时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "PDF文件": "; ".join([f.name for f in selected_pdfs]),
        "库存文件": stock_file.name,
        "库存列": stock_col,
        "提取出货数量（Sold Units）": total_sold_from_pdf
    }

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
