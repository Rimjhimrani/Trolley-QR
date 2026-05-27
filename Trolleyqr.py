import streamlit as st
import pandas as pd
import qrcode
import io
import zipfile
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from PIL import Image

st.set_page_config(page_title="Trolley QR Pick List", page_icon="🏭", layout="wide")

st.markdown("""
<style>
.main-header {
    background: linear-gradient(135deg, #1a237e, #1565c0);
    color: white; padding: 22px 32px; border-radius: 12px;
    margin-bottom: 24px; text-align: center;
}
.qr-card {
    background: white; border: 2px solid #e3e8f0;
    border-radius: 14px; padding: 20px; text-align: center;
    box-shadow: 0 4px 16px rgba(0,0,0,0.08);
}
.step-box {
    background: #f0f4ff; border-left: 5px solid #1a237e;
    border-radius: 8px; padding: 14px 18px; margin: 8px 0;
    font-size: 15px;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
  <h1 style="margin:0;font-size:2rem">🏭 Trolley QR Pick List System</h1>
  <p style="margin:8px 0 0;opacity:.85">Upload Excel → Auto-generate QR codes → Scan to download full pick list</p>
</div>
""", unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────────────────────
if "trolleys" not in st.session_state:
    st.session_state.trolleys = {}   # {trolley_name: DataFrame}

# ── Helpers ───────────────────────────────────────────────────────────────────
def safe_sum(series):
    total = 0
    for v in series:
        try:
            total += float(v)
        except (ValueError, TypeError):
            pass
    return int(total)

def make_qr(text: str, box_size: int = 8) -> Image.Image:
    qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_H,
                       box_size=box_size, border=4)
    qr.add_data(text)
    qr.make(fit=True)
    return qr.make_image(fill_color="black", back_color="white").convert("RGB")

def img_to_bytes(img: Image.Image) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()

def df_to_excel_bytes(trolley_name: str, df: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = str(trolley_name)[:31]

    thin = Side(style="thin", color="BDBDBD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_fill = PatternFill("solid", start_color="1A237E")
    alt_fill  = PatternFill("solid", start_color="E8EAF6")
    title_fill = PatternFill("solid", start_color="0D47A1")

    cols = list(df.columns)
    num_cols = len(cols)

    # Title
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
    tc = ws.cell(row=1, column=1, value=f"PICK LIST — TROLLEY: {trolley_name.upper()}")
    tc.font = Font(bold=True, size=13, color="FFFFFF")
    tc.fill = title_fill
    tc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Header
    for ci, col in enumerate(cols, 1):
        c = ws.cell(row=2, column=ci, value=col)
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.fill = hdr_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
    ws.row_dimensions[2].height = 34

    # Data
    for ri, (_, row) in enumerate(df.iterrows(), 3):
        for ci, col in enumerate(cols, 1):
            val = row[col]
            # Convert numpy types to python native
            try:
                val = val.item()
            except AttributeError:
                pass
            c = ws.cell(row=ri, column=ci, value=val)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = border
            if ri % 2 == 0:
                c.fill = alt_fill
        ws.row_dimensions[ri].height = 22

    # Auto column widths (capped)
    for ci, col in enumerate(cols, 1):
        max_len = max(
            len(str(col)),
            df[col].astype(str).str.len().max() if len(df) > 0 else 0
        )
        ws.column_dimensions[get_column_letter(ci)].width = min(max(max_len + 2, 8), 35)

    ws.freeze_panes = "A3"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ══════════════════════════════════════════════════════════════════════════════
# STEP 1 — Upload
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("## Step 1 — Upload Your Excel File")
st.markdown('<div class="step-box">📂 Each <b>sheet name</b> in the Excel file = one <b>Trolley name</b>. Each sheet contains that trolley\'s full pick list.</div>', unsafe_allow_html=True)

uploaded = st.file_uploader("Upload Excel (.xlsx) — multiple sheets supported", type=["xlsx"])

if uploaded:
    try:
        all_sheets = pd.read_excel(uploaded, sheet_name=None, dtype=str)
        loaded = {}
        for sheet_name, df in all_sheets.items():
            df = df.fillna("")
            df.columns = [str(c).strip() for c in df.columns]
            loaded[sheet_name.strip()] = df
        st.session_state.trolleys = loaded
        st.success(f"✅ Loaded **{len(loaded)}** trolley(s): {', '.join(loaded.keys())}")
    except Exception as e:
        st.error(f"Error reading file: {e}")

# show sample format hint
with st.expander("ℹ️ Expected Excel format (click to see)"):
    sample = pd.DataFrame([{
        "S.No": 1, "Part No": "P001", "Description": "Bolt M8x20",
        "Store": "ST-01", "Zone": "Z1", "Rack": "R2",
        "Picking Location (old Location)": "L-12",
        "Trolley Location": "T-A1-01", "Qty": 50,
        "Pick Qty": 0, "Pending Qty": 50,
        "Delivery Location": "LINE-1", "Family": "Fasteners"
    }])
    st.dataframe(sample, hide_index=True, use_container_width=True)
    st.caption("Sheet name = Trolley name (e.g. sheet named 'TROLLEY-A1' → trolley A1)")

st.markdown("---")

# ══════════════════════════════════════════════════════════════════════════════
# STEP 2 — QR Codes
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("## Step 2 — Generated QR Codes")

if not st.session_state.trolleys:
    st.info("Upload an Excel file above to generate QR codes.")
else:
    st.markdown('<div class="step-box">🖨️ Print each QR code and stick it on the corresponding physical trolley. Scanning downloads the full pick list Excel sheet.</div>', unsafe_allow_html=True)

    n = len(st.session_state.trolleys)
    cols_per_row = min(n, 3)
    trolley_items = list(st.session_state.trolleys.items())

    for row_start in range(0, n, cols_per_row):
        cols = st.columns(cols_per_row)
        for col_idx in range(cols_per_row):
            item_idx = row_start + col_idx
            if item_idx >= n:
                break
            t_name, df = trolley_items[item_idx]
            with cols[col_idx]:
                st.markdown('<div class="qr-card">', unsafe_allow_html=True)

                # QR encodes the trolley name
                qr_img = make_qr(t_name, box_size=7)
                st.image(qr_img, use_container_width=True)
                st.markdown(f"**🚛 {t_name}**")
                st.caption(f"{len(df)} parts")

                # Download QR PNG
                st.download_button(
                    "⬇️ QR Code (PNG)",
                    data=img_to_bytes(qr_img),
                    file_name=f"QR_{t_name}.png",
                    mime="image/png",
                    use_container_width=True,
                    key=f"qr_{t_name}"
                )
                # Download full Excel
                xl = df_to_excel_bytes(t_name, df)
                st.download_button(
                    "📥 Pick List (Excel)",
                    data=xl,
                    file_name=f"PickList_{t_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key=f"xl_{t_name}"
                )
                st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("---")

    # ── Bulk downloads ────────────────────────────────────────────────────────
    st.markdown("## Step 3 — Bulk Download (Optional)")
    c1, c2 = st.columns(2)

    with c1:
        st.markdown("#### 📱 All QR Codes")
        if st.button("Generate QR ZIP", use_container_width=True):
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for t_name, df in st.session_state.trolleys.items():
                    zf.writestr(f"QR_{t_name}.png", img_to_bytes(make_qr(t_name, box_size=10)))
            st.download_button(
                "⬇️ Download All QR Codes (.zip)",
                data=buf.getvalue(),
                file_name="All_QR_Codes.zip",
                mime="application/zip",
                use_container_width=True,
                key="zip_qr"
            )

    with c2:
        st.markdown("#### 📊 All Pick Lists")
        if st.button("Generate Excel ZIP", use_container_width=True):
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for t_name, df in st.session_state.trolleys.items():
                    zf.writestr(f"PickList_{t_name}.xlsx", df_to_excel_bytes(t_name, df))
            st.download_button(
                "⬇️ Download All Pick Lists (.zip)",
                data=buf.getvalue(),
                file_name="All_PickLists.zip",
                mime="application/zip",
                use_container_width=True,
                key="zip_xl"
            )

    # ── Scan simulation ───────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("## 🔍 Scan Simulator (Preview)")
    st.caption("This simulates what happens when a worker scans a QR code — they see and download the full pick list.")

    selected = st.selectbox("Select trolley (simulates QR scan)", options=list(st.session_state.trolleys.keys()))
    if selected:
        df_sel = st.session_state.trolleys[selected]
        st.markdown(f"### 📋 Full Pick List — {selected}")
        st.dataframe(df_sel, use_container_width=True, hide_index=True, height=400)
        xl = df_to_excel_bytes(selected, df_sel)
        st.download_button(
            f"📥 Download {selected} Pick List",
            data=xl,
            file_name=f"PickList_{selected}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
