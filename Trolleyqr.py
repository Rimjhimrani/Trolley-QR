import streamlit as st
import pandas as pd
import qrcode
import io
import zipfile
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from PIL import Image

st.set_page_config(page_title="Trolley QR System", page_icon="🏭", layout="wide")

# ── Helpers ────────────────────────────────────────────────────────────────────
def df_to_excel_bytes(trolley_name: str, df: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = str(trolley_name)[:31]          # ← correct sheet name, not Sheet1

    thin   = Side(style="thin", color="BDBDBD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    cols   = list(df.columns)

    # Title
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols))
    tc            = ws.cell(row=1, column=1, value=f"PICK LIST — {trolley_name.upper()}")
    tc.font       = Font(bold=True, size=13, color="FFFFFF")
    tc.fill       = PatternFill("solid", start_color="0D47A1")
    tc.alignment  = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Header
    for ci, col in enumerate(cols, 1):
        c            = ws.cell(row=2, column=ci, value=col)
        c.font       = Font(bold=True, color="FFFFFF", size=10)
        c.fill       = PatternFill("solid", start_color="1A237E")
        c.alignment  = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border     = border
    ws.row_dimensions[2].height = 34

    # Data
    alt = PatternFill("solid", start_color="E8EAF6")
    for ri, (_, row) in enumerate(df.iterrows(), 3):
        for ci, col in enumerate(cols, 1):
            val = row[col]
            try:   val = val.item()
            except: pass
            c           = ws.cell(row=ri, column=ci, value=val)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border    = border
            if ri % 2 == 0:
                c.fill = alt
        ws.row_dimensions[ri].height = 22

    # Column widths
    for ci, col in enumerate(cols, 1):
        max_len = max(len(str(col)),
                      df[col].astype(str).str.len().max() if len(df) else 0)
        ws.column_dimensions[get_column_letter(ci)].width = min(max(max_len + 2, 9), 36)

    ws.freeze_panes = "A3"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def make_qr_image(url: str) -> Image.Image:
    qr = qrcode.QRCode(version=1,
                       error_correction=qrcode.constants.ERROR_CORRECT_H,
                       box_size=10, border=4)
    qr.add_data(url)
    qr.make(fit=True)
    return qr.make_image(fill_color="black", back_color="white").convert("RGB")

def img_to_bytes(img: Image.Image) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()

# ══════════════════════════════════════════════════════════════════════════════
# SCAN MODE  —  phone opens  ?trolley=TROLLEY-A1
# ══════════════════════════════════════════════════════════════════════════════
scanned = st.query_params.get("trolley")

if scanned:
    trolleys = st.session_state.get("trolleys", {})

    if scanned in trolleys:
        df   = trolleys[scanned]
        xl   = df_to_excel_bytes(scanned, df)

        st.markdown(f"## 🚛 {scanned}")
        st.markdown(f"**{len(df)} parts** in this trolley's pick list.")

        st.download_button(
            label    = "📥 Download Pick List Excel",
            data     = xl,
            file_name= f"PickList_{scanned}.xlsx",
            mime     = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type     = "primary",
            use_container_width=True
        )
        st.dataframe(df, use_container_width=True, hide_index=True)

    else:
        st.error(f"Trolley **{scanned}** not found. Please ask the operator to re-upload the file.")

    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
# MAIN  —  operator screen
# ══════════════════════════════════════════════════════════════════════════════
if "trolleys" not in st.session_state:
    st.session_state.trolleys = {}

st.title("🏭 Trolley QR Pick List System")
st.caption("Upload Excel → Generate QR codes → Workers scan → Pick list opens on phone")

# ── Step 1: App URL ───────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("Step 1 — Enter this app's URL")
st.info(
    "The QR code needs to contain this app's web address so the phone knows where to go.\n\n"
    "**Streamlit Cloud:** copy the URL from your browser (e.g. `https://yourname-trolley.streamlit.app`)\n\n"
    "**Same WiFi / office network:** use your computer's IP address (e.g. `http://192.168.1.5:8501`)"
)
app_url = st.text_input(
    "App URL",
    placeholder="https://yourname-trolley.streamlit.app",
).strip().rstrip("/")

# ── Step 2: Upload ────────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("Step 2 — Upload Excel file")
st.markdown(
    "- Each **sheet = one trolley**\n"
    "- The **sheet name** becomes the trolley name on the QR code\n"
    "- Each sheet contains that trolley's full pick list"
)

uploaded = st.file_uploader("Upload .xlsx file", type=["xlsx"])

if uploaded:
    raw = pd.read_excel(uploaded, sheet_name=None, dtype=str)
    loaded = {name.strip(): df.fillna("") for name, df in raw.items()}
    st.session_state.trolleys = loaded
    st.success(f"✅ Loaded {len(loaded)} trolley(s): **{', '.join(loaded.keys())}**")

# ── Step 3: QR Codes ──────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("Step 3 — QR Codes")

if not st.session_state.trolleys:
    st.info("Upload your Excel file above.")
elif not app_url:
    st.warning("Enter the app URL in Step 1 so QR codes work on phones.")
else:
    st.success("QR codes are ready. Print them and stick on trolleys.")

    trolleys = st.session_state.trolleys
    items    = list(trolleys.items())
    n        = len(items)

    # 3 cards per row
    for row_start in range(0, n, 3):
        cols = st.columns(3)
        for ci in range(3):
            idx = row_start + ci
            if idx >= n:
                break
            t_name, df = items[idx]
            scan_url   = f"{app_url}?trolley={t_name}"
            qr_img     = make_qr_image(scan_url)

            with cols[ci]:
                st.markdown(f"#### 🚛 {t_name}")
                st.image(qr_img, use_container_width=True)
                st.caption(f"Scan → opens pick list on phone  |  {len(df)} parts")

                # Download QR
                st.download_button(
                    "⬇️ Download QR Code",
                    data      = img_to_bytes(qr_img),
                    file_name = f"QR_{t_name}.png",
                    mime      = "image/png",
                    use_container_width=True,
                    key       = f"qr_{t_name}"
                )
                # Download Excel directly too
                st.download_button(
                    "📊 Download Excel",
                    data      = df_to_excel_bytes(t_name, df),
                    file_name = f"PickList_{t_name}.xlsx",
                    mime      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key       = f"xl_{t_name}"
                )

    # Bulk ZIP
    st.markdown("---")
    st.subheader("Download all at once")
    c1, c2 = st.columns(2)

    with c1:
        if st.button("Generate ZIP — All QR Codes", use_container_width=True):
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w") as zf:
                for t_name, df in trolleys.items():
                    url = f"{app_url}?trolley={t_name}"
                    zf.writestr(f"QR_{t_name}.png", img_to_bytes(make_qr_image(url)))
            st.download_button("⬇️ Download QR ZIP", data=buf.getvalue(),
                               file_name="All_QR_Codes.zip", mime="application/zip",
                               use_container_width=True, key="zip_qr")

    with c2:
        if st.button("Generate ZIP — All Excel Files", use_container_width=True):
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w") as zf:
                for t_name, df in trolleys.items():
                    zf.writestr(f"PickList_{t_name}.xlsx", df_to_excel_bytes(t_name, df))
            st.download_button("⬇️ Download Excel ZIP", data=buf.getvalue(),
                               file_name="All_PickLists.zip", mime="application/zip",
                               use_container_width=True, key="zip_xl")
