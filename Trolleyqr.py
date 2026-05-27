import streamlit as st
import pandas as pd
import qrcode
import io
import zipfile
import json
import os
import tempfile
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from PIL import Image
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

st.set_page_config(page_title="Trolley QR System", page_icon="🏭", layout="wide")

st.markdown("""
<style>
.header {
    background: linear-gradient(135deg,#1a237e,#1565c0);
    color:white; padding:20px 28px; border-radius:12px;
    text-align:center; margin-bottom:20px;
}
.qr-card {
    background:white; border:2px solid #e0e6ff; border-radius:12px;
    padding:18px; text-align:center;
    box-shadow:0 2px 12px rgba(0,0,0,0.07); margin-bottom:12px;
}
.step {
    background:#f0f4ff; border-left:5px solid #1a237e;
    border-radius:8px; padding:12px 16px; margin:10px 0;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="header">
  <h1 style="margin:0">🏭 Trolley QR Pick List System</h1>
  <p style="margin:6px 0 0;opacity:.85">Upload Excel → QR generated → Scan → Excel opens on phone</p>
</div>
""", unsafe_allow_html=True)

# ── Helpers ────────────────────────────────────────────────────────────────────
def df_to_excel_bytes(trolley_name: str, df: pd.DataFrame) -> bytes:
    wb  = Workbook()
    ws  = wb.active
    ws.title = str(trolley_name)[:31]

    thin   = Side(style="thin", color="BDBDBD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    cols   = list(df.columns)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols))
    tc           = ws.cell(row=1, column=1, value=f"PICK LIST — {trolley_name.upper()}")
    tc.font      = Font(bold=True, size=13, color="FFFFFF")
    tc.fill      = PatternFill("solid", start_color="0D47A1")
    tc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    for ci, col in enumerate(cols, 1):
        c            = ws.cell(row=2, column=ci, value=col)
        c.font       = Font(bold=True, color="FFFFFF", size=10)
        c.fill       = PatternFill("solid", start_color="1A237E")
        c.alignment  = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border     = border
    ws.row_dimensions[2].height = 34

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

    for ci, col in enumerate(cols, 1):
        max_len = max(len(str(col)),
                      df[col].astype(str).str.len().max() if len(df) else 0)
        ws.column_dimensions[get_column_letter(ci)].width = min(max(max_len + 2, 9), 36)

    ws.freeze_panes = "A3"
    buf = io.BytesIO(); wb.save(buf)
    return buf.getvalue()

def make_qr(text: str) -> Image.Image:
    qr = qrcode.QRCode(version=1,
                       error_correction=qrcode.constants.ERROR_CORRECT_H,
                       box_size=10, border=4)
    qr.add_data(text)
    qr.make(fit=True)
    return qr.make_image(fill_color="black", back_color="white").convert("RGB")

def img_to_bytes(img: Image.Image) -> bytes:
    buf = io.BytesIO(); img.save(buf, format="PNG"); return buf.getvalue()

def get_drive_service(creds_dict: dict):
    creds = service_account.Credentials.from_service_account_info(
        creds_dict,
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)

def upload_to_drive(service, file_bytes: bytes, filename: str, folder_id: str | None = None) -> str:
    """Upload file, make public, return shareable link."""
    meta = {"name": filename}
    if folder_id:
        meta["parents"] = [folder_id]

    media = MediaIoBaseUpload(
        io.BytesIO(file_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    f = service.files().create(body=meta, media_body=media, fields="id").execute()
    fid = f["id"]

    # Make publicly readable
    service.permissions().create(
        fileId=fid,
        body={"type": "anyone", "role": "reader"}
    ).execute()

    # Direct download link
    return f"https://drive.google.com/uc?export=download&id={fid}"

# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════
if "trolleys"   not in st.session_state: st.session_state.trolleys   = {}
if "drive_links" not in st.session_state: st.session_state.drive_links = {}  # {trolley: url}
if "qr_images"  not in st.session_state: st.session_state.qr_images  = {}   # {trolley: bytes}

# ══════════════════════════════════════════════════════════════════════════════
# STEP 1 — Google Service Account credentials
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("## Step 1 — Connect Google Drive")

with st.expander("📋 How to get your Google Service Account key (one-time setup)", expanded=False):
    st.markdown("""
1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Create a project (or use existing)
3. Enable **Google Drive API**
4. Go to **IAM & Admin → Service Accounts** → Create service account
5. Click the service account → **Keys** tab → **Add Key → JSON**
6. Download the JSON file — upload it below
7. Done! This is a one-time setup.
    """)

creds_file = st.file_uploader("Upload Google Service Account JSON key", type=["json"])
creds_dict = None
drive_service = None

if creds_file:
    try:
        creds_dict = json.load(creds_file)
        drive_service = get_drive_service(creds_dict)
        st.success("✅ Google Drive connected!")
    except Exception as e:
        st.error(f"Invalid credentials: {e}")

# ══════════════════════════════════════════════════════════════════════════════
# STEP 2 — Upload Excel
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("---")
st.markdown("## Step 2 — Upload Excel File")
st.markdown('<div class="step">Each <b>sheet name</b> = one Trolley name. Each sheet = that trolley\'s pick list.</div>', unsafe_allow_html=True)

uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

if uploaded:
    raw = pd.read_excel(uploaded, sheet_name=None, dtype=str)
    st.session_state.trolleys = {
        name.strip(): df.fillna("") for name, df in raw.items()
    }
    # Reset previous links when new file uploaded
    st.session_state.drive_links = {}
    st.session_state.qr_images   = {}
    names = list(st.session_state.trolleys.keys())
    st.success(f"✅ {len(names)} trolley(s) found: **{', '.join(names)}**")

# ══════════════════════════════════════════════════════════════════════════════
# STEP 3 — Generate & Upload
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("---")
st.markdown("## Step 3 — Generate QR Codes")

if not st.session_state.trolleys:
    st.info("Upload your Excel file in Step 2.")
elif not drive_service:
    st.warning("Connect Google Drive in Step 1 so QR codes work on phones.")
else:
    if st.button("🚀 Upload to Google Drive & Generate QR Codes", type="primary", use_container_width=True):
        progress = st.progress(0, text="Starting...")
        trolleys = st.session_state.trolleys
        n        = len(trolleys)

        for i, (t_name, df) in enumerate(trolleys.items()):
            progress.progress((i) / n, text=f"Uploading {t_name}...")
            try:
                xl_bytes = df_to_excel_bytes(t_name, df)
                link     = upload_to_drive(drive_service, xl_bytes, f"PickList_{t_name}.xlsx")
                qr_img   = make_qr(link)
                st.session_state.drive_links[t_name] = link
                st.session_state.qr_images[t_name]   = img_to_bytes(qr_img)
            except Exception as e:
                st.error(f"Failed for {t_name}: {e}")

        progress.progress(1.0, text="Done!")
        st.success("✅ All trolleys uploaded! QR codes are ready below.")
        st.rerun()

# ── Show QR cards ─────────────────────────────────────────────────────────────
if st.session_state.qr_images:
    st.markdown("### 📱 Print & stick these QR codes on trolleys")
    st.caption("Worker scans QR → phone opens Google Drive → Excel file downloads automatically")

    items = list(st.session_state.qr_images.items())
    n     = len(items)

    for row_start in range(0, n, 3):
        cols = st.columns(3)
        for ci in range(3):
            idx = row_start + ci
            if idx >= n: break
            t_name, qr_bytes = items[idx]

            with cols[ci]:
                st.markdown('<div class="qr-card">', unsafe_allow_html=True)
                st.image(qr_bytes, use_container_width=True)
                st.markdown(f"**🚛 {t_name}**")
                df = st.session_state.trolleys[t_name]
                st.caption(f"{len(df)} parts")

                st.download_button(
                    "⬇️ Download QR Code (PNG)",
                    data       = qr_bytes,
                    file_name  = f"QR_{t_name}.png",
                    mime       = "image/png",
                    use_container_width=True,
                    key        = f"qr_{t_name}"
                )
                # Also allow direct Excel download
                xl = df_to_excel_bytes(t_name, df)
                st.download_button(
                    "📊 Download Excel",
                    data       = xl,
                    file_name  = f"PickList_{t_name}.xlsx",
                    mime       = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key        = f"xl_{t_name}"
                )
                st.markdown("</div>", unsafe_allow_html=True)

    # Bulk ZIP of all QR codes
    st.markdown("---")
    if st.button("🗜️ Download ALL QR Codes as ZIP", use_container_width=True):
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as zf:
            for t_name, qr_bytes in st.session_state.qr_images.items():
                zf.writestr(f"QR_{t_name}.png", qr_bytes)
        st.download_button(
            "⬇️ Download QR ZIP",
            data=buf.getvalue(),
            file_name="All_QR_Codes.zip",
            mime="application/zip",
            use_container_width=True,
            key="zip_all"
        )
