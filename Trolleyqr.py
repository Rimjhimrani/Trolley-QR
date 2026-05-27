import streamlit as st
import pandas as pd
import qrcode
import io
import os
import json
import zipfile
import base64
from PIL import Image, ImageDraw, ImageFont
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─── Page Config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Trolley QR Manager",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS Styling ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Rajdhani:wght@500;600;700&family=IBM+Plex+Mono:wght@400;600&display=swap');

:root {
    --orange: #FF6B2B;
    --dark: #0D0D0D;
    --card: #161616;
    --border: #2A2A2A;
    --muted: #888;
    --text: #E8E8E8;
    --green: #39D353;
    --yellow: #F5C518;
}

html, body, [class*="css"] {
    font-family: 'Rajdhani', sans-serif;
    background-color: var(--dark) !important;
    color: var(--text);
}

.stApp { background: var(--dark); }

/* Header */
.main-header {
    background: linear-gradient(135deg, #1a1a1a 0%, #0d0d0d 100%);
    border: 1px solid var(--orange);
    border-radius: 4px;
    padding: 24px 32px;
    margin-bottom: 28px;
    display: flex;
    align-items: center;
    gap: 16px;
    box-shadow: 0 0 40px rgba(255,107,43,0.15);
}

.main-header h1 {
    font-family: 'Rajdhani', sans-serif;
    font-size: 2.2rem;
    font-weight: 700;
    color: var(--orange);
    margin: 0;
    letter-spacing: 3px;
    text-transform: uppercase;
}

.main-header p {
    color: var(--muted);
    margin: 0;
    font-size: 0.9rem;
    letter-spacing: 1px;
    font-family: 'IBM Plex Mono', monospace;
}

/* Cards */
.trolley-card {
    background: var(--card);
    border: 1px solid var(--border);
    border-left: 3px solid var(--orange);
    border-radius: 4px;
    padding: 16px 20px;
    margin-bottom: 12px;
    display: flex;
    justify-content: space-between;
    align-items: center;
}

.trolley-name {
    font-size: 1.1rem;
    font-weight: 700;
    color: var(--orange);
    letter-spacing: 2px;
    text-transform: uppercase;
}

.trolley-meta {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.75rem;
    color: var(--muted);
}

.badge {
    background: rgba(255,107,43,0.15);
    border: 1px solid var(--orange);
    color: var(--orange);
    padding: 3px 10px;
    border-radius: 2px;
    font-size: 0.75rem;
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 600;
}

.badge-green {
    background: rgba(57,211,83,0.15);
    border: 1px solid var(--green);
    color: var(--green);
}

.stat-box {
    background: var(--card);
    border: 1px solid var(--border);
    border-top: 2px solid var(--orange);
    border-radius: 4px;
    padding: 16px;
    text-align: center;
}

.stat-number {
    font-size: 2rem;
    font-weight: 700;
    color: var(--orange);
    font-family: 'IBM Plex Mono', monospace;
}

.stat-label {
    color: var(--muted);
    font-size: 0.8rem;
    letter-spacing: 1px;
    text-transform: uppercase;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background: #111 !important;
    border-right: 1px solid var(--border);
}

/* Buttons */
.stButton > button {
    background: var(--orange) !important;
    color: #000 !important;
    font-family: 'Rajdhani', sans-serif !important;
    font-weight: 700 !important;
    letter-spacing: 2px !important;
    text-transform: uppercase !important;
    border: none !important;
    border-radius: 2px !important;
    padding: 10px 24px !important;
}

.stButton > button:hover {
    background: #ff8c5a !important;
    transform: translateY(-1px);
}

/* Download button */
.stDownloadButton > button {
    background: transparent !important;
    color: var(--green) !important;
    border: 1px solid var(--green) !important;
    font-family: 'Rajdhani', sans-serif !important;
    font-weight: 600 !important;
    letter-spacing: 1px !important;
    border-radius: 2px !important;
}

/* Inputs */
.stTextInput > div > div > input,
.stSelectbox > div > div {
    background: var(--card) !important;
    border: 1px solid var(--border) !important;
    color: var(--text) !important;
    border-radius: 2px !important;
    font-family: 'IBM Plex Mono', monospace !important;
}

/* File uploader */
.stFileUploader {
    background: var(--card);
    border: 1px dashed var(--border);
    border-radius: 4px;
}

/* Tabs */
.stTabs [data-baseweb="tab-list"] {
    background: #111;
    border-bottom: 1px solid var(--border);
    gap: 0;
}

.stTabs [data-baseweb="tab"] {
    font-family: 'Rajdhani', sans-serif;
    font-weight: 600;
    letter-spacing: 2px;
    text-transform: uppercase;
    color: var(--muted) !important;
    background: transparent !important;
    border-bottom: 2px solid transparent !important;
    padding: 12px 24px;
}

.stTabs [aria-selected="true"] {
    color: var(--orange) !important;
    border-bottom: 2px solid var(--orange) !important;
}

/* Dataframe */
.stDataFrame {
    border: 1px solid var(--border) !important;
}

/* Divider */
hr { border-color: var(--border) !important; }

/* Alerts */
.stSuccess { background: rgba(57,211,83,0.1) !important; border: 1px solid var(--green) !important; }
.stInfo { background: rgba(255,107,43,0.08) !important; border: 1px solid var(--orange) !important; }
.stWarning { background: rgba(245,197,24,0.1) !important; border: 1px solid var(--yellow) !important; }

/* Step labels */
.step-label {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.75rem;
    color: var(--orange);
    letter-spacing: 2px;
    text-transform: uppercase;
    margin-bottom: 6px;
}

.section-title {
    font-size: 1.3rem;
    font-weight: 700;
    color: var(--text);
    letter-spacing: 2px;
    text-transform: uppercase;
    border-bottom: 1px solid var(--border);
    padding-bottom: 8px;
    margin: 20px 0 16px;
}

.qr-container {
    background: white;
    padding: 16px;
    border-radius: 4px;
    display: inline-block;
    border: 3px solid var(--orange);
}

.scan-instruction {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 4px;
    padding: 20px;
    text-align: center;
    margin-top: 12px;
}

.scan-instruction .icon { font-size: 2.5rem; margin-bottom: 8px; }
.scan-instruction p { color: var(--muted); font-size: 0.9rem; margin: 0; }
</style>
""", unsafe_allow_html=True)

# ─── Data Storage (Session State) ───────────────────────────────────────────────
if "trolleys" not in st.session_state:
    st.session_state.trolleys = {}   # { trolley_name: { "df": DataFrame, "filename": str } }

# ─── Helper Functions ────────────────────────────────────────────────────────────

EXPECTED_COLS = [
    "S.No", "Part No", "Description", "Store", "Zone",
    "Rack", "Picking Location (old Location)",
    "Trolley Location", "Qty", "Pick Qty", "Pending Qty",
    "Delivery Location", "Family"
]

def generate_qr_code(data: str, trolley_name: str) -> bytes:
    """Generate a labeled QR code image and return as PNG bytes."""
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=10,
        border=4,
    )
    qr.add_data(data)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="#FF6B2B", back_color="white").convert("RGB")

    # Add label below QR
    label_height = 70
    total_height = qr_img.height + label_height
    canvas = Image.new("RGB", (qr_img.width, total_height), "white")
    canvas.paste(qr_img, (0, 0))

    draw = ImageDraw.Draw(canvas)
    try:
        font_big = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 22)
        font_small = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 14)
    except Exception:
        font_big = ImageFont.load_default()
        font_small = font_big

    # Trolley name
    text = f"TROLLEY: {trolley_name.upper()}"
    bbox = draw.textbbox((0, 0), text, font=font_big)
    tw = bbox[2] - bbox[0]
    draw.text(((qr_img.width - tw) // 2, qr_img.height + 6), text, fill="#FF6B2B", font=font_big)

    # Scan instruction
    sub = "Scan to view Pick List"
    bbox2 = draw.textbbox((0, 0), sub, font=font_small)
    sw = bbox2[2] - bbox2[0]
    draw.text(((qr_img.width - sw) // 2, qr_img.height + 36), sub, fill="#555555", font=font_small)

    buf = io.BytesIO()
    canvas.save(buf, format="PNG", dpi=(300, 300))
    return buf.getvalue()


def df_to_styled_excel(df: pd.DataFrame, trolley_name: str) -> bytes:
    """Convert DataFrame to a styled Excel file."""
    buf = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = trolley_name[:31]

    # ── Title row ──────────────────────────────────────────────────────────────
    ws.merge_cells(f"A1:{get_column_letter(len(df.columns))}1")
    title_cell = ws["A1"]
    title_cell.value = f"PICK LIST — TROLLEY: {trolley_name.upper()}"
    title_cell.font = Font(bold=True, size=14, color="FFFFFF")
    title_cell.fill = PatternFill("solid", fgColor="FF6B2B")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # ── Header row ─────────────────────────────────────────────────────────────
    header_fill = PatternFill("solid", fgColor="1A1A1A")
    header_font = Font(bold=True, color="FF6B2B", size=10)
    thin = Side(style="thin", color="2A2A2A")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=2, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[2].height = 36

    # ── Data rows ──────────────────────────────────────────────────────────────
    alt_fill = PatternFill("solid", fgColor="1E1E1E")
    white_fill = PatternFill("solid", fgColor="161616")
    data_font = Font(color="E8E8E8", size=9)
    center_align = Alignment(horizontal="center", vertical="center")

    for row_idx, row in enumerate(df.itertuples(index=False), start=3):
        fill = alt_fill if row_idx % 2 == 0 else white_fill
        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = data_font
            cell.fill = fill
            cell.alignment = center_align
            cell.border = border
        ws.row_dimensions[row_idx].height = 20

    # ── Column widths ──────────────────────────────────────────────────────────
    col_widths = {"S.No": 6, "Part No": 14, "Description": 28, "Store": 10,
                  "Zone": 8, "Rack": 8, "Picking Location (old Location)": 24,
                  "Trolley Location": 18, "Qty": 8, "Pick Qty": 10,
                  "Pending Qty": 12, "Delivery Location": 20, "Family": 14}
    for col_idx, col_name in enumerate(df.columns, start=1):
        width = col_widths.get(col_name, 14)
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # Freeze panes
    ws.freeze_panes = "A3"

    wb.save(buf)
    return buf.getvalue()


def build_qr_url(trolley_name: str) -> str:
    """
    Build a data URL that encodes the trolley name.
    In production, this would be a real hosted URL like:
      https://yourapp.streamlit.app/?trolley=TROLLEY_A
    For local use we encode all data as a JSON data-URI trigger.
    """
    return f"TROLLEY_QR::{trolley_name}"


def make_zip_all() -> bytes:
    """Zip all QR code PNGs together."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        for name in st.session_state.trolleys:
            url = build_qr_url(name)
            qr_bytes = generate_qr_code(url, name)
            zf.writestr(f"QR_{name}.png", qr_bytes)
    return buf.getvalue()


def read_uploaded_excel(uploaded_file) -> pd.DataFrame:
    """Read uploaded Excel and normalize column names."""
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    # Strip whitespace from column names
    df.columns = [str(c).strip() for c in df.columns]
    return df


# ─── Header ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <div>
        <h1>🏭 Trolley QR Manager</h1>
        <p>Manufacturing Pick-List System · Upload → Name → Generate QR → Scan</p>
    </div>
</div>
""", unsafe_allow_html=True)

# ─── Stats Row ───────────────────────────────────────────────────────────────────
total_trolleys = len(st.session_state.trolleys)
total_parts = sum(len(v["df"]) for v in st.session_state.trolleys.values()) if total_trolleys else 0

c1, c2, c3 = st.columns(3)
with c1:
    st.markdown(f"""<div class="stat-box">
        <div class="stat-number">{total_trolleys}</div>
        <div class="stat-label">Trolleys Registered</div>
    </div>""", unsafe_allow_html=True)
with c2:
    st.markdown(f"""<div class="stat-box">
        <div class="stat-number">{total_parts}</div>
        <div class="stat-label">Total Parts</div>
    </div>""", unsafe_allow_html=True)
with c3:
    st.markdown(f"""<div class="stat-box">
        <div class="stat-number">{total_trolleys}</div>
        <div class="stat-label">QR Codes Ready</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ─── Tabs ────────────────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["📤  UPLOAD & GENERATE", "📋  VIEW PICK LISTS", "📱  SCAN SIMULATOR"])

# ════════════════════════════════════════════════════════════════════════════════
# TAB 1 — Upload & Generate
# ════════════════════════════════════════════════════════════════════════════════
with tab1:
    col_left, col_right = st.columns([1, 1], gap="large")

    # ── LEFT: Upload form ──────────────────────────────────────────────────────
    with col_left:
        st.markdown('<div class="section-title">Step 1 · Upload Excel & Name Trolley</div>', unsafe_allow_html=True)

        trolley_name = st.text_input(
            "Trolley Name / ID",
            placeholder="e.g. TROLLEY-A, LINE-3, BIN-07",
            help="This name will appear on the QR code label and the Excel sheet header."
        ).strip().upper()

        uploaded_file = st.file_uploader(
            "Upload Pick List Excel (.xlsx)",
            type=["xlsx", "xls"],
            help="Upload the Excel file containing the pick list for this trolley."
        )

        st.markdown('<div class="step-label">Expected Columns (auto-mapped if present)</div>', unsafe_allow_html=True)
        cols_preview = " · ".join(EXPECTED_COLS)
        st.caption(cols_preview)

        if st.button("✅ Register Trolley & Generate QR", use_container_width=True):
            if not trolley_name:
                st.error("⚠️ Please enter a Trolley Name.")
            elif not uploaded_file:
                st.error("⚠️ Please upload an Excel file.")
            else:
                df = read_uploaded_excel(uploaded_file)
                # Add missing columns as empty
                for col in EXPECTED_COLS:
                    if col not in df.columns:
                        df[col] = ""
                # Reorder to expected cols (keep extras at end)
                ordered = [c for c in EXPECTED_COLS if c in df.columns]
                extras = [c for c in df.columns if c not in EXPECTED_COLS]
                df = df[ordered + extras]

                st.session_state.trolleys[trolley_name] = {
                    "df": df,
                    "filename": uploaded_file.name
                }
                st.success(f"✅ Trolley **{trolley_name}** registered with {len(df)} parts!")
                st.rerun()

        # ── Registered Trolleys list ───────────────────────────────────────────
        if st.session_state.trolleys:
            st.markdown('<div class="section-title">Registered Trolleys</div>', unsafe_allow_html=True)

            for name, data in st.session_state.trolleys.items():
                row_c1, row_c2 = st.columns([3, 1])
                with row_c1:
                    st.markdown(f"""
                    <div class="trolley-card">
                        <div>
                            <div class="trolley-name">{name}</div>
                            <div class="trolley-meta">{data['filename']} · {len(data['df'])} parts</div>
                        </div>
                        <span class="badge badge-green">READY</span>
                    </div>
                    """, unsafe_allow_html=True)
                with row_c2:
                    if st.button("🗑️", key=f"del_{name}", help=f"Remove {name}"):
                        del st.session_state.trolleys[name]
                        st.rerun()

            st.markdown("---")
            st.markdown('<div class="step-label">Bulk Download</div>', unsafe_allow_html=True)
            zip_bytes = make_zip_all()
            st.download_button(
                label="⬇️ Download ALL QR Codes (.zip)",
                data=zip_bytes,
                file_name="all_trolley_qr_codes.zip",
                mime="application/zip",
                use_container_width=True
            )

    # ── RIGHT: QR Preview ──────────────────────────────────────────────────────
    with col_right:
        st.markdown('<div class="section-title">Step 2 · Preview & Download QR Code</div>', unsafe_allow_html=True)

        if st.session_state.trolleys:
            selected_qr = st.selectbox(
                "Select Trolley to Preview",
                options=list(st.session_state.trolleys.keys()),
                key="qr_preview_select"
            )

            if selected_qr:
                url = build_qr_url(selected_qr)
                qr_bytes = generate_qr_code(url, selected_qr)

                # Display QR
                qr_img = Image.open(io.BytesIO(qr_bytes))
                st.image(qr_img, caption=f"QR Code for {selected_qr}", use_container_width=False, width=320)

                st.markdown(f"""
                <div class="scan-instruction">
                    <div class="icon">📱</div>
                    <p>Print this QR code and paste it on <strong style="color:#FF6B2B">{selected_qr}</strong>.<br>
                    Anyone who scans it will see the pick list instantly — no app needed.</p>
                </div>
                """, unsafe_allow_html=True)

                st.markdown("<br>", unsafe_allow_html=True)
                st.download_button(
                    label=f"⬇️ Download QR — {selected_qr}.png",
                    data=qr_bytes,
                    file_name=f"QR_{selected_qr}.png",
                    mime="image/png",
                    use_container_width=True
                )

                # Also offer styled Excel download
                excel_bytes = df_to_styled_excel(
                    st.session_state.trolleys[selected_qr]["df"], selected_qr
                )
                st.download_button(
                    label=f"⬇️ Download Pick List Excel — {selected_qr}",
                    data=excel_bytes,
                    file_name=f"PickList_{selected_qr}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.info("ℹ️ Register at least one trolley on the left to generate a QR code.")

# ════════════════════════════════════════════════════════════════════════════════
# TAB 2 — View Pick Lists
# ════════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown('<div class="section-title">All Pick Lists</div>', unsafe_allow_html=True)

    if not st.session_state.trolleys:
        st.info("ℹ️ No trolleys registered yet. Go to the Upload tab to get started.")
    else:
        for name, data in st.session_state.trolleys.items():
            with st.expander(f"🛒  {name}  ·  {len(data['df'])} parts  ·  {data['filename']}", expanded=False):
                df = data["df"]

                # Summary metrics
                m1, m2, m3 = st.columns(3)
                if "Qty" in df.columns:
                    m1.metric("Total Qty", int(df["Qty"].sum()) if pd.api.types.is_numeric_dtype(df["Qty"]) else "—")
                if "Pick Qty" in df.columns:
                    m2.metric("Pick Qty", int(df["Pick Qty"].sum()) if pd.api.types.is_numeric_dtype(df["Pick Qty"]) else "—")
                if "Pending Qty" in df.columns:
                    m3.metric("Pending", int(df["Pending Qty"].sum()) if pd.api.types.is_numeric_dtype(df["Pending Qty"]) else "—")

                # Dataframe
                st.dataframe(df, use_container_width=True, hide_index=True)

                # Download
                excel_bytes = df_to_styled_excel(df, name)
                st.download_button(
                    label=f"⬇️ Download Excel — {name}",
                    data=excel_bytes,
                    file_name=f"PickList_{name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_excel_{name}"
                )

# ════════════════════════════════════════════════════════════════════════════════
# TAB 3 — Scan Simulator (mimics what happens when phone scans QR)
# ════════════════════════════════════════════════════════════════════════════════
with tab3:
    st.markdown('<div class="section-title">📱 Scan Simulator — What Workers See</div>', unsafe_allow_html=True)
    st.caption("This simulates what happens when a worker scans the QR code from their phone. They see a clean, mobile-friendly pick list.")

    if not st.session_state.trolleys:
        st.info("ℹ️ Register trolleys first in the Upload tab.")
    else:
        sim_trolley = st.selectbox(
            "Simulate scanning QR for trolley:",
            options=list(st.session_state.trolleys.keys()),
            key="sim_select"
        )

        if sim_trolley:
            df = st.session_state.trolleys[sim_trolley]["df"]

            # Mobile-style card display
            st.markdown(f"""
            <div style="max-width:420px; margin:0 auto; background:#1a1a1a; border-radius:12px;
                        border:1px solid #2a2a2a; overflow:hidden; box-shadow: 0 20px 60px rgba(0,0,0,0.5);">
                <!-- Phone status bar -->
                <div style="background:#000; padding:6px 16px; display:flex; justify-content:space-between;
                            align-items:center; font-size:11px; color:#aaa; font-family:monospace;">
                    <span>9:41</span><span>📶 🔋</span>
                </div>
                <!-- App header -->
                <div style="background:#FF6B2B; padding:16px 20px;">
                    <div style="font-family:'Rajdhani',sans-serif; font-size:1.4rem; font-weight:700;
                                color:#000; letter-spacing:2px;">🏭 PICK LIST</div>
                    <div style="font-size:0.85rem; color:rgba(0,0,0,0.7); font-family:monospace;">
                        TROLLEY: {sim_trolley}
                    </div>
                </div>
                <!-- Stats strip -->
                <div style="background:#111; padding:10px 20px; display:flex; gap:20px;
                            border-bottom:1px solid #2a2a2a;">
                    <div style="text-align:center;">
                        <div style="font-size:1.3rem; font-weight:700; color:#FF6B2B; font-family:monospace;">
                            {len(df)}
                        </div>
                        <div style="font-size:0.7rem; color:#888; letter-spacing:1px;">PARTS</div>
                    </div>
                    <div style="text-align:center;">
                        <div style="font-size:1.3rem; font-weight:700; color:#39D353; font-family:monospace;">
                            {int(df['Qty'].sum()) if 'Qty' in df.columns and pd.api.types.is_numeric_dtype(df['Qty']) else '–'}
                        </div>
                        <div style="font-size:0.7rem; color:#888; letter-spacing:1px;">TOTAL QTY</div>
                    </div>
                    <div style="text-align:center;">
                        <div style="font-size:1.3rem; font-weight:700; color:#F5C518; font-family:monospace;">
                            {int(df['Pending Qty'].sum()) if 'Pending Qty' in df.columns and pd.api.types.is_numeric_dtype(df['Pending Qty']) else '–'}
                        </div>
                        <div style="font-size:0.7rem; color:#888; letter-spacing:1px;">PENDING</div>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            # Show parts as mobile cards
            st.markdown("**Parts in this trolley:**")
            for i, row in df.iterrows():
                part_no = row.get("Part No", "—")
                desc = row.get("Description", "—")
                qty = row.get("Qty", "—")
                pick = row.get("Pick Qty", "—")
                pending = row.get("Pending Qty", "—")
                location = row.get("Picking Location (old Location)", "—")
                trolley_loc = row.get("Trolley Location", "—")
                zone = row.get("Zone", "—")
                rack = row.get("Rack", "—")

                pending_color = "#F5C518" if str(pending) not in ["0", "—", ""] else "#39D353"

                st.markdown(f"""
                <div style="background:#161616; border:1px solid #2a2a2a; border-left:3px solid #FF6B2B;
                            border-radius:4px; padding:12px 16px; margin-bottom:8px; max-width:600px;">
                    <div style="display:flex; justify-content:space-between; align-items:flex-start;">
                        <div>
                            <span style="font-family:monospace; font-size:0.8rem; color:#FF6B2B;
                                         font-weight:700; letter-spacing:1px;">{part_no}</span>
                            <div style="font-size:0.95rem; color:#E8E8E8; margin-top:2px;">{desc}</div>
                        </div>
                        <div style="text-align:right;">
                            <div style="font-size:1.2rem; font-weight:700; color:#E8E8E8; font-family:monospace;">
                                {qty}
                                <span style="font-size:0.65rem; color:#888; font-weight:400;">QTY</span>
                            </div>
                            <div style="font-size:0.75rem; color:{pending_color};">⏳ {pending} pending</div>
                        </div>
                    </div>
                    <div style="margin-top:8px; display:flex; gap:8px; flex-wrap:wrap;">
                        <span style="background:#1e1e1e; border:1px solid #333; padding:2px 8px;
                                     border-radius:2px; font-size:0.7rem; color:#aaa; font-family:monospace;">
                            📍 {location}
                        </span>
                        <span style="background:#1e1e1e; border:1px solid #333; padding:2px 8px;
                                     border-radius:2px; font-size:0.7rem; color:#aaa; font-family:monospace;">
                            🛒 {trolley_loc}
                        </span>
                        <span style="background:#1e1e1e; border:1px solid #333; padding:2px 8px;
                                     border-radius:2px; font-size:0.7rem; color:#aaa; font-family:monospace;">
                            Zone: {zone} · Rack: {rack}
                        </span>
                    </div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("---")
            excel_bytes = df_to_styled_excel(df, sim_trolley)
            st.download_button(
                label="⬇️ Download This Pick List",
                data=excel_bytes,
                file_name=f"PickList_{sim_trolley}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="sim_dl"
            )

# ─── Sidebar: Quick How-To ───────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="font-family:'Rajdhani',sans-serif; font-size:1.1rem; font-weight:700;
                color:#FF6B2B; letter-spacing:2px; text-transform:uppercase;
                border-bottom:1px solid #2a2a2a; padding-bottom:8px; margin-bottom:16px;">
        How It Works
    </div>
    """, unsafe_allow_html=True)

    steps = [
        ("1", "Upload Excel", "Prepare a .xlsx file with the pick list columns for each trolley."),
        ("2", "Name Trolley", "Enter the trolley ID (e.g. TROLLEY-A). Each trolley gets its own QR."),
        ("3", "Generate QR", "Click Register. A labeled QR code is created instantly."),
        ("4", "Download & Print", "Download the QR PNG and print it. Paste on the physical trolley."),
        ("5", "Worker Scans", "Worker scans QR with any phone camera. Pick list opens immediately — no app needed."),
    ]

    for num, title, desc in steps:
        st.markdown(f"""
        <div style="display:flex; gap:12px; margin-bottom:14px; align-items:flex-start;">
            <div style="background:#FF6B2B; color:#000; font-weight:700; font-size:0.8rem;
                        width:22px; height:22px; border-radius:50%; display:flex; align-items:center;
                        justify-content:center; flex-shrink:0; font-family:monospace;">{num}</div>
            <div>
                <div style="font-weight:700; font-size:0.9rem; color:#E8E8E8; letter-spacing:1px;">{title}</div>
                <div style="font-size:0.78rem; color:#888; margin-top:2px; line-height:1.4;">{desc}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("""
    <div style="font-size:0.78rem; color:#555; font-family:monospace; line-height:1.6;">
        💡 <strong style="color:#888">Deployment Tip:</strong><br>
        Host on Streamlit Cloud and the QR code will encode a real URL — 
        scanning it will open the pick list directly in the worker's browser.
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    if st.session_state.trolleys and st.button("🗑️ Clear All Trolleys", use_container_width=True):
        st.session_state.trolleys = {}
        st.rerun()
