import streamlit as st
import pandas as pd
import qrcode
import io
import json
import zipfile
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from PIL import Image

st.set_page_config(
    page_title="Trolley QR Pick List System",
    page_icon="🏭",
    layout="wide"
)

# ── Styling ──────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  .main-header {
    background: linear-gradient(135deg,#1a237e,#0d47a1);
    color:white; padding:24px 32px; border-radius:12px;
    margin-bottom:24px; text-align:center;
  }
  .trolley-card {
    background:#f8f9fa; border:2px solid #e0e0e0;
    border-radius:12px; padding:20px; margin:10px 0;
    transition:border-color .2s;
  }
  .trolley-card:hover { border-color:#1976d2; }
  .qr-container {
    background:white; border-radius:12px; padding:20px;
    box-shadow:0 4px 12px rgba(0,0,0,.1); text-align:center;
  }
  .stat-box {
    background:white; border-left:4px solid #1976d2;
    padding:16px; border-radius:8px;
    box-shadow:0 2px 8px rgba(0,0,0,.08);
  }
  .success-badge {
    background:#e8f5e9; color:#2e7d32;
    padding:4px 12px; border-radius:20px;
    font-size:13px; font-weight:600;
  }
  .warning-badge {
    background:#fff3e0; color:#e65100;
    padding:4px 12px; border-radius:20px;
    font-size:13px; font-weight:600;
  }
</style>
""", unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────────────────────
if "trolleys" not in st.session_state:
    st.session_state.trolleys = {}          # {name: [rows]}
if "active_tab" not in st.session_state:
    st.session_state.active_tab = "setup"

COLUMNS = [
    "S.No", "Part No", "Description", "Store", "Zone", "Rack",
    "Picking Location (old Location)", "Trolley Location",
    "Qty", "Pick Qty", "Pending Qty", "Delivery Location", "Family"
]

# ── Helpers ───────────────────────────────────────────────────────────────────
def make_qr(data: str, size: int = 8) -> Image.Image:
    qr = qrcode.QRCode(
        version=1, error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=size, border=4
    )
    qr.add_data(data)
    qr.make(fit=True)
    return qr.make_image(fill_color="black", back_color="white").convert("RGB")


def make_excel(trolley_name: str, rows: list[dict]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = trolley_name[:31]

    # Colours
    hdr_fill  = PatternFill("solid", start_color="1a237e")
    alt_fill  = PatternFill("solid", start_color="E8EAF6")
    thin = Side(style="thin", color="BDBDBD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Title row
    ws.merge_cells("A1:M1")
    title_cell = ws["A1"]
    title_cell.value = f"PICK LIST — TROLLEY: {trolley_name.upper()}"
    title_cell.font = Font(bold=True, size=14, color="FFFFFF")
    title_cell.fill = PatternFill("solid", start_color="0D47A1")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # Header row
    for col_idx, col_name in enumerate(COLUMNS, 1):
        cell = ws.cell(row=2, column=col_idx, value=col_name)
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.fill = hdr_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[2].height = 36

    # Data rows
    for row_idx, row in enumerate(rows, 3):
        for col_idx, col_name in enumerate(COLUMNS, 1):
            val = row.get(col_name, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
            if row_idx % 2 == 0:
                cell.fill = alt_fill
        # Pending Qty formula
        qty_col   = COLUMNS.index("Qty") + 1
        pick_col  = COLUMNS.index("Pick Qty") + 1
        pend_col  = COLUMNS.index("Pending Qty") + 1
        qty_letter  = get_column_letter(qty_col)
        pick_letter = get_column_letter(pick_col)
        pend_cell   = ws.cell(row=row_idx, column=pend_col)
        pend_cell.value = f"={qty_letter}{row_idx}-{pick_letter}{row_idx}"
        pend_cell.font  = Font(color="C62828")

    # Column widths
    col_widths = [6,14,28,10,8,10,24,18,8,8,10,18,12]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Freeze panes
    ws.freeze_panes = "A3"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def qr_to_bytes(img: Image.Image) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <h1 style="margin:0;font-size:2rem">🏭 Trolley QR Pick List System</h1>
  <p style="margin:8px 0 0;opacity:.85">Generate unique QR codes for each trolley → scan to download its pick list</p>
</div>
""", unsafe_allow_html=True)

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab_setup, tab_qr, tab_scan, tab_all = st.tabs([
    "⚙️ Setup Trolleys",
    "📱 QR Codes",
    "🔍 Scan / Preview",
    "📦 Download All"
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — Setup
# ══════════════════════════════════════════════════════════════════════════════
with tab_setup:
    st.subheader("Configure Trolleys & Pick Lists")

    col_add, col_imp = st.columns([1, 1])

    # ── Add trolley manually ──────────────────────────────────────────────────
    with col_add:
        st.markdown("### ➕ Add a New Trolley")
        trolley_name = st.text_input(
            "Trolley Name", placeholder="e.g. TROLLEY-A1",
            key="new_trolley_name"
        ).strip().upper()

        st.markdown("**Add parts to this trolley:**")
        with st.form("add_part_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            part_no  = c1.text_input("Part No")
            desc     = c2.text_input("Description")
            store    = c1.text_input("Store")
            zone     = c2.text_input("Zone")
            rack     = c1.text_input("Rack")
            pick_loc = c2.text_input("Picking Location (old Location)")
            trol_loc = c1.text_input("Trolley Location")
            family   = c2.text_input("Family")
            qty      = c1.number_input("Qty", min_value=0, value=1)
            pick_qty = c2.number_input("Pick Qty", min_value=0, value=0)

            submitted = st.form_submit_button("➕ Add Part", use_container_width=True)
            if submitted:
                if not trolley_name:
                    st.error("Enter a trolley name first.")
                elif not part_no:
                    st.error("Part No is required.")
                else:
                    rows = st.session_state.trolleys.setdefault(trolley_name, [])
                    rows.append({
                        "S.No": len(rows) + 1,
                        "Part No": part_no,
                        "Description": desc,
                        "Store": store,
                        "Zone": zone,
                        "Rack": rack,
                        "Picking Location (old Location)": pick_loc,
                        "Trolley Location": trol_loc,
                        "Qty": qty,
                        "Pick Qty": pick_qty,
                        "Pending Qty": qty - pick_qty,
                        "Delivery Location": "",
                        "Family": family,
                    })
                    st.success(f"Part added to {trolley_name}!")

    # ── Import via Excel ──────────────────────────────────────────────────────
    with col_imp:
        st.markdown("### 📂 Import from Excel")
        st.info(
            "Upload an Excel file where **each sheet name = trolley name** "
            "and columns match the pick-list format."
        )
        uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
        if uploaded:
            try:
                xl = pd.read_excel(uploaded, sheet_name=None, dtype=str)
                for sheet, df in xl.items():
                    df.fillna("", inplace=True)
                    rows = []
                    for i, row in df.iterrows():
                        entry = {col: row.get(col, "") for col in COLUMNS}
                        entry["S.No"] = i + 1
                        rows.append(entry)
                    st.session_state.trolleys[sheet.strip().upper()] = rows
                st.success(f"Imported {len(xl)} trolley(s): {', '.join(xl.keys())}")
            except Exception as e:
                st.error(f"Import failed: {e}")

        st.markdown("---")
        st.markdown("### 📋 Quick Demo Data")
        if st.button("Load Sample Trolleys", use_container_width=True):
            demo = {
                "TROLLEY-A1": [
                    {"S.No":1,"Part No":"P001","Description":"Bolt M8x20","Store":"ST-01","Zone":"Z1","Rack":"R2","Picking Location (old Location)":"L-12","Trolley Location":"T-A1-01","Qty":50,"Pick Qty":0,"Pending Qty":50,"Delivery Location":"LINE-1","Family":"Fasteners"},
                    {"S.No":2,"Part No":"P002","Description":"Nut M8","Store":"ST-01","Zone":"Z1","Rack":"R2","Picking Location (old Location)":"L-13","Trolley Location":"T-A1-02","Qty":50,"Pick Qty":0,"Pending Qty":50,"Delivery Location":"LINE-1","Family":"Fasteners"},
                    {"S.No":3,"Part No":"P003","Description":"Washer 8mm","Store":"ST-02","Zone":"Z2","Rack":"R4","Picking Location (old Location)":"L-20","Trolley Location":"T-A1-03","Qty":100,"Pick Qty":0,"Pending Qty":100,"Delivery Location":"LINE-1","Family":"Fasteners"},
                ],
                "TROLLEY-B2": [
                    {"S.No":1,"Part No":"P010","Description":"Bearing 6205","Store":"ST-03","Zone":"Z3","Rack":"R5","Picking Location (old Location)":"L-30","Trolley Location":"T-B2-01","Qty":10,"Pick Qty":0,"Pending Qty":10,"Delivery Location":"LINE-2","Family":"Bearings"},
                    {"S.No":2,"Part No":"P011","Description":"Seal Ring 25mm","Store":"ST-03","Zone":"Z3","Rack":"R6","Picking Location (old Location)":"L-31","Trolley Location":"T-B2-02","Qty":20,"Pick Qty":0,"Pending Qty":20,"Delivery Location":"LINE-2","Family":"Seals"},
                ],
                "TROLLEY-C3": [
                    {"S.No":1,"Part No":"P020","Description":"Motor Bracket","Store":"ST-04","Zone":"Z4","Rack":"R8","Picking Location (old Location)":"L-40","Trolley Location":"T-C3-01","Qty":5,"Pick Qty":0,"Pending Qty":5,"Delivery Location":"LINE-3","Family":"Brackets"},
                    {"S.No":2,"Part No":"P021","Description":"Cable Harness","Store":"ST-04","Zone":"Z4","Rack":"R9","Picking Location (old Location)":"L-42","Trolley Location":"T-C3-02","Qty":5,"Pick Qty":0,"Pending Qty":5,"Delivery Location":"LINE-3","Family":"Electrical"},
                    {"S.No":3,"Part No":"P022","Description":"Connector 12-pin","Store":"ST-05","Zone":"Z4","Rack":"R9","Picking Location (old Location)":"L-43","Trolley Location":"T-C3-03","Qty":15,"Pick Qty":0,"Pending Qty":15,"Delivery Location":"LINE-3","Family":"Electrical"},
                ],
            }
            st.session_state.trolleys.update(demo)
            st.success("Sample data loaded!")

    # ── Current trolleys ──────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 🚛 Configured Trolleys")
    if not st.session_state.trolleys:
        st.info("No trolleys yet. Add one above or load sample data.")
    else:
        for t_name, t_rows in list(st.session_state.trolleys.items()):
            with st.expander(f"🚛 {t_name}  —  {len(t_rows)} part(s)"):
                df_view = pd.DataFrame(t_rows)
                st.dataframe(df_view, use_container_width=True, hide_index=True)
                if st.button(f"🗑️ Delete {t_name}", key=f"del_{t_name}"):
                    del st.session_state.trolleys[t_name]
                    st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — QR Codes
# ══════════════════════════════════════════════════════════════════════════════
with tab_qr:
    st.subheader("📱 QR Codes for All Trolleys")
    if not st.session_state.trolleys:
        st.info("Set up trolleys in the **Setup** tab first.")
    else:
        st.markdown(
            "Each QR code encodes the trolley identity. "
            "Scan with any QR reader or use the **Scan / Preview** tab."
        )
        cols = st.columns(min(len(st.session_state.trolleys), 3))
        for idx, (t_name, t_rows) in enumerate(st.session_state.trolleys.items()):
            with cols[idx % 3]:
                st.markdown(f'<div class="qr-container">', unsafe_allow_html=True)
                # QR payload: JSON with trolley id
                payload = json.dumps({"trolley": t_name, "parts": len(t_rows)})
                qr_img = make_qr(payload, size=7)
                st.image(qr_img, caption=t_name, use_container_width=True)

                # Download QR PNG
                qr_bytes = qr_to_bytes(qr_img)
                st.download_button(
                    "⬇️ Download QR",
                    data=qr_bytes,
                    file_name=f"QR_{t_name}.png",
                    mime="image/png",
                    use_container_width=True,
                    key=f"qr_dl_{t_name}"
                )
                # Download Excel directly
                xl_bytes = make_excel(t_name, t_rows)
                st.download_button(
                    "📥 Download Pick List",
                    data=xl_bytes,
                    file_name=f"PickList_{t_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key=f"xl_dl_{t_name}"
                )
                total_qty = sum(r.get("Qty", 0) for r in t_rows if str(r.get("Qty", "")).isdigit() or isinstance(r.get("Qty"), (int, float)))
                st.markdown(
                    f'<p style="margin-top:8px;font-size:13px;color:#555">'
                    f'Parts: <b>{len(t_rows)}</b> &nbsp;|&nbsp; '
                    f'Total Qty: <b>{int(total_qty) if isinstance(total_qty, float) else total_qty}</b></p>',
                    unsafe_allow_html=True
                )
                st.markdown("</div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — Scan / Preview
# ══════════════════════════════════════════════════════════════════════════════
with tab_scan:
    st.subheader("🔍 Scan QR / Preview Pick List")
    st.info(
        "In a real deployment a mobile scanner passes the trolley name via URL parameter. "
        "Here you can **select a trolley manually** to simulate a scan."
    )
    if not st.session_state.trolleys:
        st.warning("No trolleys configured yet.")
    else:
        selected = st.selectbox(
            "Select / Scan Trolley",
            options=list(st.session_state.trolleys.keys()),
            format_func=lambda x: f"🚛 {x}"
        )
        if selected:
            rows = st.session_state.trolleys[selected]
            st.markdown(f"### Pick List — {selected}")

            # Stats row
            s1, s2, s3, s4 = st.columns(4)
            total_parts = len(rows)
            total_qty   = sum(r.get("Qty", 0) for r in rows if isinstance(r.get("Qty"), (int, float)))
            picked      = sum(r.get("Pick Qty", 0) for r in rows if isinstance(r.get("Pick Qty"), (int, float)))
            pending     = total_qty - picked
            s1.metric("Total Parts", total_parts)
            s2.metric("Total Qty",   int(total_qty))
            s3.metric("Picked",      int(picked))
            s4.metric("Pending",     int(pending))

            # Table
            df_show = pd.DataFrame(rows)[COLUMNS]
            st.dataframe(df_show, use_container_width=True, hide_index=True, height=350)

            # Download
            xl = make_excel(selected, rows)
            st.download_button(
                f"📥 Download {selected} Pick List (.xlsx)",
                data=xl,
                file_name=f"PickList_{selected}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

            # QR preview
            with st.expander("Show QR Code for this trolley"):
                payload = json.dumps({"trolley": selected, "parts": len(rows)})
                qr_img = make_qr(payload, size=10)
                c1, c2, c3 = st.columns([1, 2, 1])
                with c2:
                    st.image(qr_img, caption=f"QR — {selected}", use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — Download All
# ══════════════════════════════════════════════════════════════════════════════
with tab_all:
    st.subheader("📦 Bulk Download")
    if not st.session_state.trolleys:
        st.info("Set up trolleys first.")
    else:
        st.markdown(f"**{len(st.session_state.trolleys)} trolley(s) ready.**")

        # Summary table
        summary = []
        for t_name, t_rows in st.session_state.trolleys.items():
            total = sum(r.get("Qty", 0) for r in t_rows if isinstance(r.get("Qty"), (int, float)))
            summary.append({"Trolley": t_name, "Parts": len(t_rows), "Total Qty": int(total)})
        st.dataframe(pd.DataFrame(summary), use_container_width=True, hide_index=True)

        st.markdown("---")
        col_a, col_b = st.columns(2)

        with col_a:
            st.markdown("#### 📊 All Pick Lists (.xlsx ZIP)")
            if st.button("Generate Excel ZIP", use_container_width=True):
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    for t_name, t_rows in st.session_state.trolleys.items():
                        xl_bytes = make_excel(t_name, t_rows)
                        zf.writestr(f"PickList_{t_name}.xlsx", xl_bytes)
                st.download_button(
                    "⬇️ Download Excel ZIP",
                    data=zip_buf.getvalue(),
                    file_name="All_PickLists.zip",
                    mime="application/zip",
                    use_container_width=True
                )

        with col_b:
            st.markdown("#### 📱 All QR Codes (.png ZIP)")
            if st.button("Generate QR ZIP", use_container_width=True):
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    for t_name, t_rows in st.session_state.trolleys.items():
                        payload = json.dumps({"trolley": t_name, "parts": len(t_rows)})
                        qr_img  = make_qr(payload, size=10)
                        zf.writestr(f"QR_{t_name}.png", qr_to_bytes(qr_img))
                st.download_button(
                    "⬇️ Download QR ZIP",
                    data=zip_buf.getvalue(),
                    file_name="All_QRCodes.zip",
                    mime="application/zip",
                    use_container_width=True
                )

        st.markdown("---")
        st.markdown("#### 🔄 Reset All Data")
        if st.button("⚠️ Clear All Trolleys", type="secondary"):
            st.session_state.trolleys = {}
            st.rerun()
