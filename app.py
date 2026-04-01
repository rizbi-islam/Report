import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Excel Filter & Export Tool", layout="wide")

st.title("📊 Excel Filter & Export Tool")
st.caption("Filter by zone → export ward sheets + summary pivot")

# ── Upload ────────────────────────────────────────────────────────────────────
uploaded = st.file_uploader("Upload Excel / CSV file", type=["xlsx", "xls", "csv"])

if not uploaded:
    st.info("Upload a file to get started.")
    st.stop()

@st.cache_data
def load_data(file_bytes, name):
    if name.endswith(".csv"):
        return pd.read_csv(io.BytesIO(file_bytes), dtype=str)
    return pd.read_excel(io.BytesIO(file_bytes), dtype=str)

raw_bytes = uploaded.read()
df_raw = load_data(raw_bytes, uploaded.name)
df_raw.columns = df_raw.columns.str.strip()

st.success(f"Loaded **{len(df_raw):,}** rows × **{len(df_raw.columns)}** columns")

# ── Sidebar controls ──────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Controls")

    # Detect zone/ward columns (flexible fallback)
    zone_candidates = [c for c in df_raw.columns if "zone" in c.lower()]
    ward_candidates = [c for c in df_raw.columns if "ward" in c.lower()]

    zone_col = st.selectbox("Zone column", zone_candidates or df_raw.columns.tolist(),
                            index=0)
    ward_col = st.selectbox("Ward column", ward_candidates or df_raw.columns.tolist(),
                            index=0 if not ward_candidates else 0)

    # Numeric/amount column for pivot
    num_candidates = [c for c in df_raw.columns if any(
        k in c.lower() for k in ["amount", "paid", "fee", "total", "sum"])]
    amount_col = st.selectbox("Amount column (for pivot)", num_candidates or df_raw.columns.tolist())

    st.divider()

    # Zone filter
    all_zones = sorted(df_raw[zone_col].dropna().unique().tolist())
    sel_zones = st.multiselect("Filter by Zone(s)", all_zones, default=all_zones)

    st.divider()

    # Column visibility
    st.subheader("Column selector")
    st.caption("Uncheck to hide from export")
    col_flags = {}
    for c in df_raw.columns:
        col_flags[c] = st.checkbox(c, value=True, key=f"col_{c}")

    st.divider()

    # Dash cleaner
    remove_dash = st.checkbox("Remove trailing dash (—) from E-Holding", value=True)
    eholding_candidates = [c for c in df_raw.columns if "holding" in c.lower() or "e-holding" in c.lower()]
    eholding_col = st.selectbox("E-Holding column", eholding_candidates or df_raw.columns.tolist(),
                                disabled=not remove_dash)

# ── Apply filters & transforms ────────────────────────────────────────────────
df = df_raw.copy()

# Zone filter
if sel_zones:
    df = df[df[zone_col].isin(sel_zones)]

# Column filter
keep_cols = [c for c, v in col_flags.items() if v]
df = df[keep_cols]

# Remove trailing dashes from E-Holding
if remove_dash and eholding_col in df.columns:
    df[eholding_col] = df[eholding_col].astype(str).str.rstrip("-").str.strip()

# ── Preview ───────────────────────────────────────────────────────────────────
st.subheader("📋 Preview (first 200 rows)")
st.dataframe(df.head(200), use_container_width=True, height=300)
st.caption(f"Filtered: **{len(df):,}** rows | Columns: {len(df.columns)}")

# ── Build Excel ───────────────────────────────────────────────────────────────
def build_excel(df: pd.DataFrame, zone_col: str, ward_col: str, amount_col: str) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)  # remove default sheet

    HDR_FILL  = PatternFill("solid", start_color="2E75B6")
    HDR_FONT  = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    DATA_FONT = Font(name="Arial", size=10)
    ALT_FILL  = PatternFill("solid", start_color="D9E1F2")
    CENTER    = Alignment(horizontal="center", vertical="center")
    LEFT      = Alignment(horizontal="left",   vertical="center")
    thin      = Side(style="thin", color="B0B0B0")
    BORDER    = Border(left=thin, right=thin, top=thin, bottom=thin)

    def style_header(ws, ncols):
        for col in range(1, ncols + 1):
            cell = ws.cell(row=1, column=col)
            cell.fill  = HDR_FILL
            cell.font  = HDR_FONT
            cell.alignment = CENTER
            cell.border = BORDER

    def write_sheet(ws, data: pd.DataFrame):
        # Header
        for ci, col in enumerate(data.columns, 1):
            ws.cell(row=1, column=ci, value=col)
        style_header(ws, len(data.columns))

        # Rows
        for ri, (_, row) in enumerate(data.iterrows(), 2):
            fill = ALT_FILL if ri % 2 == 0 else None
            for ci, val in enumerate(row, 1):
                cell = ws.cell(row=ri, column=ci, value=val)
                cell.font   = DATA_FONT
                cell.border = BORDER
                cell.alignment = CENTER if ci > 1 else LEFT
                if fill:
                    cell.fill = fill

        # Auto column width
        for col in ws.columns:
            max_len = max((len(str(c.value)) if c.value else 0) for c in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 40)

        ws.freeze_panes = "A2"

    # ── Ward sheets ───────────────────────────────────────────────────────────
    ward_summary = []

    for zone_val in sorted(df[zone_col].dropna().unique()):
        zone_df = df[df[zone_col] == zone_val]

        for ward_val in sorted(zone_df[ward_col].dropna().unique()):
            ward_df = zone_df[zone_df[ward_col] == ward_val]

            # Sanitize sheet name (Excel limit 31 chars, no special chars)
            raw_name = f"{zone_val}_{ward_val}"
            sheet_name = "".join(c for c in raw_name if c not in r'\/:*?[]')[:31]

            ws = wb.create_sheet(title=sheet_name)
            write_sheet(ws, ward_df)

            # Collect for summary
            if amount_col in ward_df.columns:
                total = pd.to_numeric(ward_df[amount_col], errors="coerce").sum()
            else:
                total = None

            ward_summary.append({
                "Zone":       zone_val,
                "Ward":       ward_val,
                "Row Count":  len(ward_df),
                "Total Amount": total,
            })

    # ── Summary + Pivot sheet ─────────────────────────────────────────────────
    ws_sum = wb.create_sheet(title="📊 Summary")
    summary_df = pd.DataFrame(ward_summary)

    # Write detail table
    write_sheet(ws_sum, summary_df)

    # Pivot section — zone totals
    pivot_start_row = len(summary_df) + 4

    ws_sum.cell(row=pivot_start_row, column=1, value="PIVOT — Zone Totals").font = Font(
        bold=True, size=12, name="Arial", color="2E75B6")

    pivot_headers = ["Zone", "Wards", "Total Rows", "Total Amount"]
    for ci, h in enumerate(pivot_headers, 1):
        cell = ws_sum.cell(row=pivot_start_row + 1, column=ci, value=h)
        cell.fill  = HDR_FILL
        cell.font  = HDR_FONT
        cell.alignment = CENTER
        cell.border = BORDER

    if not summary_df.empty:
        pivot = (summary_df
                 .groupby("Zone", as_index=False)
                 .agg(Wards=("Ward", "count"),
                      Total_Rows=("Row Count", "sum"),
                      Total_Amount=("Total Amount", "sum"))
                 .rename(columns={"Total_Rows": "Total Rows",
                                   "Total_Amount": "Total Amount"}))

        for ri, (_, row) in enumerate(pivot.iterrows(), pivot_start_row + 2):
            fill = ALT_FILL if ri % 2 == 0 else None
            for ci, val in enumerate(row, 1):
                cell = ws_sum.cell(row=ri, column=ci, value=val)
                cell.font   = DATA_FONT
                cell.border = BORDER
                cell.alignment = CENTER if ci > 1 else LEFT
                if fill:
                    cell.fill = fill

        # Grand total row
        last_row = pivot_start_row + 2 + len(pivot)
        ws_sum.cell(row=last_row, column=1, value="GRAND TOTAL").font = Font(
            bold=True, name="Arial")
        ws_sum.cell(row=last_row, column=2,
                    value=f"=SUM({get_column_letter(2)}{pivot_start_row+2}:{get_column_letter(2)}{last_row-1})")
        ws_sum.cell(row=last_row, column=3,
                    value=f"=SUM({get_column_letter(3)}{pivot_start_row+2}:{get_column_letter(3)}{last_row-1})")
        ws_sum.cell(row=last_row, column=4,
                    value=f"=SUM({get_column_letter(4)}{pivot_start_row+2}:{get_column_letter(4)}{last_row-1})")
        for ci in range(1, 5):
            ws_sum.cell(row=last_row, column=ci).fill   = PatternFill("solid", start_color="C6EFCE")
            ws_sum.cell(row=last_row, column=ci).font   = Font(bold=True, name="Arial")
            ws_sum.cell(row=last_row, column=ci).border = BORDER

    # Move Summary to front
    wb.move_sheet("📊 Summary", offset=-len(wb.sheetnames) + 1)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ── Export button ─────────────────────────────────────────────────────────────
st.divider()
col1, col2 = st.columns([2, 1])

with col1:
    if ward_col not in df.columns or zone_col not in df.columns:
        st.warning("Ward/Zone columns not found in filtered data — adjust column selector.")
    elif df.empty:
        st.warning("No data after filters.")
    else:
        if st.button("🚀 Generate Excel Export", type="primary", use_container_width=True):
            with st.spinner("Building Excel workbook…"):
                excel_bytes = build_excel(df, zone_col, ward_col, amount_col)

            n_wards = df.groupby([zone_col, ward_col]).ngroups
            st.success(f"✅ Done! {n_wards} ward sheet(s) + 1 summary sheet")

            st.download_button(
                label="⬇️ Download Excel",
                data=excel_bytes,
                file_name="filtered_export.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

with col2:
    if not df.empty and ward_col in df.columns and zone_col in df.columns:
        st.metric("Zones selected", len(sel_zones))
        st.metric("Total rows", f"{len(df):,}")
        st.metric("Ward sheets", df.groupby([zone_col, ward_col]).ngroups)
