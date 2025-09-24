import io
import re
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

# =========================
# App Setup & Constants
# =========================
st.set_page_config(page_title="HCHSP Services & Referrals — Tool", layout="wide")

logo_path = Path("header_logo.png")
HCHSP_NAVY = "#305496"
HCHSP_RED = "#C00000"
HCHSP_LIGHT = "#D9E1F2"

# =========================
# Header
# =========================
hdr_l, hdr_c, hdr_r = st.columns([1, 2, 1])
with hdr_c:
    if logo_path.exists():
        st.image(str(logo_path), width=320)
    st.markdown(
        "<h1 style='text-align:center; margin: 8px 0 4px;'>Hidalgo County Head Start — Services & Referrals</h1>",
        unsafe_allow_html=True,
    )
    now_ct = datetime.now(ZoneInfo("America/Chicago")).strftime("%m/%d/%y %I:%M %p CT")
    st.markdown(
        f"<p style='text-align:center; font-weight:600; color:{HCHSP_RED}; margin-top:0;'>(as of {now_ct})</p>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<p style='text-align:center; font-size:16px; margin-top:0;'>Upload the <strong>10433</strong> Services/Referrals report (Excel). The tool will build: <em>Services & Referrals</em>, <em>PIR Summary</em>, <em>Author Fix List</em>, <em>PIR Dashboard</em>, and <em>PIS Dashboard</em>.</p>",
        unsafe_allow_html=True,
    )
st.divider()

# =========================
# Sidebar Options
# =========================
with st.sidebar:
    st.header("Settings")
    cutoff = st.date_input(
        "Cutoff (Service Date on/after)", value=pd.to_datetime("2025-08-11")
    ).strftime("%Y-%m-%d")
    st.caption("Only rows with Service Date ≥ this date are included.")
    st.checkbox("Require 'PIR' in Detailed Service", value=True, key="require_pir")
    st.caption("PIR must contain 'PIR' and a C.44 letter code (e.g., 'PIR C.44 n').")

# =========================
# Input
# =========================
inp_l, inp_c, inp_r = st.columns([1, 2, 1])
with inp_c:
    sref_file = st.file_uploader("Upload *10433.xlsx*", type=["xlsx"], key="sref")
    process = st.button("Process & Download")

# =========================
# Helpers
# =========================
def _clean_header(h: str) -> str:
    return re.sub(r"^(ST:|FD:)\s*", "", str(h).strip(), flags=re.I)

def _parse_to_dt(series: pd.Series) -> pd.Series:
    """Robust datetime parser for mixed Excel date formats (strings, serials)."""
    dt1 = pd.to_datetime(series, errors="coerce", infer_datetime_format=True)
    num = pd.to_numeric(series, errors="coerce")
    serial_mask = num.notna() & num.between(10000, 70000)
    dt2 = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    if serial_mask.any():
        dt2.loc[serial_mask] = pd.to_datetime(
            num.loc[serial_mask], unit="D", origin="1899-12-30", errors="coerce"
        )
    dt = dt1.copy()
    dt[dt.isna()] = dt2[dt.isna()]
    return dt

def _extract_pir_code(text: str) -> str | None:
    """Extracts normalized PIR code like 'C.44 n' (lowercased letter) from detail text."""
    if not isinstance(text, str):
        text = str(text)
    m = re.search(r"(?i)\bC\s*\.?\s*44\s*([a-z])\b", text)
    return f"C.44 {m.group(1).lower()}" if m else None

def _format_pid_display(val) -> str:
    """Pretty PID for output (removes trailing .0 if numeric-like)."""
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if re.fullmatch(r"-?\d+(\.0+)?", s):
        try:
            return str(int(float(s)))
        except Exception:
            return s
    return s

def _pid_key(val: object) -> str:
    """
    Stable string key for PID:
      - trims spaces
      - normalizes '322515.0' -> '322515'
      - keeps any non-numeric characters to avoid losing distinct IDs
    """
    s = str(val).strip()
    return s.split(".")[0] if re.fullmatch(r"-?\d+(\.0+)?", s) else s

def _fid_key(val: object) -> str:
    """Stable string key for FID (trim spaces)."""
    return str(val).strip()

def _col_letter(idx: int) -> str:
    s = ""
    n = idx
    while n >= 0:
        s = chr(n % 26 + 65) + s
        n = n // 26 - 1
    return s

# =========================
# Core Builder
# =========================
def build_workbook(df_raw: pd.DataFrame, cutoff: str, require_pir: bool = True) -> bytes:
    df = df_raw.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # Column detection
    FID_COL   = "Family ID"
    PID_COL   = next((c for c in df.columns if "participant pid" in c.lower()
                      or ("pid" in c.lower() and "participant" in c.lower())), "ST: Participant PID")
    LNAME_COL = next((c for c in df.columns if "last name" in c.lower()), "ST: Participant Last Name")
    FNAME_COL = next((c for c in df.columns if "first name" in c.lower()), "ST: Participant First Name")
    GEN_COL   = next((c for c in df.columns if "general service" in c.lower()), "FD: Services - General Service")
    DET_COL   = next((c for c in df.columns if "detail service" in c.lower()), "FD: Services - Detail Service")
    RES_COL   = next((c for c in df.columns if "result" in c.lower()), "FD: Services - Result")
    DATE_COL  = next((c for c in df.columns if "date" in c.lower()), "FD: Services - Date")
    AUTH_COL  = next((c for c in df.columns if ("author" in c.lower() and "service" in c.lower())
                      or "worker" in c.lower() or "staff" in c.lower()), None)
    CENTER_COL = next((c for c in df.columns if "center" in c.lower() or "campus" in c.lower()), None)

    # Parse and filter by date
    df[DATE_COL] = _parse_to_dt(df[DATE_COL])
    df = df[df[DATE_COL].notna() & (df[DATE_COL] >= pd.Timestamp(cutoff))].copy()

    # Normalized keys (CRITICAL FIX so Summary matches Services sheet)
    df["_PID_KEY"] = df[PID_COL].apply(_pid_key)
    df["_FID_KEY"] = df[FID_COL].apply(_fid_key)

    # PIR logic
    df["_Result_norm"] = df[RES_COL].astype(str).str.strip().str.lower()
    valid_result = df["_Result_norm"].isin({"service ongoing", "service completed"})
    has_pir = df[DET_COL].astype(str).str.contains("pir", case=False, na=False) if require_pir else True
    df["_has_PIR"] = has_pir
    df["_PIR_CODE"] = df[DET_COL].astype(str).map(_extract_pir_code)

    # Month bucket for PIS dashboard
    df["_month"] = df[DATE_COL].dt.to_period("M").dt.to_timestamp()

    # Candidate + de-dupe per (pid_key, PIR code)
    count_candidate = df["_has_PIR"] & valid_result & df["_PIR_CODE"].notna()
    dup_mask = pd.Series(False, index=df.index)
    if count_candidate.any():
        sub = pd.DataFrame({
            "pid_key": df.loc[count_candidate, "_PID_KEY"],
            "code": df.loc[count_candidate, "_PIR_CODE"].astype(str).str.strip().str.lower(),
        }, index=df.index[count_candidate])
        dup_mask.loc[count_candidate] = sub.duplicated(subset=["pid_key", "code"], keep="first").values

    df["Counts for PIR"] = (count_candidate & ~dup_mask).map({True: "Yes", False: "No"})

    def reason_fn(row):
        if row["Counts for PIR"] == "Yes":
            return ""
        gen = str(row.get(GEN_COL, "")).strip()
        det = str(row.get(DET_COL, "")).strip()
        res = str(row.get(RES_COL, "")).strip().lower()
        if gen == "" or gen.lower() == "nan":
            return "Missing General Service"
        if det == "" or det.lower() == "nan":
            return "Missing Detailed Service"
        if res not in {"service ongoing", "service completed"}:
            return "Invalid/Missing Result"
        if row.name in dup_mask.index and dup_mask.loc[row.name]:
            return "Duplicate Entry"
        if pd.isna(row[DATE_COL]):
            return "Missing Service Date"
        return ""

    df["Reason (if not counted)"] = df.apply(reason_fn, axis=1)

    # =========================
    # Sheet 1: Services & Referrals (detail rows)
    # =========================
    cols = [FID_COL, PID_COL, LNAME_COL, FNAME_COL]
    if CENTER_COL: cols.append(CENTER_COL)
    cols += [DATE_COL, GEN_COL, DET_COL]
    if AUTH_COL: cols.append(AUTH_COL)
    cols += [RES_COL, "Counts for PIR", "Reason (if not counted)"]

    details = df[cols].copy()
    rename_map = {c: _clean_header(c) for c in details.columns if c not in ["Counts for PIR", "Reason (if not counted)"]}
    details.rename(columns=rename_map, inplace=True)

    date_out = _clean_header(DATE_COL)
    pid_out = _clean_header(PID_COL)
    details[date_out] = _parse_to_dt(details[date_out]).dt.strftime("%m/%d/%y")
    details[pid_out] = details[pid_out].apply(_format_pid_display)
    details = details.fillna("")

    # =========================
    # PIR Summary (CRITICAL FIX: use string keys for distinct counts)
    # =========================
    pir_rows = df[df["Counts for PIR"] == "Yes"].copy()

    # Distinct children by (General, Detail, PIR code) using _PID_KEY
    per_child = (
        pir_rows
        .drop_duplicates(subset=["_PID_KEY", GEN_COL, "_PIR_CODE"])
        .groupby([GEN_COL, DET_COL])["_PID_KEY"]
        .nunique()
        .rename("Distinct Children (PID)")
        .reset_index()
    )

    # Distinct families by (General, Detail, PIR code) using _FID_KEY
    per_family = (
        pir_rows
        .drop_duplicates(subset=["_FID_KEY", GEN_COL, "_PIR_CODE"])
        .groupby([GEN_COL, DET_COL])["_FID_KEY"]
        .nunique()
        .rename("PIR (Distinct Families)")
        .reset_index()
    )

    summary = per_child.merge(per_family, on=[GEN_COL, DET_COL], how="outer").fillna(0)
    summary.rename(columns={GEN_COL: "GENERAL service", DET_COL: "DETAILED services"}, inplace=True)
    summary = summary[["GENERAL service", "DETAILED services", "Distinct Children (PID)", "PIR (Distinct Families)"]]

    # =========================
    # Author Fix List (actionable)
    # =========================
    author_col_name = _clean_header(AUTH_COL) if AUTH_COL else None
    actionable = {"Missing General Service", "Missing Detailed Service", "Invalid/Missing Result", "Missing Service Date"}

    fix_rows = details[(details["Counts for PIR"] == "No") & (details["Reason (if not counted)"].isin(actionable))].copy()
    if author_col_name and author_col_name in fix_rows.columns:
        pids_by_group = (
            fix_rows
            .groupby([author_col_name, "Reason (if not counted)"])[_clean_header(PID_COL)]
            .apply(lambda s: ", ".join(sorted({_format_pid_display(x) for x in s if str(x).strip() != ""})))
        )
        author_fix = pids_by_group.reset_index().rename(columns={_clean_header(PID_COL): "PIDs to Fix"})
        author_fix["Count of PIDs"] = author_fix["PIDs to Fix"].apply(lambda x: 0 if x == "" else len([p for p in x.split(", ") if p]))
    else:
        author_fix = pd.DataFrame(columns=[author_col_name or "Author", "Reason (if not counted)", "PIDs to Fix", "Count of PIDs"])

    # =========================
    # PIS Dashboard (General Services by Month)
    # =========================
    gen_month = (
        df.groupby([df["_month"], GEN_COL])
          .agg(Services=("Family ID", "count"))
          .reset_index()
          .rename(columns={"_month": "Month", GEN_COL: "GENERAL service"})
    )
    gen_month["Month"] = pd.to_datetime(gen_month["Month"]).dt.strftime("%b %Y")

    monthly = (
        df.groupby(df["_month"])
          .size()
          .rename("Services")
          .reset_index()
          .rename(columns={"_month": "Month"})
    )

    # =========================
    # Excel Writer & Formatting
    # =========================
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book
        hdr_fmt     = wb.add_format({"bold": True, "font_color": "white", "bg_color": HCHSP_NAVY, "align": "center", "valign": "vcenter", "text_wrap": True, "border": 1})
        border_all  = wb.add_format({"border": 1})
        title_fmt   = wb.add_format({"bold": True, "font_size": 18, "align": "center", "font_color": HCHSP_NAVY})
        subtitle_fmt= wb.add_format({"bold": True, "font_size": 12, "align": "center"})
        red_fmt     = wb.add_format({"bold": True, "font_size": 12, "font_color": HCHSP_RED})
        light_red   = wb.add_format({"bg_color": "#F8D7DA"})
        bold_center = wb.add_format({"bold": True, "align": "center"})
        kpi_lbl     = wb.add_format({"bold": True, "align": "center", "valign": "vcenter", "font_size": 12, "bg_color": HCHSP_LIGHT, "border": 1})
        kpi_val     = wb.add_format({"bold": True, "align": "center", "valign": "vcenter", "font_size": 16, "bg_color": "#FFFFFF", "border": 1})
        yellow_total= wb.add_format({"bold": True, "bg_color": "#FFF2CC", "border": 1, "align": "center", "font_size": 14})
        total_fmt   = wb.add_format({"bold": True, "bg_color": "#E2EFDA", "border": 1})
        c44_fmt     = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})

        now_ct = datetime.now(ZoneInfo("America/Chicago")).strftime("%m/%d/%y %I:%M %p CT")

        def _set_widths(ws, cols):
            for idx, name in enumerate(cols):
                n = str(name).lower()
                w = 20
                if "center" in n or "campus" in n: w = 26
                if "general service" in n: w = 30
                if "detail service" in n or "detailed service" in n: w = 40
                if "date" in n: w = 14
                if "result" in n: w = 20
                if "author" in n: w = 24
                if "pid" in n or "family id" in n: w = 16
                if "last name" in n or "first name" in n: w = 22
                if "counts for pir" in n: w = 18
                if "reason" in n: w = 34
                ws.set_column(idx, idx, w)

        # ---- Services & Referrals sheet
        details.to_excel(writer, index=False, sheet_name="Services & Referrals", startrow=3)
        ws1 = writer.sheets["Services & Referrals"]
        ws1.hide_gridlines(0); ws1.set_row(0,24); ws1.set_row(1,22); ws1.set_row(2,20)

        last_col_0 = details.shape[1] - 1
        if logo_path.exists():
            ws1.set_column(0, 0, 16)
            ws1.insert_image(0, 0, str(logo_path),
                             {"x_offset":2, "y_offset":2, "x_scale":0.53, "y_scale":0.53, "object_position": 1})
        ws1.merge_range(0, 1, 0, last_col_0, "Hidalgo County Head Start Program", title_fmt)
        ws1.merge_range(1, 1, 1, last_col_0, "", subtitle_fmt)
        ws1.write_rich_string(1, 1, subtitle_fmt, "Services & Referrals — 2025-2026 as of ", red_fmt, f"({now_ct})", subtitle_fmt)

        ws1.set_row(3,26)
        for c, col in enumerate(details.columns):
            ws1.write(3, c, col, hdr_fmt)
        ws1.freeze_panes(4, 0)

        last_row_0 = len(details) + 3
        ws1.autofilter(3, 0, last_row_0, last_col_0)
        ws1.conditional_format(3, 0, last_row_0, last_col_0, {"type":"no_errors","format":border_all})
        _set_widths(ws1, details.columns)

        # Highlight missing key fields
        name_to_idx = {name: idx for idx, name in enumerate(details.columns)}
        for name in details.columns:
            if ("general service" in name.lower() or "detail service" in name.lower()
                or "author" in name.lower() or "center" in name.lower() or "campus" in name.lower()):
                ws1.conditional_format(4, name_to_idx[name], last_row_0, name_to_idx[name],
                                       {"type":"blanks","format": light_red})

        # Helper column for SUBTOTAL
        helper_idx = last_col_0 + 1
        ws1.write(3, helper_idx, "_helper_")
        for r in range(4, last_row_0 + 1):
            ws1.write_number(r, helper_idx, 1)
        ws1.set_column(helper_idx, helper_idx, None, None, {"hidden":1})

        # Totals row
        totals_row = last_row_0 + 1
        ws1.write(totals_row, 0, "Total", wb.add_format({"bold":True,"align":"right"}))
        helper_col_letter = _col_letter(helper_idx)
        headers = list(details.columns)
        try:
            gs_idx = next(i for i, h in enumerate(headers) if "general service" in h.lower())
        except StopIteration:
            gs_idx = 5
        ws1.write_formula(totals_row, gs_idx,
                          f"=SUBTOTAL(109,{helper_col_letter}5:{helper_col_letter}{last_row_0+1})",
                          bold_center)

        # ---- PIR Summary sheet
        summary.to_excel(writer, index=False, sheet_name="PIR Summary", startrow=1)
        ws2 = writer.sheets["PIR Summary"]
        ws2.hide_gridlines(0); ws2.set_row(0,24)

        if logo_path.exists():
            ws2.set_column(0, 0, 16)
            ws2.insert_image(0, 0, str(logo_path),
                             {"x_offset":2, "y_offset":2, "x_scale":0.53, "y_scale":0.53, "object_position": 1})
        ws2.write(0, 1, "PIR Summary", wb.add_format({"bold":True,"font_size":14,"align":"left","font_color":HCHSP_NAVY}))

        ws2.set_row(1,26)
        for c, col in enumerate(summary.columns):
            ws2.write(1, c, col, hdr_fmt)
        last_row2 = len(summary) + 1
        last_col2 = len(summary.columns) - 1
        ws2.autofilter(1, 0, last_row2, last_col2)
        ws2.conditional_format(1, 0, last_row2, last_col2, {"type":"no_errors","format": border_all})

        def _set_widths2(ws, cols):
            for idx, name in enumerate(cols):
                n = str(name).lower()
                w = 24
                if "detailed services" in n: w = 44
                ws.set_column(idx, idx, w)
        _set_widths2(ws2, summary.columns)

        # Dynamic SUBTOTALS block (respects filters)
        start_excel_row = 3
        end_excel_row = last_row2 + 1
        children_col = _col_letter(2)
        families_col = _col_letter(3)
        ws2.write(last_row2 + 2, 1, "Dynamic Totals", total_fmt)
        ws2.write_formula(last_row2 + 2, 2, f"=SUBTOTAL(109,{children_col}{start_excel_row}:{children_col}{end_excel_row})", total_fmt)
        ws2.write_formula(last_row2 + 2, 3, f"=SUBTOTAL(109,{families_col}{start_excel_row}:{families_col}{end_excel_row})", total_fmt)

        # C.44 grand total (ignores filters on purpose)
        ws2.write(last_row2 + 3, 1, "C.44 – Sum of PIR Families (TOTAL)", c44_fmt)
        ws2.write_formula(last_row2 + 3, 3, f"=SUM({families_col}{start_excel_row}:{families_col}{end_excel_row})", c44_fmt)

        # ---- Author Fix List
        author_fix.to_excel(writer, index=False, sheet_name="Author Fix List", startrow=1)
        ws3 = writer.sheets["Author Fix List"]
        ws3.hide_gridlines(0); ws3.set_row(0,24)
        if logo_path.exists():
            ws3.set_column(0, 0, 16)
            ws3.insert_image(0, 0, str(logo_path),
                             {"x_offset":2, "y_offset":2, "x_scale":0.53, "y_scale":0.53, "object_position": 1})
        ws3.write(0, 1, "Author Fix List (Actionable only)", wb.add_format({"bold":True,"font_size":14,"align":"left","font_color":HCHSP_NAVY}))
        ws3.set_row(1,26)
        for c, col in enumerate(author_fix.columns):
            ws3.write(1, c, col, hdr_fmt)
        ws3.autofilter(1, 0, len(author_fix) + 1, len(author_fix.columns) - 1)
        ws3.conditional_format(1, 0, len(author_fix) + 1, len(author_fix.columns) - 1, {"type":"no_errors","format": border_all})
        for idx, name in enumerate(author_fix.columns):
            w = 22
            if "reason" in name.lower(): w = 30
            if "pids" in name.lower(): w = 50
            ws3.set_column(idx, idx, w)

        # ---- PIR Dashboard (KPIs + Top Detailed Services)
        ws4 = wb.add_worksheet("PIR Dashboard")
        ws4.hide_gridlines(0); ws4.set_row(0,24)
        if logo_path.exists():
            ws4.set_column(0, 0, 16)
            ws4.insert_image(0, 0, str(logo_path),
                             {"x_offset":2, "y_offset":2, "x_scale":0.53, "y_scale":0.53, "object_position": 1})
        ws4.merge_range(0, 1, 0, 6, "PIR Dashboard", title_fmt)

        k_children = int(summary["Distinct Children (PID)"].sum()) if len(summary) else 0
        k_families = int(summary["PIR (Distinct Families)"].sum()) if len(summary) else 0
        k_details = int(summary.shape[0]) if len(summary) else 0

        ws4.merge_range(2, 1, 3, 2, "PIR Children", kpi_lbl); ws4.merge_range(4, 1, 5, 2, k_children, kpi_val)
        ws4.merge_range(2, 3, 3, 4, "PIR Families", kpi_lbl); ws4.merge_range(4, 3, 5, 4, k_families, kpi_val)
        ws4.merge_range(2, 5, 3, 6, "Number of Detailed Services", kpi_lbl); ws4.merge_range(4, 5, 5, 6, k_details, kpi_val)

        top_det = summary.sort_values("PIR (Distinct Families)", ascending=False).reset_index(drop=True)
        start_r, start_c = 8, 1
        top_det.to_excel(writer, index=False, sheet_name="PIR Dashboard", startrow=start_r, startcol=start_c)
        for c, col in enumerate(top_det.columns):
            ws4.write(start_r, start_c + c, col, hdr_fmt)
        end_r = start_r + len(top_det)

        chart1 = wb.add_chart({"type":"column"})
        chart1.set_title({"name":"Top Detailed Services by PIR Families"})
        chart1.set_y_axis({"name":"Families"})
        chart1.add_series({
            "name":      ["PIR Dashboard", start_r, start_c + 3],
            "categories":["PIR Dashboard", start_r + 1, start_c + 1, end_r, start_c + 1],
            "values":    ["PIR Dashboard", start_r + 1, start_c + 3, end_r, start_c + 3],
            "fill":      {"color": HCHSP_NAVY},
            "border":    {"color": HCHSP_NAVY}
        })
        chart1.set_size({"width":820,"height":360})
        ws4.insert_chart(2, 8, chart1)

        # ---- PIS Dashboard (General Services)
        ws5 = wb.add_worksheet("PIS Dashboard")
        ws5.hide_gridlines(0); ws5.set_row(0,24)
        if logo_path.exists():
            ws5.set_column(0, 0, 16)
            ws5.insert_image(0, 0, str(logo_path),
                             {"x_offset":2, "y_offset":2, "x_scale":0.53, "y_scale":0.53, "object_position": 1})
        ws5.merge_range(0, 1, 0, 14, "PIS Dashboard (General Services)", title_fmt)

        gen_month_sorted = gen_month.sort_values(["Month", "Services"], ascending=[True, False]).reset_index(drop=True)
        gen_month_sorted = gen_month_sorted[["Month", "GENERAL service", "Services"]]
        gs_r, gs_c = 2, 1
        gen_month_sorted.to_excel(writer, index=False, sheet_name="PIS Dashboard", startrow=gs_r, startcol=gs_c)
        for c, col in enumerate(gen_month_sorted.columns):
            ws5.write(gs_r, gs_c + c, col, hdr_fmt)
        gs_end_r = gs_r + len(gen_month_sorted)

        ws5.add_table(gs_r, gs_c, gs_end_r, gs_c + len(gen_month_sorted.columns) - 1,
                      {"name":"tblGenMonth","columns":[{"header":h} for h in gen_month_sorted.columns],"style":"Table Style Medium 2"})

        services_col_letter = _col_letter(gs_c + 2)
        start_excel_row = gs_r + 2
        end_excel_row = gs_end_r + 1
        ws5.write(gs_r - 1, gs_c + 1, "Total Services ➜", yellow_total)
        ws5.write_formula(gs_r - 1, gs_c + 2,
                          f"=SUBTOTAL(109,{services_col_letter}{start_excel_row}:{services_col_letter}{end_excel_row})",
                          yellow_total)

        chart2 = wb.add_chart({"type":"column"})
        chart2.set_title({"name":"Services and Referrals"})
        chart2.set_y_axis({"name":"Total Services"})
        chart2.add_series({
            "name":      ["PIS Dashboard", gs_r, gs_c + 2],
            "categories":["PIS Dashboard", gs_r + 1, gs_c + 1, gs_end_r, gs_c + 1],
            "values":    ["PIS Dashboard", gs_r + 1, gs_c + 2, gs_end_r, gs_c + 2],
            "fill":      {"color": HCHSP_NAVY},
            "border":    {"color": HCHSP_NAVY},
            "data_labels":{"value":True,"position":"outside_end","font":{"bold":True,"size":14}}
        })
        chart2.set_size({"width":760,"height":320})
        ws5.insert_chart(4, 7, chart2)

        ws5.write(gs_end_r + 3, gs_c + 1, "Totals ➜", wb.add_format({"bold":True,"align":"right"}))
        ws5.write_formula(gs_end_r + 3, gs_c + 2,
                          f"=SUBTOTAL(109,{services_col_letter}{start_excel_row}:{services_col_letter}{end_excel_row})",
                          wb.add_format({"bold":True,"border":1}))

        # Result distribution
        res_counts = df[RES_COL].value_counts().reset_index()
        res_counts.columns = ["Result", "Count"]
        rc_r, rc_c = gs_end_r + 8, 1
        res_counts.to_excel(writer, index=False, sheet_name="PIS Dashboard", startrow=rc_r, startcol=rc_c)
        for c, col in enumerate(res_counts.columns):
            ws5.write(rc_r, rc_c + c, col, hdr_fmt)
        rc_end_r = rc_r + len(res_counts)

        chart3 = wb.add_chart({"type":"pie"})
        points = [{"data_labels":{"percentage":True,"value":False,"position":"outside_end","leader_lines":True,"font":{"bold":True,"size":14,"color":"black"}}}
                  for _ in res_counts["Count"].tolist()]
        chart3.add_series({
            "name":"Results",
            "categories":["PIS Dashboard", rc_r + 1, rc_c + 0, rc_end_r, rc_c + 0],
            "values":    ["PIS Dashboard", rc_r + 1, rc_c + 1, rc_end_r, rc_c + 1],
            "points": points
        })
        chart3.set_title({"name":"Result Distribution"})
        chart3.set_size({"width":520,"height":340})
        ws5.insert_chart(rc_r, rc_c + 5, chart3)

        # Monthly trend
        monthly_sorted = monthly.sort_values("Month").reset_index(drop=True)
        monthly_sorted["Month"] = monthly_sorted["Month"].dt.strftime("%b %Y")
        m_r, m_c = rc_end_r + 8, 1
        monthly_sorted.to_excel(writer, index=False, sheet_name="PIS Dashboard", startrow=m_r, startcol=m_c)
        for c, col in enumerate(monthly_sorted.columns):
            ws5.write(m_r, m_c + c, col, hdr_fmt)
        m_end_r = m_r + len(monthly_sorted)

        chart4 = wb.add_chart({"type":"line"})
        chart4.set_title({"name":"Monthly Services Trend"})
        chart4.set_y_axis({"name":"Services"})
        chart4.add_series({
            "name":      ["PIS Dashboard", m_r, m_c + 1],
            "categories":["PIS Dashboard", m_r + 1, m_c + 0, m_end_r, m_c + 0],
            "values":    ["PIS Dashboard", m_r + 1, m_c + 1, m_end_r, m_c + 1],
            "marker":{"type":"circle"},
            "data_labels":{"value":True,"font":{"bold":True,"size":14}}
        })
        chart4.set_size({"width":980,"height":360})
        ws5.insert_chart(m_r, m_c + 5, chart4)

    return buf.getvalue()

# =========================
# Run
# =========================
if process and sref_file:
    try:
        # IMPORTANT: The 10433 raw file’s real headers start on row 5 (0-based index 4)
        raw = pd.read_excel(sref_file, header=4)
        xlsx = build_workbook(raw, cutoff, require_pir=st.session_state.get("require_pir", True))
        st.download_button(
            "Download Styled Workbook (Excel)",
            data=xlsx,
            file_name="HCHSP_Services_Referrals_PIR.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.success("Workbook generated.")
    except Exception as e:
        st.error(f"Processing error: {e}")
