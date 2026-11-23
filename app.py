"""
CBR Report Builder  (Streamlit)

Upload 3 Excel files  ->  get the finished CBR report back:
  1. CBR Monthly report  (last 6 months, downloaded from the CSD portal)
  2. CBR Report template (the blank template)
  3. Customer price file (sent by the client)

Nothing is stored. Files are processed in memory and are gone when the page closes.
"""
import io
import re
from collections import Counter
from copy import copy
from datetime import datetime, date

import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.formatting.rule import CellIsRule, Rule
from openpyxl.formula.translate import Translator
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles.cell_style import StyleArray
from openpyxl.utils import get_column_letter

FIRST_DATA_ROW = 3          # template data starts on row 3 (header is row 2)
MONTH_COLS = range(8, 14)   # H..M
YELLOW = PatternFill("solid", fgColor="FFFF00")


# ----------------------------------------------------------------------------
# helpers
# ----------------------------------------------------------------------------
def to_number_if_possible(v):
    """Step 4: 'Convert to numbers' for Product Number columns (C and D).
    Text such as '10053413' becomes the number 10053413.
    IDs with leading zeros (e.g. '00123') are left as text so they are not damaged."""
    if isinstance(v, str):
        s = v.strip()
        if re.fullmatch(r"[1-9]\d*|0", s):
            return int(s)
        return s
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def norm_key(v):
    """Key used to match product numbers (case-insensitive like VLOOKUP, number == 'number')."""
    v = to_number_if_possible(v)
    if v is None:
        return ""
    return str(v).strip().upper()


def excel_sort_key(v):
    """Mimics Excel sort: numbers first, then text A-Z (Excel ignores hyphens)."""
    if isinstance(v, (int, float)):
        return (0, v, "")
    s = "" if v is None else str(v).upper().replace("-", "").replace("'", "")
    return (1, 0, s)


def first_of_month(v):
    d = pd.to_datetime(v)
    return datetime(d.year, d.month, 1)


def find_total_row(ws):
    """Row that holds '*Total excludes Market Price Items*'."""
    for row in ws.iter_rows(min_row=FIRST_DATA_ROW):
        for c in row:
            if isinstance(c.value, str) and "excludes" in c.value.lower():
                return c.row
    raise ValueError("Template: could not find the row with '*Total excludes Market Price Items*'.")


def suggested_max(month_values, max_qty):
    """Python copy of the template formula in column R (Suggested Max)."""
    vals = sorted([float(x or 0) for x in month_values], reverse=True)[:3]
    avg = sum(vals) / 3.0
    s = -(-avg // 1)  # ROUNDUP(avg, 0) for positive numbers
    s = int(s)
    try:
        if max_qty and -0.05 < (s - max_qty) / max_qty < 0.05:
            return max_qty
    except TypeError:
        pass
    return s


# ----------------------------------------------------------------------------
# reading the inputs
# ----------------------------------------------------------------------------
def read_monthly(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.worksheets[0]
    head = [str(ws.cell(1, c).value or "").strip().lower() for c in range(1, 8)]
    if "product number" not in head[2] or "warehouse" not in head[0]:
        raise ValueError(
            "The 'CBR Monthly report' file does not look right. Expected headers "
            "'Warehouse' in A1 and 'Product Number' in C1."
        )
    months = [first_of_month(ws.cell(1, c).value) for c in MONTH_COLS]     # Step 3
    rows = []
    for r in range(2, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, 17)]                 # A..P only
        if vals[2] in (None, "") and vals[0] in (None, ""):
            continue
        vals[2] = to_number_if_possible(vals[2])                           # Step 4 (C)
        vals[3] = to_number_if_possible(vals[3])                           # Step 4 (D)
        rows.append(vals)
    if not rows:
        raise ValueError("The 'CBR Monthly report' has no data rows.")
    return months, rows


def read_prices(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb["Customer Price File"] if "Customer Price File" in wb.sheetnames else None
    if ws is None:
        for cand in wb.worksheets:
            if str(cand.cell(6, 1).value or "").strip().lower() == "product id":
                ws = cand
                break
    if ws is None:
        raise ValueError("Price file: could not find the 'Customer Price File' sheet.")
    start = 7
    for r in range(1, 20):
        if str(ws.cell(r, 1).value or "").strip().lower() == "product id":
            start = r + 1
            break
    prices = {}
    for r in range(start, ws.max_row + 1):
        pid = ws.cell(r, 1).value
        if pid in (None, ""):
            continue
        k = norm_key(pid)
        if k not in prices:                      # VLOOKUP returns the first match
            prices[k] = (ws.cell(r, 3).value, ws.cell(r, 4).value)
    return prices


# ----------------------------------------------------------------------------
# main processing
# ----------------------------------------------------------------------------
def build_report(template_bytes, monthly_bytes, price_bytes, ascending=True):
    months, rows = read_monthly(monthly_bytes)
    prices = read_prices(price_bytes)

    # Step 4: sort on Product number (column D)
    rows.sort(key=lambda v: excel_sort_key(v[3] if v[3] not in (None, "") else v[2]),
              reverse=not ascending)

    # Step 5: duplicates in product number (column C)
    counts = Counter(norm_key(v[2]) for v in rows)
    duplicates = sorted(k for k, n in counts.items() if n > 1)

    # Step 7: new items (added within the 6 months) flagged "Remove" -> Q,R = O,P
    window_start = min(months)
    step7_rows = []
    for v in rows:
        added = v[6]
        if isinstance(added, datetime) and added >= window_start:
            if suggested_max(v[7:13], v[15]) == 0:
                step7_rows.append(v[2])

    # ---- write into the template
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    ws = wb["CBR"] if "CBR" in wb.sheetnames else wb.worksheets[0]
    max_col = 28  # A..AB  (AC "Delete Column" is removed, as in the finished sample)
    total_row = find_total_row(ws)
    capacity = total_row - FIRST_DATA_ROW
    n = len(rows)

    # remember the total row (styles + values) and the data-row formulas/styles
    total_cells = [(copy(ws.cell(total_row, c)._style), ws.cell(total_row, c).value)
                   for c in range(1, max_col + 1)]
    style_row = total_row - 1                       # an untouched formatted data row
    row_styles = [copy(ws.cell(style_row, c)._style) for c in range(1, max_col + 1)]
    formula_cols = {}
    for c in range(1, max_col + 1):
        val = ws.cell(FIRST_DATA_ROW, c).value
        if isinstance(val, str) and val.startswith("="):
            formula_cols[c] = val

    for mr in list(ws.merged_cells.ranges):
        if mr.min_row == total_row:
            ws.unmerge_cells(str(mr))

    # clear everything from the first row that will be unused, down to the old total row
    last_data_row = FIRST_DATA_ROW + n - 1
    new_total_row = last_data_row + 1
    for r in range(new_total_row, max(total_row, new_total_row) + 1):
        for c in range(1, 30):
            cell = ws.cell(r, c)
            cell.value = None
            cell._style = StyleArray()   # blank default style

    # make sure every data row has styles + formulas (also when there are more rows than the template)
    for r in range(FIRST_DATA_ROW, last_data_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(r, c)
            if r >= total_row or cell.value is None and c in formula_cols:
                cell._style = copy(row_styles[c - 1])
                if c in formula_cols:
                    cell.value = Translator(formula_cols[c], origin=f"{get_column_letter(c)}{FIRST_DATA_ROW}")\
                        .translate_formula(f"{get_column_letter(c)}{r}")

    # Step 3: months into H2:M2
    for i, c in enumerate(MONTH_COLS):
        ws.cell(2, c).value = months[i]

    # Steps 2 + 6: data A..P into row 3 onward;  Step 9: prices into Y / Z
    missing = []
    preview = []
    step7_keys = {norm_key(x) for x in step7_rows}
    for i, v in enumerate(rows):
        r = FIRST_DATA_ROW + i
        for c in range(1, 17):
            ws.cell(r, c).value = v[c - 1]
        cur, new = prices.get(norm_key(v[2]), (None, None))
        found = norm_key(v[2]) in prices
        ws.cell(r, 25).value = cur            # Y  Current Price  (values, not VLOOKUP)
        ws.cell(r, 26).value = new            # Z  New Price
        if not found:
            ws.cell(r, 25).fill = YELLOW
            ws.cell(r, 26).fill = YELLOW
            missing.append(v[2])
        if norm_key(v[2]) in step7_keys:      # Step 7
            ws.cell(r, 17).value = v[14]      # Q = O
            ws.cell(r, 18).value = v[15]      # R = P
        preview.append({
            "Product Number": v[2], "Description": v[4], "6-Month Usage": v[13],
            "Min": v[14], "Max": v[15], "Current Price": cur, "New Price": new,
            "Price found": "Yes" if found else "NO",
        })

    # Step 8: total row goes straight after the last product row
    for c in range(1, max_col + 1):
        style, value = total_cells[c - 1]
        cell = ws.cell(new_total_row, c)
        cell._style = copy(style)
        cell.value = value
    ws.cell(new_total_row, 28).value = f"=SUM(AB{FIRST_DATA_ROW}:AB{last_data_row})"
    ws.merge_cells(start_row=new_total_row, start_column=25, end_row=new_total_row, end_column=27)

    # remove column AC ("Delete Column" / notes) like the finished sample
    for r in (1, 2):
        ws.cell(r, 29).value = None
        ws.cell(r, 29)._style = StyleArray()
    ws.auto_filter.ref = f"A2:{get_column_letter(max_col)}2"

    # conditional formats: yellow "Remove" in W, red duplicates in C (as in the finished sample)
    from openpyxl.formatting.formatting import ConditionalFormattingList
    ws.conditional_formatting = ConditionalFormattingList()
    ws.conditional_formatting.add(
        f"W{FIRST_DATA_ROW}:W{last_data_row}",
        CellIsRule(operator="equal", formula=['"Remove"'], fill=PatternFill(bgColor="FFFF00", fill_type="solid")))
    dup = Rule(type="duplicateValues",
               dxf=DifferentialStyle(font=Font(color="9C0006"),
                                     fill=PatternFill(bgColor="FFC7CE", fill_type="solid")))
    ws.conditional_formatting.add(f"C{FIRST_DATA_ROW}:C{last_data_row}", dup)

    wb.calculation.fullCalcOnLoad = True      # Excel recalculates all formulas when opened
    out = io.BytesIO()
    wb.save(out)

    crib = rows[0][0]
    info = {
        "rows": n, "crib": crib, "months": months,
        "duplicates": duplicates, "missing": missing,
        "step7": step7_rows, "preview": pd.DataFrame(preview),
        "extra_rows_removed": max(capacity - n, 0),
    }
    return out.getvalue(), info


# ----------------------------------------------------------------------------
# Streamlit page
# ----------------------------------------------------------------------------
def main():
    st.set_page_config(page_title="CBR Report Builder", page_icon="📊", layout="wide")
    st.title("📊 CBR Report Builder")
    st.caption("Upload the 3 Excel files, click **Create report**, download the finished CBR report. "
               "Nothing is saved - files only live in memory while you use this page.")

    c1, c2, c3 = st.columns(3)
    monthly = c1.file_uploader("1️⃣ CBR Monthly report (from CSD portal)", type=["xlsx"], key="m")
    template = c2.file_uploader("2️⃣ CBR Report template", type=["xlsx"], key="t")
    price = c3.file_uploader("3️⃣ Customer price file (from client)", type=["xlsx"], key="p")

    sort_choice = st.radio(
        "Sort order on Product number (column D)",
        ["A → Z (numbers first, then letters)  - same as your sample report",
         "Z → A (reverse)"],
        horizontal=True)

    if st.button("Create report", type="primary", disabled=not (monthly and template and price)):
        try:
            with st.spinner("Building the report..."):
                data, info = build_report(template.getvalue(), monthly.getvalue(), price.getvalue(),
                                          ascending=sort_choice.startswith("A"))
            st.session_state["result"] = (data, info)
        except Exception as e:                                   # show a readable message
            st.session_state.pop("result", None)
            st.error(f"Could not build the report: {e}")

    if "result" in st.session_state:
        data, info = st.session_state["result"]
        crib = str(info["crib"]).strip()
        default_name = f"{crib}_Crib_Business_Report_{date.today():%m-%d-%Y}.xlsx"
        st.success(f"Report ready: {info['rows']} products for crib **{crib}**. "
                   f"Months: {', '.join(m.strftime('%b %Y') for m in info['months'])}.")
        fname = st.text_input("File name", value=default_name)
        if not fname.lower().endswith(".xlsx"):
            fname += ".xlsx"
        st.download_button("⬇️ Download finished report", data=data, file_name=fname,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary")

        st.subheader("Checks")
        if info["duplicates"]:
            st.warning(f"Step 5 - duplicate Product Numbers ({len(info['duplicates'])}): "
                       + ", ".join(map(str, info["duplicates"]))
                       + ".  They are shown in red in column C of the report.")
        else:
            st.info("Step 5 - no duplicate Product Numbers.")
        if info["missing"]:
            st.error(f"Step 9 - {len(info['missing'])} product(s) NOT found in the price file "
                     "(Current/New Price left empty and highlighted yellow): "
                     + ", ".join(map(str, info["missing"])))
        else:
            st.info("Step 9 - every product was found in the price file.")
        if info["step7"]:
            st.warning("Step 7 - added in the last 6 months and flagged 'Remove'; "
                       "Suggested Min/Max set equal to Min/Max for: "
                       + ", ".join(map(str, info["step7"])))
        else:
            st.info("Step 7 - no products added in the last 6 months were flagged 'Remove'.")

        with st.expander("Preview of rows"):
            st.dataframe(info["preview"], use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
