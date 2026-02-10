# QR.py — Full Updated Version With:
# - Collapsible History Maintenance
# - Global Search + Column Filters (Option C)
# - Row Selection + Delete
# - Logic Editor that updates Database + History + Recalculates All Fields

import streamlit as st
import pandas as pd
import datetime
import os
from io import BytesIO
from pathlib import Path
import re
from typing import Optional, Dict, Any
from openpyxl import load_workbook

# =========================================
# COMPATIBILITY RERUN HELPER
# =========================================
def _safe_rerun():
    if hasattr(st, "rerun"):
        st.rerun()
    elif hasattr(st, "experimental_rerun"):
        st.experimental_rerun()

# =========================================
# PATH SETUP
# =========================================
try:
    BASE_DIR = Path(__file__).resolve().parent
except NameError:
    BASE_DIR = Path(os.getcwd()).resolve()

APP_DB_FILE = BASE_DIR / "Database.xlsx"
APP_LOGIC_FILE = BASE_DIR / "Logic.xlsx"
APP_HISTORY_FILE = BASE_DIR / "History.xlsx"
APP_BARCODE_FILE = BASE_DIR / "barcode.xlsx"

# =========================================
# NORMALIZATION HELPERS
# =========================================
ARABIC_TO_ASCII = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")

def normalize_material(value: str) -> str:
    if pd.isna(value):
        return ""
    s = str(value).strip().translate(ARABIC_TO_ASCII)
    s = re.sub(r"\s+", "", s)
    digits = re.findall(r"\d", s)
    return "".join(digits)

def normalize_material_series(series: pd.Series) -> pd.Series:
    return series.astype(str).map(normalize_material)

# =========================================
# FILE HELPERS
# =========================================
def save_df_to_excel(df: pd.DataFrame, path: Path, sheet: str = "Sheet1"):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet)
    Path(path).write_bytes(out.getvalue())

def probe_file_access(path: Path) -> dict:
    info = {"exists": False, "readable": False, "error": None,
            "abs_path": str(path.resolve()), "size": 0}
    try:
        if path.exists():
            info["exists"] = True
            info["size"] = path.stat().st_size
            with open(path, "rb") as fh:
                _ = fh.read(1024)
            info["readable"] = True
    except Exception as e:
        info["error"] = f"{type(e).__name__}: {e}"
        info["readable"] = False
    return info

# =========================================
# READERS: A→B (Material → Code)
# =========================================
def read_first_sheet_ab(path: Optional[Path] = None, bytes_buf: Optional[bytes] = None):
    if path is None and bytes_buf is None:
        return pd.DataFrame(columns=["material", "symbol"])
    try:
        wb = load_workbook(filename=(BytesIO(bytes_buf) if bytes_buf else str(path)), data_only=True)
    except:
        return pd.DataFrame(columns=["material", "symbol"])

    ws = wb.worksheets[0]
    mapping = {}
    for row in ws.iter_rows(values_only=True):
        a_val = row[0] if row and len(row) > 0 else None
        b_val = row[1] if row and len(row) > 1 else None
        a = normalize_material("" if a_val is None else str(a_val))
        b = "" if b_val is None else str(b_val).strip()
        if a and b:
            mapping[a] = b

    if not mapping:
        return pd.DataFrame(columns=["material", "symbol"])

    df = pd.DataFrame({"material": list(mapping.keys()),
                       "symbol": list(mapping.values())})
    df = df.drop_duplicates(subset=["material"], keep="last")
    return df[["material", "symbol"]]

def read_barcode_ab(path: Optional[Path] = None):
    if path is None or not path.exists():
        return pd.DataFrame(columns=["material", "barcode"])
    try:
        wb = load_workbook(filename=str(path), data_only=True)
    except:
        return pd.DataFrame(columns=["material", "barcode"])

    ws = wb.worksheets[0]
    mapping = {}
    for row in ws.iter_rows(values_only=True):
        a_val = row[0]
        b_val = row[1] if len(row) > 1 else None
        a = normalize_material("" if a_val is None else str(a_val))
        b = "" if b_val is None else str(b_val).strip()
        if a and b:
            mapping[a] = b

    df = pd.DataFrame({"material": list(mapping.keys()),
                       "barcode": list(mapping.values())})
    df = df.drop_duplicates(subset=["material"], keep="last")
    return df

# =========================================
# HISTORY LOADER
# =========================================
def ensure_history_file():
    cols = [
        "CookerName","Quantity","MaterialCode","OrderNumber",
        "Month","Year","Y","x","symbol","code34","code_length",
        "الباركود الدولي","LastSerial"
    ]
    if not APP_HISTORY_FILE.exists():
        hist = pd.DataFrame(columns=cols)
        save_df_to_excel(hist, APP_HISTORY_FILE, sheet="HISTORY")
        return hist
    try:
        hist = pd.read_excel(APP_HISTORY_FILE, sheet_name="HISTORY", engine="openpyxl")
        for c in cols:
            if c not in hist.columns:
                hist[c] = pd.Series(dtype="object")
        hist["MaterialCode"] = normalize_material_series(hist["MaterialCode"])
        return hist[cols]
    except:
        hist = pd.DataFrame(columns=cols)
        save_df_to_excel(hist, APP_HISTORY_FILE, sheet="HISTORY")
        return hist

# =========================================
# LOGIC LOADER (Capacity, Family, Option, Color)
# =========================================
def load_logic():
    if not APP_LOGIC_FILE.exists():
        upl = st.file_uploader("Upload Logic.xlsx", type=["xlsx"])
        if upl is None:
            st.stop()
        APP_LOGIC_FILE.write_bytes(upl.getbuffer())

    try:
        xls = pd.ExcelFile(APP_LOGIC_FILE, engine="openpyxl")
    except:
        st.error("Cannot read Logic.xlsx")
        st.stop()

    logic = {}
    for sh in ["Capacity","Family","option","Color"]:
        try:
            df = pd.read_excel(xls, sheet_name=sh, dtype=str)
        except:
            st.error(f"Missing sheet: {sh}")
            st.stop()
        df = df.iloc[:, :2]
        df.columns = ["name","code"]
        df = df.dropna()
        logic[sh.lower()] = df
    return logic

# =========================================
# DATABASE LOADER A→B
# =========================================
def load_database_from_disk_or_upload():
    info = probe_file_access(APP_DB_FILE)
    fallback = None

    with st.sidebar:
        st.subheader("Diagnostics")
        st.code(
            f"exists: {info['exists']}\n"
            f"readable: {info['readable']}\n"
            f"path: {info['abs_path']}\n"
            f"size: {info['size']}"
        )
        repl = st.file_uploader("Replace Database.xlsx", type=["xlsx"])
        if repl:
            APP_DB_FILE.write_bytes(repl.getbuffer())
            _safe_rerun()

        if not info["readable"]:
            tmp = st.file_uploader("Upload Session Database.xlsx", type=["xlsx"])
            if tmp:
                fallback = tmp.getvalue()

    if fallback:
        df = read_first_sheet_ab(bytes_buf=fallback)
    else:
        df = read_first_sheet_ab(path=APP_DB_FILE)

    if df.empty:
        empty = pd.DataFrame(columns=["Material","symbol"])
        save_df_to_excel(empty, APP_DB_FILE, sheet="DATABASE")
        return pd.DataFrame(columns=["material","symbol"])
    return df

def save_database(df):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df.rename(columns={"material":"Material"}).to_excel(w, index=False, sheet_name="DATABASE")
    APP_DB_FILE.write_bytes(out.getvalue())

# =========================================
# BARCODE LOADER
# =========================================
def load_barcode():
    return read_barcode_ab(APP_BARCODE_FILE)

# =========================================
# LOGIC BUILD
# =========================================
def build_x_from_logic(sel: dict, maps: dict):
    c = maps["capacity"].get(sel["capacity"],"")
    f = maps["family"].get(sel["family"],"")
    o = maps["option"].get(sel["option"],"")
    col = maps["color"].get(sel["color"],"")
    return f"{c}{f}{o}{col}"

# =========================================
# SERIAL RECALCULATOR
# =========================================
def recompute_last_serial(df_hist: pd.DataFrame):
    df = df_hist.copy()
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0).astype(int)
    df = df.sort_index()
    if not df.empty:
        df["LastSerial"] = df.groupby(["Year","Y"])["Quantity"].cumsum()
    else:
        df["LastSerial"] = 0
    return df

# =========================================
# STREAMLIT PAGE SETUP
# =========================================
st.set_page_config(page_title="Cooker Production Code Builder", layout="wide")

history_df = ensure_history_file()
logic = load_logic()
barcode_df = load_barcode()
db = load_database_from_disk_or_upload()

st.session_state.history = history_df
st.session_state.logic = logic
st.session_state.barcode = barcode_df
st.session_state.db = db
st.session_state.setdefault("added_materials",set())

# =========================================
# HEADER
# =========================================
st.title("Cooker Production Code Builder")

col1, col2 = st.columns([2,1])

with col1:
    ptype = st.radio("Product type",
                     ["Gas-cooker","Built-in","CKD"],
                     horizontal=True)
    Y = {"Gas-cooker":"GC", "Built-in":"BI","CKD":"Ck"}[ptype]

with col2:
    today = datetime.date.today()
    month = st.selectbox("Production month",
                         list(range(1,13)),
                         index=today.month-1,
                         format_func=lambda m: datetime.date(2000,m,1).strftime("%b"))
    year = st.number_input("Production year",
                           min_value=2000,max_value=2100,
                           value=today.year)

st.markdown("---")
st.markdown("**Enter rows** → `CookerName`, `Quantity`, `SAP Material Code`, `OrderNumber`")

# =========================================
# INPUT GRID
# =========================================
default_df = pd.DataFrame({
    "CookerName":[""],
    "Quantity":[0],
    "MaterialCode":[""],
    "OrderNumber":[""]
})

df_input = st.data_editor(
    default_df,
    num_rows="dynamic",
    height=220,
    use_container_width=True
)

# CLEAN INPUT
df_input = df_input.copy()
df_input["MaterialCode"] = normalize_material_series(df_input["MaterialCode"])
df_input = df_input[df_input["MaterialCode"].str.len()>0]
df_input["CookerName"] = df_input["CookerName"].astype(str).str.strip()
df_input["OrderNumber"] = df_input["OrderNumber"].astype(str).str.strip()
df_input["Quantity"] = pd.to_numeric(df_input["Quantity"], errors="coerce").fillna(0).astype(int)

if not df_input.empty:
    st.subheader("Input Preview")
    st.dataframe(df_input, use_container_width=True)
# ==============================
# CONTINUATION — PART 2
# ==============================

    # Logic maps (capacity/family/option/color)
    logic_maps = {
        cat: dict(zip(
            st.session_state.logic[cat]["name"],
            st.session_state.logic[cat]["code"]
        ))
        for cat in ["capacity", "family", "option", "color"]
    }

    # DB mapping
    db = st.session_state.db
    db_materials = set(db["material"])
    db_map = dict(zip(db["material"], db["symbol"]))

    # Find missing and matched materials
    unique_materials = df_input["MaterialCode"].unique().tolist()
    missing_codes = [m for m in unique_materials if m not in db_materials]
    matched_codes = [m for m in unique_materials if m in db_materials]

    with st.expander("Matching diagnostics"):
        st.write(f"Matched materials: {len(matched_codes)}")
        if matched_codes:
            st.code("\n".join(matched_codes[:50]))

        st.write(f"Missing materials: {len(missing_codes)}")
        if missing_codes:
            st.code("\n".join(missing_codes[:50]))

        st.write("Sample of materials in Database.xlsx:")
        st.code("\n".join(list(db_materials)[:50]))

    # Collect logic for missing codes
    selections = {}
    if missing_codes:
        st.info(f"{len(missing_codes)} material codes not in the database — define logic below.")
        for mat in missing_codes:
            with st.expander(f"Define logic for: {mat}", expanded=True):
                selections[mat] = {
                    "capacity": st.selectbox("Capacity", st.session_state.logic["capacity"]["name"], key=f"{mat}_cap"),
                    "family":   st.selectbox("Family",   st.session_state.logic["family"]["name"],   key=f"{mat}_fam"),
                    "option":   st.selectbox("Option",   st.session_state.logic["option"]["name"],   key=f"{mat}_opt"),
                    "color":    st.selectbox("Color",    st.session_state.logic["color"]["name"],    key=f"{mat}_col"),
                }

    # ==========================================
    # MAIN GENERATE & SAVE ACTION
    # ==========================================
    if st.button("Generate & Save"):
        df_input["x"] = df_input["MaterialCode"].map(db_map)
        new_pairs = []

        # Build logic for missing codes
        for mat in missing_codes:
            sel = selections.get(mat)
            if not sel:
                st.error(f"Missing logic selections for: {mat}")
                st.stop()
            x_new = build_x_from_logic(sel, logic_maps)
            df_input.loc[df_input["MaterialCode"] == mat, "x"] = x_new
            new_pairs.append({"material": mat, "symbol": x_new})
            st.session_state.added_materials.add(mat)

        # Update DB if new codes added
        if new_pairs:
            new_df = pd.DataFrame(new_pairs)
            new_df["material"] = normalize_material_series(new_df["material"])
            db = pd.concat([db, new_df], ignore_index=True).drop_duplicates(
                subset=["material"], keep="last"
            )
            st.session_state.db = db
            save_database(db)
            st.success("Database.xlsx updated.")

            db_materials = set(db["material"])
            db_map = dict(zip(db["material"], db["symbol"]))

        # Add production symbol MMYY + Y
        df_input["Month"] = int(month)
        df_input["Year"] = int(year)
        df_input["Y"] = Y
        df_input["symbol"] = df_input.apply(
            lambda r: f"{int(r['Month']):02d}{r['Year'] % 100:02d}{r['Y']}",
            axis=1
        )

        # Add code34 and its length
        df_input["code34"] = df_input.apply(
            lambda r: f"{r['x']}&PO={r['OrderNumber']}&PC={r['MaterialCode']}",
            axis=1
        )
        df_input["code_length"] = df_input["code34"].str.len()

        # Add barcode
        barcode_map = dict(zip(st.session_state.barcode["material"],
                               st.session_state.barcode["barcode"]))
        df_input["الباركود الدولي"] = df_input["MaterialCode"].map(barcode_map).fillna("")

        # Append to history
        hist = pd.concat([st.session_state.history, df_input], ignore_index=True)
        hist = recompute_last_serial(hist)
        st.session_state.history = hist
        save_df_to_excel(hist, APP_HISTORY_FILE, sheet="HISTORY")

        st.success("History updated and saved.")

        # Show summary of batch
        st.subheader("Processed Entries")
        display_cols = [
            "CookerName", "Quantity", "MaterialCode", "OrderNumber",
            "symbol", "الباركود الدولي", "code34"
        ]
        st.dataframe(df_input[display_cols], use_container_width=True)

        # Download updated DB
        out_db = BytesIO()
        with pd.ExcelWriter(out_db, engine="openpyxl") as writer:
            st.session_state.db.rename(columns={"material":"Material"}).to_excel(
                writer, index=False, sheet_name="DATABASE"
            )
        out_db.seek(0)
        st.download_button(
            "⬇️ Download Updated Database.xlsx",
            out_db,
            file_name="Database_updated.xlsx"
        )

        # Download this batch
        out_batch = BytesIO()
        with pd.ExcelWriter(out_batch, engine="openpyxl") as writer:
            df_input.to_excel(writer, index=False, sheet_name="BATCH")
        out_batch.seek(0)
        st.download_button(
            "⬇️ Download This Batch",
            out_batch,
            file_name=f"Batch_{year}_{month:02d}.xlsx"
        )

# END OF PART 2
# ==============================
# PART 3 — HISTORY MAINTENANCE + LOGIC EDITOR
# ==============================

st.markdown("---")
st.header("History (Persistent View by Type)")

# Display history by type
hist = recompute_last_serial(st.session_state.history)
st.session_state.history = hist

tabs = st.tabs(["Gas-cooker (GC)", "Built-in (BI)", "CKD (Ck)"])
for y_val, tab in zip(["GC", "BI", "Ck"], tabs):
    with tab:
        filtered = hist[hist["Y"] == y_val]
        if filtered.empty:
            st.info(f"No entries for {y_val}.")
        else:
            st.dataframe(
                filtered[[
                    "CookerName","Quantity","MaterialCode","OrderNumber",
                    "symbol","الباركود الدولي","code34","LastSerial"
                ]],
                use_container_width=True,
                height=350
            )

# ==========================================================
# NEW: LOGIC EDITOR — UPDATES BOTH DB AND HISTORY
# ==========================================================
with st.expander("Edit Material Logic (Updates Database + History)", expanded=False):
    st.write("Enter a MaterialCode to edit its logic. This will update:")
    st.write("- Database.xlsx (material → new x)")
    st.write("- All History rows using this MaterialCode")
    st.write("- Recalculate x, symbol, code34, barcode, LastSerial")

    mat_to_edit = st.text_input("MaterialCode to edit:", "")
    mat_to_edit_norm = normalize_material(mat_to_edit)

    if mat_to_edit_norm:
        db = st.session_state.db
        if mat_to_edit_norm not in db["material"].values:
            st.error("Material not found in database.")
        else:
            # Current symbol (x)
            old_x = db.loc[db["material"] == mat_to_edit_norm, "symbol"].iloc[0]

            st.markdown("### Current Logic:")

            # Decode x into components
            logic = st.session_state.logic
            def reverse_lookup(code, table):
                df = logic[table]
                match = df[df["code"] == code]
                return match["name"].iloc[0] if not match.empty else None

            # Attempt to split old_x based on code lengths in Logic.xlsx
            cap_len = logic["capacity"]["code"].str.len().mode()[0]
            fam_len = logic["family"]["code"].str.len().mode()[0]
            opt_len = logic["option"]["code"].str.len().mode()[0]
            col_len = logic["color"]["code"].str.len().mode()[0]

            c_code = old_x[:cap_len]
            f_code = old_x[cap_len:cap_len+fam_len]
            o_code = old_x[cap_len+fam_len:cap_len+fam_len+opt_len]
            col_code = old_x[cap_len+fam_len+opt_len:cap_len+fam_len+opt_len+col_len]

            picked = {
                "capacity": reverse_lookup(c_code, "capacity"),
                "family": reverse_lookup(f_code, "family"),
                "option": reverse_lookup(o_code, "option"),
                "color": reverse_lookup(col_code, "color")
            }

            # UI for editing logic
            st.markdown("### New Logic Selection")

            sel_capacity = st.selectbox(
                "Capacity", logic["capacity"]["name"], index=
                logic["capacity"].index[logic["capacity"]["name"] == picked["capacity"]].tolist()[0]
                if picked["capacity"] in logic["capacity"]["name"].values else 0
            )
            sel_family = st.selectbox(
                "Family", logic["family"]["name"], index=
                logic["family"].index[logic["family"]["name"] == picked["family"]].tolist()[0]
                if picked["family"] in logic["family"]["name"].values else 0
            )
            sel_option = st.selectbox(
                "Option", logic["option"]["name"], index=
                logic["option"].index[logic["option"]["name"] == picked["option"]].tolist()[0]
                if picked["option"] in logic["option"]["name"].values else 0
            )
            sel_color = st.selectbox(
                "Color", logic["color"]["name"], index=
                logic["color"].index[logic["color"]["name"] == picked["color"]].tolist()[0]
                if picked["color"] in logic["color"]["name"].values else 0
            )

            # Build new x
            maps = {
                cat: dict(zip(logic[cat]["name"], logic[cat]["code"]))
                for cat in ["capacity","family","option","color"]
            }

            new_x = (
                maps["capacity"][sel_capacity] +
                maps["family"][sel_family] +
                maps["option"][sel_option] +
                maps["color"][sel_color]
            )

            st.info(f"New generated x code: **{new_x}**")

            if st.button("Save Updated Logic"):
                # Update DB
                db.loc[db["material"] == mat_to_edit_norm, "symbol"] = new_x
                st.session_state.db = db
                save_database(db)

                # Update history rows
                hist = st.session_state.history.copy()

                # Update x field
                hist.loc[hist["MaterialCode"] == mat_to_edit_norm, "x"] = new_x

                # Rebuild symbol
                hist["symbol"] = hist.apply(
                    lambda r: f"{r['Month']:02d}{r['Year'] % 100:02d}{r['Y']}",
                    axis=1
                )

                # Recompute code34
                hist["code34"] = hist.apply(
                    lambda r: f"{r['x']}&PO={r['OrderNumber']}&PC={r['MaterialCode']}",
                    axis=1
                )
                hist["code_length"] = hist["code34"].str.len()

                # Update barcode
                bmap = dict(zip(st.session_state.barcode["material"],
                                st.session_state.barcode["barcode"]))
                hist["الباركود الدولي"] = hist["MaterialCode"].map(bmap).fillna("")

                # Recompute LastSerial
                hist = recompute_last_serial(hist)

                st.session_state.history = hist
                save_df_to_excel(hist, APP_HISTORY_FILE, sheet="HISTORY")

                st.success("Logic updated in Database and History. Recalculation completed.")
                _safe_rerun()

# ==========================================================
# COLLAPSIBLE HISTORY MAINTENANCE SECTION
# ==========================================================
with st.expander("History Maintenance (Search + Filters + Delete Rows)", expanded=False):

    st.markdown("### Search & Filters")

    hist = st.session_state.history.copy()
    hist = hist.reset_index().rename(columns={"index": "RowID"})

    # Global search
    global_search = st.text_input(
        "Global search (search all columns):",
        ""
    ).strip().lower()

    # Column filters
    with st.expander("Column Filters"):
        for col in hist.columns:
            if col in ["Select"]:
                continue
            unique_vals = hist[col].dropna().unique()
            if len(unique_vals) <= 30:
                selected = st.multiselect(f"Filter '{col}'", unique_vals.tolist(), [])
                if selected:
                    hist = hist[hist[col].isin(selected)]
            else:
                txt = st.text_input(f"Filter '{col}' contains:", "")
                if txt.strip():
                    hist = hist[hist[col].astype(str).str.contains(txt.strip(), case=False, na=False)]

    # Apply global search
    if global_search:
        mask = pd.Series(False, index=hist.index)
        for col in hist.columns:
            mask = mask | hist[col].astype(str).str.lower().str.contains(global_search, na=False)
        hist = hist[mask]

    # Add selection checkbox
    if "Select" not in hist.columns:
        hist["Select"] = False

    # Display interactive table
    edited_hist = st.data_editor(
        hist,
        use_container_width=True,
        hide_index=True,
        height=400,
        column_config={
            "Select": st.column_config.CheckboxColumn("Select")
        }
    )

    # Determine selected rows
    to_delete = edited_hist.loc[edited_hist["Select"] == True, "RowID"].tolist()

    colA, colB = st.columns(2)
    with colA:
        if st.button("Delete Selected Rows"):
            if not to_delete:
                st.warning("No rows selected.")
            else:
                hist2 = st.session_state.history.copy()
                safe_ids = [rid for rid in to_delete if rid in hist2.index]
                hist2 = hist2.drop(index=safe_ids)
                hist2 = recompute_last_serial(hist2)
                st.session_state.history = hist2
                save_df_to_excel(hist2, APP_HISTORY_FILE, sheet="HISTORY")
                st.success(f"Deleted {len(safe_ids)} row(s).")
                _safe_rerun()

    with colB:
        if st.button("Download History.xlsx"):
            st.download_button(
                "Download",
                data=APP_HISTORY_FILE.read_bytes(),
                file_name="History.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

st.markdown("---")
st.caption("All updates saved automatically. Database and History stay in sync.")
# ==== Centered footer ====
footer_css = """
<style>
.app-footer {
  position: fixed;
  left: 50%;
  bottom: 12px;
  transform: translateX(-50%);
  z-index: 9999;
  background: rgba(255,255,255,0.85);
  border: 1px solid #e6e6e6;
  border-radius: 14px;
  padding: 8px 14px;
  font-weight: 600;
  font-size: 14px;
}
</style>
"""
footer_html = """
<div class="app-footer">✨ تم التنفيذ بواسطة م / بيتر عمانوئيل – جميع الحقوق محفوظة © 2025 ✨</div>
"""
st.markdown(footer_css, unsafe_allow_html=True)
st.markdown(footer_html, unsafe_allow_html=True)


