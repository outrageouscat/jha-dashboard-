
import streamlit as st
import pandas as pd
import io
import os
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

st.set_page_config(page_title="JHA Interactive v3", layout="wide")

DEFAULT_XLSX = "JHA by Division.xlsx"

def find_file():
    if os.path.exists(DEFAULT_XLSX):
        return DEFAULT_XLSX
    for f in os.listdir("."):
        if f.lower().endswith(".xlsx") or f.lower().endswith(".xls"):
            return f
    return None

@st.cache_data
def load_sheets(path):
    xls = pd.ExcelFile(path)
    sheets = xls.sheet_names
    data = {}
    raw = {}
    for i, s in enumerate(sheets):
        df_raw = pd.read_excel(xls, sheet_name=s, header=None, dtype=object)
        raw[s] = df_raw
        if i >= 1 and len(df_raw) >= 2:
            hdr1 = df_raw.iloc[0].fillna('').astype(str)
            hdr2 = df_raw.iloc[1].fillna('').astype(str)
            cols = []
            for a, b in zip(hdr1, hdr2):
                a = a.strip()
                b = b.strip()
                if a and b:
                    cols.append(f"{a} — {b}")
                elif a:
                    cols.append(a)
                elif b:
                    cols.append(b)
                else:
                    cols.append("Unnamed")
            df = df_raw.iloc[2:].copy().reset_index(drop=True)
            df.columns = cols
        else:
            df = df_raw.copy()
        data[s] = df
    return data, sheets, raw

def to_excel_bytes(df):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Filtered")
    out.seek(0)
    return out.getvalue()

def make_pdf_text(title, rows):
    out = io.BytesIO()
    c = canvas.Canvas(out, pagesize=letter)
    width, height = letter
    c.setFont("Helvetica-Bold", 16)
    c.drawString(40, height-40, title)
    y = height - 70
    c.setFont("Helvetica", 10)
    for r in rows:
        line = str(r)
        for i in range(0, len(line), 200):
            part = line[i:i+200]
            if y < 60:
                c.showPage()
                y = height - 40
                c.setFont("Helvetica", 10)
            c.drawString(40, y, part)
            y -= 12
    c.save()
    out.seek(0)
    return out.getvalue()

file = find_file()
if not file:
    st.error("No Excel file found in this folder. Please place 'JHA by Division.xlsx' with the app.")
    st.stop()

data_dict, sheets, raw_sheets = load_sheets(file)

st.sidebar.title("Navigation")
page = st.sidebar.radio("Go to", ["Home (Overview)", "Search / Edit", "Analytics", "Download"])

# Map sheets by index
# 0: landing; 1: Key JHAs by Division; 2: Critical JHAs by Division; 3: Critical JHAs Summary; 4: Primary Hazards; 5: Primary Controls
if page == "Home (Overview)":
    st.title("Overview")
    landing = data_dict[sheets[0]]
    if landing.shape[1] == 1:
        text = "\\n\\n".join(str(x[0]) for x in landing.values.tolist() if pd.notna(x[0]))
        st.markdown(text)
    else:
        for i in range(min(len(landing), 40)):
            row = landing.iloc[i].dropna().astype(str).tolist()
            if row:
                st.markdown(" ".join(row))

elif page == "Search / Edit":
    st.title("Search Hazards & Controls")
    if len(sheets) <= 1:
        st.error("Workbook missing required sheets. Expected at least 6 sheets as described.")
        st.stop()
    key_sheet = data_dict[sheets[1]]
    hazards_sheet = data_dict[sheets[4]] if len(sheets) > 4 else pd.DataFrame()
    controls_sheet = data_dict[sheets[5]] if len(sheets) > 5 else pd.DataFrame()

    def find_col(df, patterns):
        for p in patterns:
            for c in df.columns:
                if p.lower() in str(c).lower():
                    return c
        return None

    division_col = find_col(key_sheet, ["division"])
    task_col = find_col(key_sheet, ["task", "sequence", "job step", "what am i doing"])
    hazard_col_haz = find_col(hazards_sheet, ["hazard", "primary hazard"])
    control_col_ctrl = find_col(controls_sheet, ["control", "primary control"])

    divisions = []
    if division_col:
        divisions = sorted(key_sheet[division_col].dropna().astype(str).unique().tolist())
    else:
        if hazard_col_haz:
            divisions = sorted(hazards_sheet[hazard_col_haz].dropna().astype(str).unique().tolist())
    divisions = ["-- Select Division --"] + divisions

    sel_div = st.selectbox("Select Division", divisions)
    if sel_div == "-- Select Division --":
        st.info("Pick a division to view related tasks, hazards, and controls.")
    else:
        filtered_key = key_sheet[key_sheet[division_col].astype(str) == str(sel_div)] if division_col is not None else key_sheet
        st.subheader("What am I doing? (Tasks)")
        tasks = []
        if task_col and task_col in filtered_key.columns:
            tasks = filtered_key[task_col].fillna("").astype(str).tolist()
        else:
            other_cols = [c for c in filtered_key.columns if c != division_col]
            if other_cols:
                tasks = filtered_key[other_cols[0]].fillna("").astype(str).tolist()
        default_task = "\\n\\n".join(tasks) if tasks else ""
        if "edited_task" not in st.session_state:
            st.session_state.edited_task = default_task
        st.text_area("Edit task (temporary)", key="edited_task", height=200)

        st.subheader("How can I hurt myself? (Primary Hazards)")
        hazard_text = ""
        if hazard_col_haz and division_col and hazard_col_haz in hazards_sheet.columns:
            hazard_div_col = find_col(hazards_sheet, ["division"])
            if hazard_div_col:
                matched = hazards_sheet[hazard_div_col].astype(str) == str(sel_div)
                if matched.any():
                    hazard_text = "\\n\\n".join(hazards_sheet.loc[matched, hazard_col_haz].fillna("").astype(str).tolist())
        if not hazard_text and hazard_col_haz and hazard_col_haz in hazards_sheet.columns:
            hazard_text = "\\n\\n".join(hazards_sheet[hazard_col_haz].dropna().astype(str).unique().tolist())
        if "edited_hazard" not in st.session_state:
            st.session_state.edited_hazard = hazard_text
        st.text_area("Edit primary hazards (temporary)", key="edited_hazard", height=200)

        st.subheader("How do I protect myself? (Primary Controls)")
        control_text = ""
        if control_col_ctrl and division_col and control_col_ctrl in controls_sheet.columns:
            control_div_col = find_col(controls_sheet, ["division"])
            if control_div_col:
                matched = controls_sheet[control_div_col].astype(str) == str(sel_div)
                if matched.any():
                    control_text = "\\n\\n".join(controls_sheet.loc[matched, control_col_ctrl].fillna("").astype(str).tolist())
        if not control_text and control_col_ctrl and control_col_ctrl in controls_sheet.columns:
            control_text = "\\n\\n".join(controls_sheet[control_col_ctrl].dropna().astype(str).unique().tolist())
        if "edited_control" not in st.session_state:
            st.session_state.edited_control = control_text
        st.text_area("Edit primary controls (temporary)", key="edited_control", height=200)

        combined_df = pd.DataFrame([{{"Division": sel_div,
                                     "Task": st.session_state.edited_task,
                                     "Primary Hazards": st.session_state.edited_hazard,
                                     "Primary Controls": st.session_state.edited_control}}])
        csv_bytes = combined_df.to_csv(index=False).encode('utf-8')
        st.download_button("Download combined as CSV", data=csv_bytes, file_name="jha_combined.csv", mime="text/csv")
        st.download_button("Download combined as Excel (.xlsx)", data=to_excel_bytes(combined_df), file_name="jha_combined.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        if st.button("Download combined as PDF"):
            rows = [f"Division: {sel_div}", "", "Task:", st.session_state.edited_task, "", "Primary Hazards:", st.session_state.edited_hazard, "", "Primary Controls:", st.session_state.edited_control]
            pdf = make_pdf_text(f"JHA — {sel_div}", rows)
            st.download_button("Click to save PDF", data=pdf, file_name=f"jha_{sel_div}.pdf", mime="application/pdf")

elif page == "Analytics":
    st.title("Analytics")
    key_sheet = data_dict[sheets[1]] if len(sheets) > 1 else pd.DataFrame()
    def find_col_simple(df, pat):
        for c in df.columns:
            if pat.lower() in str(c).lower():
                return c
        return None
    divc = find_col_simple(key_sheet, "division") if not key_sheet.empty else None
    if divc and not key_sheet.empty:
        counts = key_sheet[divc].fillna("Unknown").astype(str).value_counts()
        st.bar_chart(counts)
    else:
        st.info("No Division column detected in Key JHAs sheet for analytics.")

else:
    st.title("Download Center")
    for s in sheets:
        st.write(f"Sheet: {s} — rows: {{len(data_dict[s])}} columns: {{len(data_dict[s].columns)}}")
        csv = data_dict[s].to_csv(index=False).encode('utf-8')
        st.download_button(f"Download {s} as CSV", data=csv, file_name=f"{s}.csv", mime="text/csv")
    if st.button("Download entire workbook as single Excel"):
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as w:
            for s in sheets:
                data_dict[s].to_excel(w, sheet_name=s[:31], index=False)
        out.seek(0)
        st.download_button("Click to save workbook", data=out.getvalue(), file_name="jha_workbook_export.xlsx", mime="application/vnd.openxmlformats-officedocument-spreadsheetml.sheet")

st.sidebar.markdown("---")
st.sidebar.caption("Editable fields are temporary for this session only.")
