
import streamlit as st
import pandas as pd
import io
import plotly.express as px
import os
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

st.set_page_config(page_title="JHA Interactive v2", layout="wide")

DEFAULT_FILE = "JHA by Division.xlsx"

def find_excel_file():
    # prefer exact default filename in same folder, otherwise look for any .xlsx in cwd
    if os.path.exists(DEFAULT_FILE):
        return DEFAULT_FILE
    for f in os.listdir("."):
        if f.lower().endswith(".xlsx") or f.lower().endswith(".xls"):
            return f
    return None

@st.cache_data
def load_workbook(path):
    xls = pd.ExcelFile(path)
    sheets = xls.sheet_names
    data = {s: pd.read_excel(xls, sheet_name=s, dtype=object) for s in sheets}
    # normalize columns and forward-fill division-like columns
    for name, df in data.items():
        df.columns = [str(c).strip() for c in df.columns]
        div_col = next((c for c in df.columns if "division" in c.lower()), None)
        if div_col:
            df[div_col] = df[div_col].ffill()
        data[name] = df
    return data, sheets

def to_excel_bytes(df):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Filtered")
    out.seek(0)
    return out.getvalue()

def make_pdf_report(title, chart_fig, df_rows):
    # chart_fig is a plotly figure; we'll render it to PNG via kaleido (fig.to_image)
    img_bytes = None
    try:
        img_bytes = chart_fig.to_image(format="png", width=800, height=400, scale=2)
    except Exception as e:
        img_bytes = None
    out = io.BytesIO()
    c = canvas.Canvas(out, pagesize=letter)
    width, height = letter
    c.setFont("Helvetica-Bold", 16)
    c.drawString(40, height - 40, title)
    y = height - 70
    if img_bytes:
        img = ImageReader(io.BytesIO(img_bytes))
        iw, ih = img.getSize()
        # scale to page width margins
        max_w = width - 80
        scale = min(1.0, max_w / iw)
        c.drawImage(img, 40, y - int(ih*scale), width=iw*scale, height=ih*scale)
        y = y - int(ih*scale) - 20
    # write table rows (limit to first 50 rows to keep PDF reasonable)
    c.setFont("Helvetica", 10)
    rows = df_rows if len(df_rows) <= 50 else df_rows[:50]
    for i, r in enumerate(rows):
        text = " | ".join([f"{k}: {str(v)[:80]}" for k,v in r.items()])
        text_lines = [text[j:j+200] for j in range(0, len(text), 200)]
        for tl in text_lines:
            if y < 60:
                c.showPage()
                y = height - 40
                c.setFont("Helvetica", 10)
            c.drawString(40, y, tl)
            y -= 12
    c.save()
    out.seek(0)
    return out.getvalue()

st.title("JHA Interactive — v2 (Multi-sheet + Charts + Exports)")

excel_file = find_excel_file()
if not excel_file:
    st.error("No Excel file found in this folder. Please place your Excel file (e.g., 'JHA by Division.xlsx') in the same folder as this app.")
    st.stop()

data_dict, sheets = load_workbook(excel_file)

# Sidebar: global controls
st.sidebar.header("Global Controls")
sheet_choice = st.sidebar.selectbox("Select sheet", sheets)
search_q = st.sidebar.text_input("Search (any column)")
dark_mode = st.sidebar.checkbox("Dark mode", value=False)

# Per-sheet data
df = data_dict[sheet_choice].copy()

# detect useful columns
division_col = next((c for c in df.columns if "division" in c.lower()), None)
risk_col = next((c for c in df.columns if "risk" in c.lower()), None)
hazard_col = next((c for c in df.columns if "hazard" in c.lower()), None)
control_col = next((c for c in df.columns if "control" in c.lower()), None)

# Filters UI
st.sidebar.subheader("Filters for sheet: " + sheet_choice)
sel_div = "All"
if division_col:
    divisions = ["All"] + sorted(df[division_col].dropna().astype(str).unique().tolist())
    sel_div = st.sidebar.selectbox("Division", divisions)
if risk_col:
    risks = ["All"] + sorted(df[risk_col].dropna().astype(str).unique().tolist())
    sel_risk = st.sidebar.selectbox("Risk Level", risks)
else:
    sel_risk = "All"

# Apply filters
filtered = df.copy()
if division_col and sel_div != "All":
    filtered = filtered[filtered[division_col].astype(str) == str(sel_div)]
if risk_col and sel_risk != "All":
    filtered = filtered[filtered[risk_col].astype(str) == str(sel_risk)]
if search_q:
    q = search_q.lower()
    mask = filtered.apply(lambda row: row.astype(str).str.lower().str.contains(q, na=False).any(), axis=1)
    filtered = filtered[mask]

st.sidebar.markdown(f"Filtered rows: **{len(filtered)}**")

# Main layout: two columns
col1, col2 = st.columns((2,1))

with col1:
    st.subheader(f"Data — {sheet_choice}")
    st.dataframe(filtered, use_container_width=True)

    # Row details: select one row index to expand
    st.subheader("Row details")
    if len(filtered) == 0:
        st.info("No rows to show details for.")
    else:
        idx = st.number_input("Select row number (0-based index)", min_value=0, max_value=max(0, len(filtered)-1), value=0, step=1)
        row = filtered.iloc[int(idx)].to_dict()
        st.json(row)

    # Exports
    st.download_button("Download filtered CSV", data=filtered.to_csv(index=False).encode('utf-8'),
                       file_name="jha_filtered.csv", mime="text/csv")
    st.download_button("Download filtered Excel (.xlsx)", data=to_excel_bytes(filtered),
                       file_name="jha_filtered.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # PDF export
    if st.button("Generate PDF report (first 50 rows + chart)"):
        # build a chart for the PDF: prefer Division counts
        chart_fig = None
        if division_col:
            counts = df[division_col].fillna('Unknown').value_counts().reset_index()
            counts.columns = ["division", "count"]
            chart_fig = px.bar(counts, x='count', y='division', orientation='h', title='JHAs by Division')
        elif risk_col:
            counts = df[risk_col].fillna('Unknown').value_counts().reset_index()
            counts.columns = ["risk", "count"]
            chart_fig = px.pie(counts, values='count', names='risk', title='JHAs by Risk Level')
        else:
            chart_fig = px.scatter(x=[0], y=[0], title='No chart available')

        pdf_bytes = make_pdf_report(f"JHA Report — {sheet_choice}", chart_fig, filtered.to_dict(orient='records'))
        st.download_button("Download PDF report", data=pdf_bytes, file_name="jha_report.pdf", mime="application/pdf")

with col2:
    st.subheader("Charts & Dashboard")
    # JHAs by Division
    if division_col:
        counts = df[division_col].fillna('Unknown').value_counts().reset_index()
        counts.columns = ["division", "count"]
        fig_div = px.bar(counts, x='count', y='division', orientation='h', title='JHAs by Division', height=400)
        st.plotly_chart(fig_div, use_container_width=True)
    # JHAs by Risk
    if risk_col:
        rc = df[risk_col].fillna('Unknown').value_counts().reset_index()
        rc.columns = ["risk", "count"]
        fig_risk = px.pie(rc, values='count', names='risk', title='JHAs by Risk Level', height=300)
        st.plotly_chart(fig_risk, use_container_width=True)

    # Hazards vs Controls cross-tab (if available)
    if hazard_col and control_col:
        ct = pd.crosstab(df[hazard_col].fillna('Unknown'), df[control_col].fillna('Unknown'))
        st.subheader('Hazards vs Controls (sample)')
        st.dataframe(ct.reset_index().head(200))

st.markdown("---")
st.caption("Place your Excel file in the same folder as this script. This prototype supports multi-sheet browsing, exports, and PDF generation.")
