import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import io
import matplotlib.pyplot as plt

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch

st.set_page_config(page_title="Sterilizer Staggering", layout="wide")

st.title("📊 Sterilizer Staggering Report Generator")

# Helper Functions
@st.cache_data
def process_data(df):
    df = df.copy()

    df['Date'] = pd.to_datetime(df['Date'])
    df['process_time'] = pd.to_datetime(df['process_time'], format="%H:%M:%S")

    df['datetime'] = pd.to_datetime(
        df['Date'].astype(str) + ' ' + df['process_time'].dt.strftime('%H:%M:%S')
    )

    df.loc[(df['Shift'] == 'Night') & (df['datetime'].dt.hour < 12), 'datetime'] += pd.Timedelta(days=1)

    df['Staggering'] = np.where(
        df['Sterilizer'].isin(['A','B','C','D']),
        'Bariquands',
        'Retorts'
    )

    df['label'] = df['Sterilizer'] + "," + df['process_time'].dt.strftime("%H:%M")

    df = df.drop_duplicates(subset='label', keep='first')
    df = df.sort_values('datetime')
    df['sequence'] = range(1, len(df) + 1)

    return df.reset_index(drop=True)

def create_chart(df):
    plot_date = pd.to_datetime(df['Date'].iloc[0]).strftime("%d/%m/%Y")

    fig = px.line(
        df,
        x='datetime',
        y='sequence',
        title=f"Sterilizers Staggering - {plot_date} - {df['Shift'].iloc[0]}",
        markers=True,
        text='label'
    )

    fig.update_traces(
        marker=dict(size=8, symbol='diamond'),
        line=dict(width=2),
        textposition='top left'
    )

    fig.update_xaxes(
        tickformat="%H:%M",
        dtick=3600000,
        title="Process Time"
    )

    fig.update_yaxes(title="Process Order")

    fig.update_layout(height=700, width=1700)

    return fig

# Matplotlib Chart for PDF

def create_matplotlib_chart(df):
    fig, ax = plt.subplots(figsize=(14, =6))

    ax.plot(df['datetime'], df['sequence'], marker='D')

    for _, row in df.iterrows():
        ax.text(row['datetime'], row['sequence'], row['label'], fontsize=7)

    ax.set_title(f"Sterilizer Staggering - {df['Shift'].iloc[0]}")
    ax.set_xlabel("Time")
    ax.set_ylabel("Process Order")

    fig.autofmt_xdate()

    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=300, bbox_inches='tight')
    plt.close(fig)

    buf.seek(0)
    return buf

# PDF Generator (A4)
def generate_pdf(df):
    buffer = io.BytesIO()

    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4))
    styles = getSampleStyleSheet()

    elements = []

    # --- Header / Branding ---
    title = Paragraph("<b>Sterilizer Staggering Report</b>", styles['Title'])
    subtitle = Paragraph(
        f"Date: {df['Date'].iloc[0].strftime('%d %b %Y')} | Shift: {df['Shift'].iloc[0]}",
        styles['Normal']
    )

    elements.append(title)
    elements.append(Spacer(1, 10))
    elements.append(subtitle)
    elements.append(Spacer(1, 20))

    # --- Chart ---
    chart_img = create_matplotlib_chart(df)
    page_width, page_height = landscape(A4)

    # margins (default ~72 each side → subtract ~144 total)
    usable_width = page_width - 100  

    img = Image(chart_img, width=usable_width, height=usable_width * 0.55)

    elements.append(img)
    elements.append(Spacer(1, 20))

    # --- Footer ---
    footer = Paragraph("Generated via Streamlit | Josphat Njau", styles['Normal'])
    elements.append(footer)

    doc.build(elements)

    buffer.seek(0)
    return buffer

# UI
uploaded_file = st.file_uploader("📂 Upload Excel File", type=["xlsx"])

if uploaded_file:
    excel_file = pd.ExcelFile(uploaded_file)
    sheet = st.selectbox("📄 Select Sheet", excel_file.sheet_names)

    df = pd.read_excel(uploaded_file, sheet_name=sheet)

    required_cols = ['Date', 'process_time', 'Shift', 'Sterilizer']

    if not all(col in df.columns for col in required_cols):
        st.error(f"Missing required columns: {required_cols}")
    else:
        df = process_data(df)

        st.success("✅ Data processed successfully")

        with st.expander("🔍 Preview Data"):
            st.dataframe(df.head())

        fig = create_chart(df)
        st.plotly_chart(fig, width='stretch')

        if st.button("📄 Generate PDF Report"):
            pdf_buffer = generate_pdf(df)

            st.success("PDF report generated")

            st.download_button(
                "⬇️ Download PDF",
                data=pdf_buffer,
                file_name="staggering_report.pdf",
                mime="application/pdf"
            )

        st.markdown("---")
        st.caption("Built with Streamlit | Josphat Njau")
