import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import io
import matplotlib.pyplot as plt

from reportlab.platypus import SimpleDocTemplate, Image
from reportlab.lib.pagesizes import A4, landscape

st.set_page_config(page_title="Sterilizer Staggering", layout="wide")

st.title("📊 Sterilizer Staggering Report Generator")

# DATA PROCESSING
@st.cache_data
def process_data(df):
    df = df.copy()

    df['Date'] = pd.to_datetime(df['Date'])
    df['process_time'] = pd.to_datetime(df['process_time'], format="%H:%M:%S")

    df['datetime'] = pd.to_datetime(
        df['Date'].astype(str) + ' ' + df['process_time'].dt.strftime('%H:%M:%S')
    )

    # Night shift correction
    df.loc[
        (df['Shift'] == 'Night') & (df['datetime'].dt.hour < 12),
        'datetime'
    ] += pd.Timedelta(days=1)

    df['Staggering'] = np.where(
        df['Sterilizer'].isin(['A', 'B', 'C', 'D']),
        'Bariquands',
        'Retorts'
    )

    df['label'] = df['Sterilizer'] + "," + df['process_time'].dt.strftime("%H:%M")

    df = df.drop_duplicates(subset='label', keep='first')
    df = df.sort_values('datetime')
    df['sequence'] = range(1, len(df) + 1)

    return df.reset_index(drop=True)

# STREAMLIT CHART (DISPLAY ONLY)
def create_chart(df):
    plot_date = pd.to_datetime(df['Date'].iloc[0]).strftime("%d/%m/%y")

    fig = px.line(
        df,
        x='datetime',
        y='sequence',
        markers=True,
        text='label'
    )

    fig.update_traces(
        marker=dict(size=9, symbol='diamond'),
        line=dict(width=2),
        textposition='top center'
    )

    fig.update_layout(
        title=f"Sterilizer Staggering - Day - {plot_date}",
        height=700,
        width=1700,
        showlegend=False,
        margin=dict(l=40, r=40, t=80, b=40)
    )

    return fig

# PDF EXPORT (MATPLOTLIB ONLY)
def export_pdf(df):
    buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=0,
        rightMargin=0,
        topMargin=0,
        bottomMargin=0
    )

    # Matplotlib chart (print)
    fig, ax = plt.subplots(figsize=(17, 10))

    ax.plot(df['datetime'], df['sequence'], marker='D', linewidth=2)

    # Labels above points
    for _, row in df.iterrows():
        ax.text(
            row['datetime'],
            row['sequence'] + 0.15,
            row['label'],
            fontsize=9,
            ha='center'
        )

    # Title format required
    plot_date = pd.to_datetime(df['Date'].iloc[0]).strftime("%d/%m/%y")
    ax.set_title(f"Staggering - Day - {plot_date}", fontsize=20, pad=20)

    ax.set_xlabel("")
    ax.set_ylabel("")
    ax.grid(True, linestyle="--", alpha=0.4)

    fig.autofmt_xdate()

    # Save high-res image
    img_buffer = io.BytesIO()
    plt.savefig(img_buffer, format="png", dpi=200, bbox_inches="tight")
    plt.close(fig)

    img_buffer.seek(0)

    # PDF (full page image)
    page_width, page_height = landscape(A4)

    img = Image(img_buffer)
    img.drawWidth = page_width
    img.drawHeight = page_height

    doc.build([img])

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

        # Display chart
        fig = create_chart(df)
        st.plotly_chart(fig, use_container_width=True)

        # PDF button
        if st.button("📄 Generate PDF Report"):

            pdf_buffer = export_pdf(df)

            st.success("PDF generated successfully")

            st.download_button(
                "⬇️ Download PDF",
                data=pdf_buffer,
                file_name=f"staggering_{df['Shift'].iloc[0]}.pdf",
                mime="application/pdf"
            )

        st.markdown("---")
        st.caption("Built with Streamlit | Josphat Njau")
