import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import io

st.set_page_config(page_title="Sterilizer Staggering", layout="wide")

st.title("📊 Sterilizer Staggering Report Generator")

# ----------------------------
# DATA PROCESSING
# ----------------------------
@st.cache_data
def process_data(df):
    df = df.copy()

    df['Date'] = pd.to_datetime(df['Date'])
    df['process_time'] = pd.to_datetime(df['process_time'], format="%H:%M:%S")

    df['datetime'] = pd.to_datetime(
        df['Date'].astype(str) + ' ' + df['process_time'].dt.strftime('%H:%M:%S')
    )

    # Night shift correction
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


# ----------------------------
# CHART (UI PREVIEW - Plotly)
# ----------------------------
def create_chart(df):
    date_str = pd.to_datetime(df['Date'].iloc[0]).strftime("%d/%m/%Y")
    shift = df['Shift'].iloc[0]

    fig = px.line(
        df,
        x='datetime',
        y='sequence',
        markers=True,
        text='label'
    )

    fig.update_traces(
        marker=dict(size=8, symbol='diamond'),
        line=dict(width=2),
        textposition='top center'
    )

    fig.update_layout(
        title=f"Staggering - {shift} - {date_str}",
        height=700,
        width=1700,
        title_font=dict(size=20),
        xaxis_title="Process Time",
        yaxis_title="Process Order"
    )

    return fig


# ----------------------------
# PDF EXPORT (MATPLOTLIB ONLY)
# ----------------------------
def export_pdf(df):
    date_str = pd.to_datetime(df['Date'].iloc[0]).strftime("%d/%m/%y")
    shift = df['Shift'].iloc[0]

    title = f"Staggering - {shift} - {date_str}"

    buffer = io.BytesIO()

    with PdfPages(buffer) as pdf:

        fig, ax = plt.subplots(figsize=(17, 10))  # landscape A4 ratio

        ax.plot(df['datetime'], df['sequence'], marker='D', linewidth=2)

        # Labels above points
        for i, row in df.iterrows():
            ax.text(
                row['datetime'],
                row['sequence'] + 0.2,
                row['label'],
                ha='center',
                fontsize=8
            )

        ax.set_title(title, fontsize=18, fontweight='bold')
        ax.set_xlabel("Process Time")
        ax.set_ylabel("Process Order")

        ax.grid(True, linestyle="--", alpha=0.5)

        plt.xticks(rotation=45)

        # FORCE SINGLE PAGE
        pdf.savefig(fig, bbox_inches='tight')
        plt.close()

    buffer.seek(0)
    return buffer, f"{shift}_staggering_{date_str}.pdf"


# ----------------------------
# UI
# ----------------------------
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

        st.success("Data processed successfully")

        with st.expander("Preview"):
            st.dataframe(df.head())

        fig = create_chart(df)
        st.plotly_chart(fig, use_container_width=True)

        if st.button("Generate PDF Report"):

            pdf_buffer, filename = export_pdf(df)

            st.success("PDF generated successfully")

            st.download_button(
                "Download PDF",
                data=pdf_buffer,
                file_name=filename,
                mime="application/pdf"
            )

st.caption("Built with Streamlit")
