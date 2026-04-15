import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import time
from pypdf import PdfReader, PdfWriter
import os

st.set_page_config(page_title="Sterilizer Staggering", layout="wide")

st.title("📊 Sterilizer Staggering Report Generator")

# -------------------------------
# Helper Functions
# -------------------------------

@st.cache_data
def process_data(df):
    df = df.copy()

    df['Date'] = pd.to_datetime(df['Date'])
    df['process_time'] = pd.to_datetime(df['process_time'], format="%H:%M:%S")

    df['datetime'] = pd.to_datetime(
        df['Date'].astype(str) + ' ' + df['process_time'].dt.strftime('%H:%M:%S')
    )

    # Fix night shift rollover
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

    fig.update_layout(
        height=700,
        width=1200,
        title_font=dict(family="Times New Roman", weight="bold", size=20)
    )

    return fig


def export_pdf(fig, df):
    f_date = pd.to_datetime(df['Date'].iloc[0]).strftime("%d-%m-%Y")
    filename = f"{df['Shift'].iloc[0]}_staggering_{f_date}.pdf"

    fig.write_image(
        filename,
        format="pdf",
        engine="kaleido",
        width=1700,
        height=1000,
        scale=2
    )

    return filename


def update_master_pdf(new_pdf, master_pdf="Master_Staggering_Report.pdf"):
    writer = PdfWriter()

    # Add new report first
    reader_new = PdfReader(new_pdf)
    for page in reader_new.pages:
        writer.add_page(page)

    # Append existing master
    if os.path.exists(master_pdf):
        reader_old = PdfReader(master_pdf)
        for page in reader_old.pages:
            writer.add_page(page)

    temp_output = "temp_master.pdf"

    with open(temp_output, "wb") as f:
        writer.write(f)

    os.replace(temp_output, master_pdf)

    return master_pdf


# -------------------------------
# UI
# -------------------------------

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

        # Show preview
        with st.expander("🔍 Preview Data"):
            st.dataframe(df.head())

        fig = create_chart(df)

        st.plotly_chart(fig, width='stretch')

        col1, col2 = st.columns(2)

        with col1:
            if st.button("📄 Generate PDF"):
                filename = export_pdf(fig, df)
                st.success(f"PDF generated: {filename}")

                with open(filename, "rb") as f:
                    st.download_button(
                        "⬇️ Download PDF",
                        f,
                        file_name=filename
                    )

        with col2:
            if st.button("📚 Update Master Report"):
                filename = export_pdf(fig, df)
                master = update_master_pdf(filename)

                with open(master, "rb") as f:
                    st.download_button(
                        "⬇️ Download Master Report",
                        f,
                        file_name=master
                    )

        st.markdown("---")
        st.caption("Built with Streamlit | Josphat Njau")
