import streamlit as st
import pandas as pd
from io import BytesIO

# ============================
#  CONFIG TEMA & LAYOUT
# ============================
st.set_page_config(
    page_title="Excel Merge",
    layout="centered"
)

# ============================
# SEMBUNYIKAN PREVIEW FILE
# ============================
st.markdown("""
<style>
.uploadedFile {display: none !important;}
.stUploadedFile {display: none !important;}
</style>
""", unsafe_allow_html=True)

# ============================
# BALIK URUTAN TAMPILAN FILE
# ============================
st.markdown("""
<style>
.stFileUploader > div:nth-child(1) > div {
    display: flex;
    flex-direction: column-reverse;
}
</style>
""", unsafe_allow_html=True)

# ============================
# PERBESAR LIST FILE
# ============================
st.markdown("""
<style>
.stFileUploader > div > div {
    max-height: 450px !important;
    overflow-y: auto !important;
}
</style>
""", unsafe_allow_html=True)

# CSS agar width max 750px
st.markdown("""
<style>
.main {
    max-width: 750px;
    margin: auto;
    padding-top: 20px;
}
</style>
""", unsafe_allow_html=True)

# ============================
#  HEADER
# ============================
st.markdown("""
<div style='text-align:left; padding: 10px 0;'>
    <h1 style='margin-bottom: 0; font-size: 40px; color:#150F3D;'>
        Excel Merge
    </h1>
    <p style='font-size:16px; color:#555;'>
        This tool automatically merges multiple Excel files into a single dataset with unified formatting
    </p>
    <hr style='margin-top:15px;'>
</div>
""", unsafe_allow_html=True)

# ============================
#  UPLOAD EXCEL
# ============================
uploaded_files = st.file_uploader(
    "Upload Excel Files",
    type=["xlsx"],
    accept_multiple_files=True
)

# Urutkan file dari terbaru → teratas
if uploaded_files:
    uploaded_files = list(uploaded_files)[::-1]

# ============================================
#  PROSES MERGE EXCEL + PROGRESS BAR
# ============================================
if uploaded_files:

    st.info("Uploading & Processing Files...")

    total = len(uploaded_files)
    progress = st.progress(0)
    status_text = st.empty()

    merged_list = []
    total_rows_all = 0   # <- tambahan hitung total baris

    for i, file in enumerate(uploaded_files):

        status_text.write(f"Processing: **{i+1}/{total}**")

        # Baca semua sheet
        excel_data = pd.read_excel(file, sheet_name=None)

        # Gabungkan semua sheet dari file tersebut
        for sheet_name, df in excel_data.items():

            # Hitung jumlah baris tanpa header
            total_rows_all += max(len(df) - 1, 0)

            df["Source File"] = file.name
            df["Sheet"] = sheet_name
            merged_list.append(df)

        progress.progress((i + 1) / total)

    # ======================================
    # SHOW "Successfully processed..." + ROWS
    # ======================================
    status_text.write(
        f"✅ Successfully processed {total} files, {total_rows_all} total rows"
    )

    progress.progress(1.0)

    # Gabungkan
    merged_df = pd.concat(merged_list, ignore_index=True)

    # Index mulai dari 1
    merged_df.index = merged_df.index + 1

    st.markdown('<div class="small-subheader">Merged Excel Results</div>', unsafe_allow_html=True)
    st.dataframe(merged_df, use_container_width=True)

    # ============================================
    # DOWNLOAD EXCEL MERGED
    # ============================================
    def to_excel(df):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='openpyxl')
        df.to_excel(writer, index=False, sheet_name='Merged Data')
        writer.close()
        return output.getvalue()

    excel_file = to_excel(merged_df)

    st.download_button(
        label="Download Merged Excel",
        data=excel_file,
        file_name="Merged_Excel.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
