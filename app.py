import streamlit as st
import pandas as pd
import io

def convert_excel(input_file, input_format, output_format):
    try:
        if input_format == '.xls':
            df = pd.read_excel(input_file, engine='xlrd')
        elif input_format in ['.xlsx', '.xlsb', '.xlsm']:
            df = pd.read_excel(input_file, engine='openpyxl')
        else:
            st.error("Format input tidak didukung.")
            return None

        output = io.BytesIO()
        if output_format == '.xlsx':
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
        elif output_format == '.xls':
            with pd.ExcelWriter(output, engine='xlwt') as writer:
                df.to_excel(writer, index=False)
        elif output_format in ['.xlsb', '.xlsm']:
            with pd.ExcelWriter(output, engine='openpyxl', mode='xlsb' if output_format == '.xlsb' else 'xlsm') as writer:
                df.to_excel(writer, index=False)
        else:
            st.error("Format output tidak didukung.")
            return None

        return output.getvalue()
    except Exception as e:
        st.error(f"Terjadi error saat konversi: {str(e)}")
        return None

st.title("Excel File Converter")

input_format = st.selectbox("Pilih format input:", ['.xls', '.xlsx', '.xlsb', '.xlsm'])
output_format = st.selectbox("Pilih format output:", ['.xls', '.xlsx', '.xlsb', '.xlsm'])

uploaded_file = st.file_uploader("Pilih file Excel", type=['xls', 'xlsx', 'xlsb', 'xlsm'])

if uploaded_file is not None:
    if st.button("Konversi"):
        with st.spinner('Sedang mengkonversi file...'):
            converted_file = convert_excel(uploaded_file, input_format, output_format)
        if converted_file:
            st.success("Konversi berhasil!")
            st.download_button(
                label="Download file hasil konversi",
                data=converted_file,
                file_name=f"converted{output_format}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Gagal mengkonversi file. Pastikan format input dan output yang dipilih benar.")
