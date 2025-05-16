import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Kiểm tra TTK In lỗi", layout="wide")
st.title("📘 Kiểm tra TTK In hỏng & Hết dòng từ File Excel")

uploaded_file = st.file_uploader("📤 Upload file Excel (như Muc18_1504_GTCG.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, dtype={'ACC_NO': str})
    df['INVT_TRAN_DATE'] = pd.to_datetime(df['INVT_TRAN_DATE'], errors='coerce')
    df['INVT_TRAN_DATE_STR'] = df['INVT_TRAN_DATE'].dt.strftime('%m/%d/%Y')
    df = df.sort_values(by='INVT_SRL_NUM', ascending=True)

    # In hỏng
    df_hong = df[df['PASSBOOK_STATUS'] == 'F'].copy()
    so_lan_hong = df_hong.groupby('ACC_NO').size().reset_index(name='SỐ_LẦN_IN_HỎNG')
    df_hong['ACC_DATE'] = df_hong['ACC_NO'] + '|' + df_hong['INVT_TRAN_DATE_STR']
    fail_ngay = df_hong.groupby('ACC_DATE').size().reset_index(name='FAIL_COUNT')
    fail_ngay['TTK IN HỎNG NHIỀU TRONG 01 NGÀY'] = fail_ngay['FAIL_COUNT'].apply(lambda x: 'X' if x >= 2 else '')
    fail_ngay[['ACC_NO', 'INVT_TRAN_DATE_STR']] = fail_ngay['ACC_DATE'].str.split('|', expand=True)

    # In hết dòng
    df_hetdong = df[df['PASSBOOK_STATUS'] == 'U'].copy()
    so_lan_hetdong = df_hetdong.groupby('ACC_NO').size().reset_index(name='SỐ_LẦN_IN_HẾT_DÒNG')
    df_hetdong['ACC_DATE'] = df_hetdong['ACC_NO'] + '|' + df_hetdong['INVT_TRAN_DATE_STR']
    hetdong_ngay = df_hetdong.groupby('ACC_DATE').size().reset_index(name='HETDONG_COUNT')
    hetdong_ngay['TTK IN HẾT DÒNG NHIỀU TRONG 01 NGÀY'] = hetdong_ngay['HETDONG_COUNT'].apply(lambda x: 'X' if x >= 2 else '')
    hetdong_ngay[['ACC_NO', 'INVT_TRAN_DATE_STR']] = hetdong_ngay['ACC_DATE'].str.split('|', expand=True)

    # Vừa in hỏng vừa hết dòng
    merged_hh = pd.merge(
        fail_ngay[['ACC_NO', 'INVT_TRAN_DATE_STR']],
        hetdong_ngay[['ACC_NO', 'INVT_TRAN_DATE_STR']],
        on=['ACC_NO', 'INVT_TRAN_DATE_STR'],
        how='inner'
    )
    merged_hh['TTK VỪA IN HỎNG VỪA HẾT DÒNG TRONG 01 NGÀY'] = 'X'

    # Tóm tắt
    summary = df[['ACC_NO', 'INVT_TRAN_DATE_STR', 'PASSBOOK_STATUS']].drop_duplicates('ACC_NO')
    summary = pd.merge(summary, so_lan_hong, on='ACC_NO', how='left')
    summary = pd.merge(summary, fail_ngay[['ACC_NO', 'INVT_TRAN_DATE_STR', 'TTK IN HỎNG NHIỀU TRONG 01 NGÀY']],
                       on=['ACC_NO', 'INVT_TRAN_DATE_STR'], how='left')
    summary = pd.merge(summary, so_lan_hetdong, on='ACC_NO', how='left')
    summary = pd.merge(summary, hetdong_ngay[['ACC_NO', 'INVT_TRAN_DATE_STR', 'TTK IN HẾT DÒNG NHIỀU TRONG 01 NGÀY']],
                       on=['ACC_NO', 'INVT_TRAN_DATE_STR'], how='left')
    summary = pd.merge(summary, merged_hh[['ACC_NO', 'INVT_TRAN_DATE_STR', 'TTK VỪA IN HỎNG VỪA HẾT DÒNG TRONG 01 NGÀY']],
                       on=['ACC_NO', 'INVT_TRAN_DATE_STR'], how='left')

    # Xuất ra file Excel trong bộ nhớ
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Raw dữ liệu đầy đủ', index=False)
        summary.to_excel(writer, sheet_name='Tóm tắt In lỗi', index=False)
    output.seek(0)

    st.success("✅ Xử lý hoàn tất!")
    st.download_button(
        label="📥 Tải về file kết quả",
        data=output,
        file_name="output_TTK_in_hong_het_dong.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
