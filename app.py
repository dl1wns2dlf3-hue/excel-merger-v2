import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import re
import io
import os
from datetime import datetime

# --- [공통 함수] 시드 파일에서 데이터 추출 ---
def load_seed_data(file_obj):
    empty_df = pd.DataFrame(columns=['이름_key', '생년월일_val', '업체_val', '직종_val'])
    if file_obj is None: return empty_df
    
    try:
        sheets = pd.read_excel(file_obj, sheet_name=None, header=None)
    except Exception:
        return empty_df
        
    df_list = []
    for sheet_name, df_raw in sheets.items():
        if df_raw is None or df_raw.empty: continue
            
        header_idx = -1
        for idx, row in df_raw.head(15).iterrows():
            row_clean = [str(cell).replace(' ', '').replace('\n', '').strip() for cell in row]
            if '성명' in row_clean or '이름' in row_clean:
                header_idx = idx
                break
        
        if header_idx != -1:
            df = df_raw.iloc[header_idx + 1:].copy()
            cols = [str(c).replace(' ', '').replace('\n', '').strip() for c in df_raw.iloc[header_idx].values]
            df.columns = cols
            df = df.loc[:, ~df.columns.duplicated()].copy()
            df_list.append(df)
            
    if not df_list: return empty_df
    
    combined_df = pd.concat(df_list, ignore_index=True)
    if combined_df.empty: return empty_df
    
    result_df = pd.DataFrame()
    name_col = next((c for c in ['이름', '성명'] if c in combined_df.columns), None)
    if not name_col: return empty_df
    
    result_df['이름_key'] = combined_df[name_col].astype(str).str.replace(r'\s+', '', regex=True)
    
    def clean_date(x):
        if pd.isna(x) or str(x).strip() in ['', 'nan', 'NaT', 'None']: return ''
        s = str(x).strip()
        if s.endswith('.0'): s = s[:-2]
        s = s.split(' ')[0]
        return re.sub(r'\D', '', s)

    if '생년월일' in combined_df.columns:
        result_df['생년월일_val'] = combined_df['생년월일'].apply(clean_date)
    else: result_df['생년월일_val'] = ''
        
    comp_col = next((c for c in ['업체명', '업체', '소속'] if c in combined_df.columns), None)
    result_df['업체_val'] = combined_df[comp_col].astype(str).str.strip().replace('nan', '') if comp_col else ''
    
    job_col = next((c for c in ['직종명', '직종', '공종', '직책'] if c in combined_df.columns), None)
    result_df['직종_val'] = combined_df[job_col].astype(str).str.strip().replace('nan', '') if job_col else ''
    
    result_df = result_df[(result_df['이름_key'] != '') & (result_df['이름_key'] != 'nan')]
    return result_df

# --- [웹 UI] ---
st.set_page_config(page_title="근로자 데이터 병합기", page_icon="👷", layout="centered")

st.title("👷 근로자 데이터 자동 병합 시스템")
st.markdown("엑셀 파일을 업로드하면 **이름을 기준**으로 정보를 자동으로 채워줍니다.")

seed1_file = st.file_uploader("1. 첫 번째 시드 파일 (필수)", type=["xlsx"])
seed2_file = st.file_uploader("2. 두 번째 시드 파일 (선택)", type=["xlsx"])
target_file = st.file_uploader("3. 작성하려는 원본 파일 (필수)", type=["xlsx"])

if st.button("데이터 병합 실행 🚀"):
    if not target_file or not seed1_file:
        st.warning("파일을 모두 업로드해주세요.")
    else:
        with st.spinner('처리 중...'):
            try:
                df_s1 = load_seed_data(seed1_file)
                df_s2 = load_seed_data(seed2_file)
                df_seeds = pd.concat([df_s1, df_s2], ignore_index=True)
                
                if df_seeds.empty:
                    st.error("데이터를 찾지 못했습니다.")
                    st.stop()
                
                df_seeds['score'] = (df_seeds['생년월일_val'] != '').astype(int) + (df_seeds['업체_val'] != '').astype(int) + (df_seeds['직종_val'] != '').astype(int)
                df_seeds = df_seeds.sort_values('score', ascending=False).drop_duplicates(subset=['이름_key'], keep='first')

                wb = openpyxl.load_workbook(target_file)
                target_sheet = None
                header_row_idx = 1
                col_indices = {}
                
                for sheetname in wb.sheetnames:
                    ws = wb[sheetname]
                    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=15, values_only=True), 1):
                        row_strs = [str(cell).replace(' ', '').replace('\n', '').strip() if cell else '' for cell in row]
                        if '이름' in row_strs or '성명' in row_strs:
                            target_sheet = ws
                            header_row_idx = row_idx
                            for col_idx, val in enumerate(row_strs, 1):
                                if val: col_indices[val] = col_idx
                            break
                    if target_sheet: break
                        
                if not target_sheet:
                    st.error("양식에서 '이름' 열을 못 찾았습니다.")
                    st.stop()

                name_key = '이름' if '이름' in col_indices else '성명'
                t_dob = '생년월일' if '생년월일' in col_indices else None
                t_comp = next((k for k in ['업체명', '업체', '소속'] if k in col_indices), None)
                t_job = next((k for k in ['직종명', '직종', '공종', '직책'] if k in col_indices), None)

                for r in range(header_row_idx + 1, target_sheet.max_row + 1):
                    name = str(target_sheet.cell(row=r, column=col_indices[name_key]).value or '').replace(' ', '').strip()
                    if not name or name == 'None': continue
                    
                    match = df_seeds[df_seeds['이름_key'] == name]
                    if not match.empty:
                        s_row = match.iloc[0]
                        if t_dob:
                            cell = target_sheet.cell(row=r, column=col_indices[t_dob])
                            if not cell.value: cell.value = s_row['생년월일_val']
                        if t_comp:
                            cell = target_sheet.cell(row=r, column=col_indices[t_comp])
                            if not cell.value: cell.value = s_row['업체_val']
                        if t_job:
                            cell = target_sheet.cell(row=r, column=col_indices[t_job])
                            if not cell.value: cell.value = s_row['직종_val']

                out = io.BytesIO()
                wb.save(out)
                out.seek(0)
                
                st.success("✅ 완료!")
                st.download_button("📥 다운로드", out, f"{datetime.now().strftime('%Y%m%d')}_{target_file.name}", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"오류: {e}")