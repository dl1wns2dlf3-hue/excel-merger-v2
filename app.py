import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import re
import io
import os
from datetime import datetime

# --- [공통 함수] 시드 파일에서 데이터 추출 (강력한 예외 처리) ---
def load_seed_data(file_obj):
    # 기본 반환 구조 정의
    empty_df = pd.DataFrame(columns=['이름_key', '생년월일_val', '업체_val', '직종_val'])
    if file_obj is None: 
        return empty_df
    
    try:
        # 모든 시트를 읽어옴
        sheets = pd.read_excel(file_obj, sheet_name=None, header=None, engine='openpyxl')
    except Exception as e:
        st.error(f"파일을 읽는 중 오류 발생 ({file_obj.name}): {e}")
        return empty_df
        
    df_list = []
    for sheet_name, df_raw in sheets.items():
        if df_raw is None or df_raw.empty: continue
            
        header_idx = -1
        # 상위 20행 내에서 헤더 검색 (범위 확장)
        for idx, row in df_raw.head(20).iterrows():
            row_clean = [str(cell).replace(' ', '').replace('\n', '').strip() for cell in row]
            if '성명' in row_clean or '이름' in row_clean:
                header_idx = idx
                break
        
        if header_idx != -1:
            df = df_raw.iloc[header_idx + 1:].copy()
            # 열 이름 정제 (특수문자 제거)
            cols = [str(c).replace(' ', '').replace('\n', '').strip() for c in df_raw.iloc[header_idx].values]
            df.columns = cols
            df = df.loc[:, ~df.columns.duplicated()].copy()
            df_list.append(df)
            
    if not df_list: 
        return empty_df
    
    combined_df = pd.concat(df_list, ignore_index=True)
    
    # 필요한 열(이름)이 있는지 확인
    name_col = next((c for c in ['이름', '성명'] if c in combined_df.columns), None)
    if not name_col: 
        return empty_df
    
    result_df = pd.DataFrame()
    # 이름 정제 (성명 매칭 핵심)
    result_df['이름_key'] = combined_df[name_col].astype(str).str.replace(r'\s+', '', regex=True).str.strip()
    
    # 생년월일 정제 함수
    def clean_date(x):
        if pd.isna(x) or str(x).strip() in ['', 'nan', 'NaT', 'None']: return ''
        s = str(x).strip()
        if s.endswith('.0'): s = s[:-2]
        s = s.split(' ')[0]
        return re.sub(r'\D', '', s)

    # 생년월일 추출
    if '생년월일' in combined_df.columns:
        result_df['생년월일_val'] = combined_df['생년월일'].apply(clean_date)
    else:
        result_df['생년월일_val'] = ''
        
    # 업체명 추출 (업체명, 업체, 소속)
    comp_col = next((c for c in ['업체명', '업체', '소속'] if c in combined_df.columns), None)
    result_df['업체_val'] = combined_df[comp_col].astype(str).str.strip().replace('nan', '') if comp_col else ''
    
    # 직종명 추출 (직종명, 직종, 공종, 직책)
    job_col = next((c for c in ['직종명', '직종', '공종', '직책'] if c in combined_df.columns), None)
    result_df['직종_val'] = combined_df[job_col].astype(str).str.strip().replace('nan', '') if job_col else ''
    
    # 유효한 이름이 있는 데이터만 남김
    result_df = result_df[result_df['이름_key'].notna() & (result_df['이름_key'] != '') & (result_df['이름_key'] != 'nan')]
    return result_df

# --- [웹 UI 설정] ---
st.set_page_config(page_title="근로자 데이터 병합기", page_icon="👷", layout="centered")

st.title("👷 근로자 데이터 자동 병합 시스템")
st.info("시드 파일(1번, 2번) 중 하나만 올려도 작동합니다. 타겟 파일의 빈칸을 자동으로 채워줍니다.")

# 파일 업로드 섹션
seed1_file = st.file_uploader("1. 첫 번째 시드 파일 (선택)", type=["xlsx"])
seed2_file = st.file_uploader("2. 두 번째 시드 파일 (선택)", type=["xlsx"])
target_file = st.file_uploader("3. 작성하려는 타겟 파일 (필수)", type=["xlsx"])

if st.button("데이터 병합 실행 🚀"):
    # 최소 조건: 타겟 파일 필수 + 시드 파일 중 최소 하나 필수
    if not target_file:
        st.error("작성하려는 타겟 파일을 업로드해주세요.")
    elif not seed1_file and not seed2_file:
        st.error("데이터를 가져올 시드 파일(1번 혹은 2번)을 최소 하나는 올려주세요.")
    else:
        with st.spinner('데이터를 분석하고 병합하는 중입니다...'):
            try:
                # 1. 시드 데이터 로드 및 통합
                df_s1 = load_seed_data(seed1_file)
                df_s2 = load_seed_data(seed2_file)
                
                # 두 시드 데이터를 하나로 합침
                df_seeds = pd.concat([df_s1, df_s2], ignore_index=True)
                
                if df_seeds.empty:
                    st.error("시드 파일에서 유효한 명단 데이터를 찾지 못했습니다. 파일의 '이름' 혹은 '성명' 열을 확인하세요.")
                    st.stop()
                
                # 동명이인 처리: 정보가 더 많은 행을 우선순위로 둠
                df_seeds['score'] = (df_seeds['생년월일_val'] != '').astype(int) + \
                                   (df_seeds['업체_val'] != '').astype(int) + \
                                   (df_seeds['직종_val'] != '').astype(int)
                df_seeds = df_seeds.sort_values('score', ascending=False).drop_duplicates(subset=['이름_key'], keep='first')

                # 2. 타겟 파일 처리 (openpyxl로 직접 수정)
                wb = openpyxl.load_workbook(target_file)
                target_sheet = None
                header_row_idx = 1
                col_indices = {}
                
                # 이름이 있는 시트 찾기
                for sheetname in wb.sheetnames:
                    ws = wb[sheetname]
                    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=20, values_only=True), 1):
                        row_strs = [str(cell).replace(' ', '').replace('\n', '').strip() if cell else '' for cell in row]
                        if '이름' in row_strs or '성명' in row_strs:
                            target_sheet = ws
                            header_row_idx = row_idx
                            for col_idx, val in enumerate(row_strs, 1):
                                if val: col_indices[val] = col_idx
                            break
                    if target_sheet: break
                        
                if not target_sheet:
                    st.error("타겟 파일에서 '이름' 또는 '성명' 열이 있는 시트를 찾지 못했습니다.")
                    st.stop()

                # 타겟 파일의 열 이름 매칭
                name_key = '이름' if '이름' in col_indices else '성명'
                t_dob = '생년월일' if '생년월일' in col_indices else None
                t_comp = next((k for k in ['업체명', '업체', '소속'] if k in col_indices), None)
                t_job = next((k for k in ['직종명', '직종', '공종', '직책'] if k in col_indices), None)

                # 데이터 채우기 시작
                fill_count = 0
                for r in range(header_row_idx + 1, target_sheet.max_row + 1):
                    name_val = target_sheet.cell(row=r, column=col_indices[name_key]).value
                    if not name_val: continue
                    
                    name = str(name_val).replace(' ', '').replace('\n', '').strip()
                    if not name or name == 'None': continue
                    
                    # 시드 데이터에서 매칭 시도
                    match = df_seeds[df_seeds['이름_key'] == name]
                    if not match.empty:
                        s_row = match.iloc[0]
                        # 빈칸인 경우에만 시드 데이터로 채움
                        if t_dob:
                            cell = target_sheet.cell(row=r, column=col_indices[t_dob])
                            if not cell.value or str(cell.value).strip() == '':
                                cell.value = s_row['생년월일_val']
                                fill_count += 1
                        if t_comp:
                            cell = target_sheet.cell(row=r, column=col_indices[t_comp])
                            if not cell.value or str(cell.value).strip() == '':
                                cell.value = s_row['업체_val']
                        if t_job:
                            cell = target_sheet.cell(row=r, column=col_indices[t_job])
                            if not cell.value or str(cell.value).strip() == '':
                                cell.value = s_row['직종_val']

                # 3. 메모리 파일로 저장 및 다운로드 제공
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                today = datetime.now().strftime("%Y%m%d")
                st.success(f"✅ 병합 완료! (총 {fill_count}개의 항목을 확인/업데이트 했습니다.)")
                st.download_button(
                    label="📥 병합된 결과 파일 다운로드",
                    data=output,
                    file_name=f"{today}_병합결과_{target_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"처리 중 예상치 못한 오류가 발생했습니다: {e}")
