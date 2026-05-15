import streamlit as st
import pandas as pd
import openpyxl
import io
import re
from datetime import datetime

# --- [전처리] 메모리 점유율을 최소화하여 엑셀 읽기 ---
def load_data_fast(file_obj):
    if file_obj is None:
        return pd.DataFrame()
    try:
        # 파일을 읽을 때 상위 5000줄까지만 제한하여 유령 데이터 방지
        # engine='openpyxl'을 명시하고, 수식이 아닌 결과값만 가져오도록 설정
        df = pd.read_excel(file_obj, header=None, engine='openpyxl', nrows=5000)
        
        # 1. 헤더(이름/성명) 위치 찾기
        header_idx = -1
        for idx, row in df.head(20).iterrows():
            row_vals = [str(c).replace(' ', '').strip() for c in row if pd.notna(c)]
            if '성명' in row_vals or '이름' in row_vals:
                header_idx = idx
                break
        
        if header_idx == -1:
            return pd.DataFrame()

        # 2. 헤더 기준 데이터 재설정
        df.columns = [str(c).replace(' ', '').strip() for c in df.iloc[header_idx]]
        df = df.iloc[header_idx + 1:].reset_index(drop=True)
        
        # 3. 필요한 열만 필터링 (메모리 절약)
        name_col = next((c for c in ['이름', '성명'] if c in df.columns), None)
        if not name_col:
            return pd.DataFrame()

        res = pd.DataFrame()
        res['이름_key'] = df[name_col].astype(str).str.replace(r'\s+', '', regex=True).str.strip()
        
        # 생년월일/업체/직종 추출 (열이 있을 때만)
        dob_col = '생년월일' if '생년월일' in df.columns else None
        comp_col = next((c for c in ['업체명', '업체', '소속'] if c in df.columns), None)
        job_col = next((c for c in ['직종명', '직종', '공종', '직책'] if c in df.columns), None)

        if dob_col:
            res['생년월일_val'] = df[dob_col].astype(str).apply(lambda x: re.sub(r'\D', '', x.split(' ')[0]) if x != 'nan' else '')
        else: res['생년월일_val'] = ''
        
        res['업체_val'] = df[comp_col].astype(str).str.strip().replace('nan', '') if comp_col else ''
        res['직종_val'] = df[job_col].astype(str).str.strip().replace('nan', '') if job_col else ''

        # 유효한 이름 데이터만 남기기
        return res[res['이름_key'].str.len() > 1].drop_duplicates('이름_key')
    except Exception as e:
        st.error(f"파일 처리 실패: {e}")
        return pd.DataFrame()

# --- [UI] ---
st.set_page_config(page_title="현장 데이터 병합기", layout="centered")
st.title("👷 건설현장 근로자 데이터 병합")
st.info("5MB 이상의 무거운 파일도 최적화하여 읽어옵니다. 시드 파일은 1번이나 2번 중 하나만 있어도 작동합니다.")

s1 = st.file_uploader("1번 시드 (정보 원본)", type=["xlsx"])
s2 = st.file_uploader("2번 시드 (추가 정보)", type=["xlsx"])
target = st.file_uploader("타겟 파일 (빈칸 채울 양식)", type=["xlsx"])

if st.button("병합 시작 🚀"):
    if not target or (not s1 and not s2):
        st.warning("파일을 업로드해주세요.")
    else:
        with st.spinner('대용량 파일 최적화 중...'):
            # 데이터 로드
            df1 = load_data_fast(s1)
            df2 = load_data_fast(s2)
            seeds = pd.concat([df1, df2]).drop_duplicates('이름_key', keep='first')

            if seeds.empty:
                st.error("시드 파일에서 명단을 찾을 수 없습니다. 파일 양식을 확인해주세요.")
                st.stop()

            # 타겟 파일 처리 (Openpyxl 직접 수정)
            wb = openpyxl.load_workbook(target, data_only=True)
            ws = wb.active # 첫 번째 시트 사용
            
            # 헤더 찾기
            h_idx = -1
            cols = {}
            for r in range(1, 21):
                row = [str(ws.cell(r, c).value).replace(' ', '').strip() for c in range(1, ws.max_column + 1)]
                if '이름' in row or '성명' in row:
                    h_idx = r
                    for i, val in enumerate(row, 1):
                        if val and val != 'None': cols[val] = i
                    break

            if h_idx == -1:
                st.error("타겟 양식에서 '이름' 열을 찾지 못했습니다.")
                st.stop()

            n_k = '이름' if '이름' in cols else '성명'
            t_d = '생년월일' if '생년월일' in cols else None
            t_c = next((k for k in ['업체명', '업체', '소속'] if k in cols), None)
            t_j = next((k for k in ['직종명', '직종', '공종', '직책'] if k in cols), None)

            count = 0
            for r in range(h_idx + 1, ws.max_row + 1):
                name_cell = ws.cell(r, cols[n_k]).value
                if not name_cell: continue
                name = str(name_cell).replace(' ', '').strip()
                
                match = seeds[seeds['이름_key'] == name]
                if not match.empty:
                    res = match.iloc[0]
                    if t_d and not ws.cell(r, cols[t_d]).value: ws.cell(r, cols[t_d]).value = res['생년월일_val']
                    if t_c and not ws.cell(r, cols[t_c]).value: ws.cell(r, cols[t_c]).value = res['업체_val']
                    if t_j and not ws.cell(r, cols[t_j]).value: ws.cell(r, cols[t_j]).value = res['직종_val']
                    count += 1

            out = io.BytesIO()
            wb.save(out)
            
            st.success(f"✅ {count}명 매칭 완료!")
            st.download_button(
                label="📥 결과 다운로드", 
                data=out.getvalue(), 
                file_name=f"병합완료_{datetime.now().strftime('%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
