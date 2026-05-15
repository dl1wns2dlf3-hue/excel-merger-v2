import streamlit as st
import pandas as pd
import openpyxl
import io
import re
from datetime import datetime

# --- [전처리] 어떤 형태의 엑셀 데이터가 들어와도 에러 없이 읽기 ---
def load_data_fast(file_obj):
    if file_obj is None:
        return pd.DataFrame()
    try:
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
        
        # 3. 이름 추출
        name_col = next((c for c in ['이름', '성명'] if c in df.columns), None)
        if not name_col:
            return pd.DataFrame()

        res = pd.DataFrame()
        res['이름_key'] = df[name_col].astype(str).str.replace(r'\s+', '', regex=True).str.strip()
        
        # --- [핵심 수정] 숫자, 날짜, 결측치 완벽 방어 함수 ---
        def clean_date(x):
            if pd.isna(x): return ''
            s = str(x).strip() # 무조건 문자로 강제 변환하여 에러 원천 차단
            if s.lower() in ['nan', 'nat', 'none', '']: return ''
            if s.endswith('.0'): s = s[:-2] # 19970306.0 처럼 소수점이 붙는 경우 제거
            s = s.split(' ')[0] # 날짜 뒤에 시간이 붙어있을 경우 띄어쓰기 기준으로 자름
            return re.sub(r'\D', '', s) # 숫자만 깔끔하게 남김

        def clean_text(x):
            if pd.isna(x): return ''
            s = str(x).strip()
            return '' if s.lower() in ['nan', 'none'] else s

        dob_col = '생년월일' if '생년월일' in df.columns else None
        comp_col = next((c for c in ['업체명', '업체', '소속'] if c in df.columns), None)
        job_col = next((c for c in ['직종명', '직종', '공종', '직책'] if c in df.columns), None)

        # 안전한 함수 적용
        res['생년월일_val'] = df[dob_col].apply(clean_date) if dob_col else ''
        res['업체_val'] = df[comp_col].apply(clean_text) if comp_col else ''
        res['직종_val'] = df[job_col].apply(clean_text) if job_col else ''

        # 유효한 이름(2글자 이상) 데이터만 남기기
        return res[res['이름_key'].str.len() > 1].drop_duplicates('이름_key')
    
    except Exception as e:
        st.error(f"파일 처리 실패: {e}")
        return pd.DataFrame()

# --- [웹 UI] ---
st.set_page_config(page_title="현장 데이터 병합기", layout="centered")
st.title("👷 건설현장 근로자 데이터 자동 병합")
st.info("양식이 다른 여러 개의 명단을 한 번에 합쳐줍니다.")

s1 = st.file_uploader("1번 시드 (필수 아님)", type=["xlsx"])
s2 = st.file_uploader("2번 시드 (필수 아님)", type=["xlsx"])
target = st.file_uploader("타겟 파일 (빈칸을 채울 대상 - 필수)", type=["xlsx"])

if st.button("데이터 병합 실행 🚀"):
    if not target or (not s1 and not s2):
        st.warning("타겟 파일과 최소 1개 이상의 시드 파일을 업로드해주세요.")
    else:
        with st.spinner('서로 다른 엑셀 양식을 분석하고 병합하는 중입니다...'):
            df1 = load_data_fast(s1)
            df2 = load_data_fast(s2)
            
            # 정보가 많은(비어있지 않은) 데이터를 우선순위로 병합
            seeds = pd.concat([df1, df2], ignore_index=True)
            if not seeds.empty:
                seeds['score'] = (seeds['생년월일_val'] != '').astype(int) + (seeds['업체_val'] != '').astype(int) + (seeds['직종_val'] != '').astype(int)
                seeds = seeds.sort_values('score', ascending=False).drop_duplicates('이름_key', keep='first')

            if seeds.empty:
                st.error("시드 파일에서 명단을 찾을 수 없습니다. 파일 양식을 확인해주세요.")
                st.stop()

            # 타겟 파일 처리
            wb = openpyxl.load_workbook(target, data_only=True)
            ws = wb.active 
            
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
                st.error("타겟 양식에서 '이름' 또는 '성명' 열을 찾지 못했습니다.")
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
                    # 빈칸일 경우에만 데이터 채우기
                    if t_d and (not ws.cell(r, cols[t_d]).value or str(ws.cell(r, cols[t_d]).value).strip() == ''):
                        ws.cell(r, cols[t_d]).value = res['생년월일_val']
                    if t_c and (not ws.cell(r, cols[t_c]).value or str(ws.cell(r, cols[t_c]).value).strip() == ''):
                        ws.cell(r, cols[t_c]).value = res['업체_val']
                    if t_j and (not ws.cell(r, cols[t_j]).value or str(ws.cell(r, cols[t_j]).value).strip() == ''):
                        ws.cell(r, cols[t_j]).value = res['직종_val']
                    count += 1

            out = io.BytesIO()
            wb.save(out)
            
            st.success(f"✅ 총 {count}명의 데이터를 성공적으로 매칭하여 채웠습니다!")
            st.download_button(
                label="📥 병합된 결과 엑셀 다운로드", 
                data=out.getvalue(), 
                file_name=f"병합결과_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
