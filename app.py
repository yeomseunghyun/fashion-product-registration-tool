import hashlib
import io
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="패션 상품 등록 관리 도구", page_icon="🧥", layout="wide")

REQUIRED_COLUMNS = ["상품코드", "상품명", "브랜드", "카테고리", "색상", "사이즈", "판매가", "소비자가", "이미지파일명"]
OPTIONAL_COLUMNS = ["소재", "제조국", "시즌", "상품설명"]
ALL_COLUMNS = REQUIRED_COLUMNS + OPTIONAL_COLUMNS

SAMPLE = pd.DataFrame([
    ["TS001", "오버핏 코튼 티셔츠", "AMPLE", "상의/반팔티", "화이트", "M", 39000, 49000, "TS001_white_01.jpg", "면 100%", "대한민국", "2026SS", "여유로운 실루엣의 코튼 티셔츠"],
    ["JK002", "데님 워크 재킷", "AMPLE", "아우터/재킷", "인디고", "L", 129000, 149000, "", "면 100%", "대한민국", "2026FW", "데일리 데님 워크 재킷"],
    ["PT003", "와이드 치노 팬츠", "AMPLE", "하의/팬츠", "베이지", "", 79000, 69000, "PT003_beige_01.jpg", "면 혼방", "중국", "2026FW", ""],
], columns=ALL_COLUMNS)

def read_upload(uploaded):
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        try:
            return pd.read_csv(uploaded, dtype={"상품코드": str})
        except UnicodeDecodeError:
            uploaded.seek(0)
            return pd.read_csv(uploaded, encoding="cp949", dtype={"상품코드": str})
    return pd.read_excel(uploaded, dtype={"상품코드": str})

def normalize_frame(df):
    df = df.copy()
    for col in ALL_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[ALL_COLUMNS]
    for col in ALL_COLUMNS:
        df[col] = df[col].fillna("")
    for col in ["판매가", "소비자가"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    return df

def validate(df):
    findings = []
    for idx, row in df.iterrows():
        excel_row = idx + 2
        code = str(row.get("상품코드", "")).strip()
        for col in ["상품코드", "상품명", "브랜드", "카테고리", "색상", "사이즈", "판매가", "소비자가", "이미지파일명"]:
            val = row.get(col, "")
            if pd.isna(val) or str(val).strip() == "":
                findings.append({"행": excel_row, "상품코드": code, "항목": col, "유형": "오류", "내용": f"{col} 값이 비어 있습니다."})
        if code and df["상품코드"].astype(str).str.strip().duplicated(keep=False).iloc[idx]:
            findings.append({"행": excel_row, "상품코드": code, "항목": "상품코드", "유형": "오류", "내용": "중복된 상품코드입니다."})
        price, retail = row.get("판매가"), row.get("소비자가")
        if pd.notna(price) and pd.notna(retail):
            if price < 0 or retail < 0:
                findings.append({"행": excel_row, "상품코드": code, "항목": "가격", "유형": "오류", "내용": "가격은 0 이상이어야 합니다."})
            if price > retail:
                findings.append({"행": excel_row, "상품코드": code, "항목": "판매가", "유형": "경고", "내용": "판매가가 소비자가보다 높습니다. 가격을 확인하세요."})
        if code and not re.fullmatch(r"[A-Za-z0-9_-]+", code):
            findings.append({"행": excel_row, "상품코드": code, "항목": "상품코드", "유형": "경고", "내용": "상품코드에 영문·숫자·_·- 외 문자가 포함되어 있습니다."})
    return pd.DataFrame(findings, columns=["행", "상품코드", "항목", "유형", "내용"])

def to_excel(df, sheet_name="상품등록"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for col_cells in ws.columns:
            max_len = max(len(str(c.value or "")) for c in col_cells)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max(max_len + 2, 12), 40)
        # Show comma-separated thousands for price columns while keeping numeric values.
        headers = {cell.value: cell.column for cell in ws[1]}
        for price_col in ("판매가", "소비자가"):
            if price_col in headers:
                for row in range(2, ws.max_row + 1):
                    ws.cell(row=row, column=headers[price_col]).number_format = '#,##0'
    output.seek(0)
    return output.getvalue()

if "products" not in st.session_state:
    st.session_state.products = SAMPLE.copy()

st.title("🧥 패션 상품 등록 관리 도구")
st.caption("상품 정보 정리 · 입력 오류 검수 · 쇼핑몰 업로드용 엑셀 생성")
st.info("현재 버전은 업로드 파일을 로컬 세션에서 처리합니다. 쇼핑몰에 직접 등록하거나 외부 서버로 전송하지 않습니다.")

with st.sidebar:
    st.header("데이터 가져오기")
    uploaded = st.file_uploader("상품 엑셀 또는 CSV 업로드", type=["xlsx", "xls", "csv"])
    if uploaded is not None:
        upload_bytes = uploaded.getvalue()
        upload_id = hashlib.sha256(upload_bytes).hexdigest()
        if upload_id != st.session_state.get("last_uploaded_id"):
            try:
                loaded = normalize_frame(read_upload(uploaded))
                st.session_state.products = loaded
                st.session_state.last_uploaded_id = upload_id
                st.success(f"{len(loaded)}개 상품을 불러왔습니다.")
            except Exception as exc:
                st.error(f"파일을 읽지 못했습니다: {exc}")
    if st.button("샘플 데이터로 초기화", use_container_width=True):
        st.session_state.products = SAMPLE.copy()
        st.rerun()
    st.download_button("입력용 엑셀 양식 다운로드", data=to_excel(pd.DataFrame(columns=ALL_COLUMNS)), file_name="상품등록_입력양식.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    st.divider()
    st.caption("필수 열 이름은 입력 양식과 동일하게 유지해 주세요.")

products = normalize_frame(st.session_state.products)
issues = validate(products)
errors = int((issues["유형"] == "오류").sum()) if not issues.empty else 0
warnings = int((issues["유형"] == "경고").sum()) if not issues.empty else 0
valid_rows = max(len(products) - products.index.isin(issues.loc[issues["유형"] == "오류", "행"] - 2).sum(), 0) if not issues.empty else len(products)

c1, c2, c3, c4 = st.columns(4)
c1.metric("전체 상품", f"{len(products):,}개")
c2.metric("검수 오류", f"{errors:,}건")
c3.metric("검수 경고", f"{warnings:,}건")
c4.metric("오류 없는 상품", f"{len(products) - len(set((issues.loc[issues['유형']=='오류','행'] - 2).tolist())) if not issues.empty else len(products):,}개")

tab1, tab2, tab3 = st.tabs(["상품 정보 관리", "오류 검수", "파일 내보내기"])
with tab1:
    st.subheader("상품 목록")
    st.caption("셀을 직접 수정할 수 있습니다. 새 상품은 표 마지막 행에 추가하세요.")
    edited = st.data_editor(
        products,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "판매가": st.column_config.NumberColumn("판매가(원)", min_value=0, step=1000, format="%,d"),
            "소비자가": st.column_config.NumberColumn("소비자가(원)", min_value=0, step=1000, format="%,d"),
            "상품설명": st.column_config.TextColumn("상품설명", width="large"),
        },
        key="product_editor",
    )
    st.session_state.products = normalize_frame(edited)
    st.caption("상품코드는 영문, 숫자, 밑줄(_), 하이픈(-) 조합을 권장합니다.")

with tab2:
    st.subheader("검수 결과")
    current = normalize_frame(st.session_state.products)
    current_issues = validate(current)
    if current_issues.empty:
        st.success("검수 오류 및 경고가 없습니다.")
    else:
        filt = st.multiselect("표시할 유형", ["오류", "경고"], default=["오류", "경고"])
        st.dataframe(current_issues[current_issues["유형"].isin(filt)], use_container_width=True, hide_index=True)
        if (current_issues["유형"] == "오류").any():
            st.warning("오류가 남아 있는 상품은 업로드 파일에서 자동 제외할 수 있습니다. 경고는 확인 후 MD가 판단하세요.")
    st.download_button("검수 결과 다운로드", data=to_excel(current_issues, "검수결과"), file_name="상품정보_검수결과.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", disabled=current_issues.empty)

with tab3:
    st.subheader("업로드 파일 생성")
    current = normalize_frame(st.session_state.products)
    current_issues = validate(current)
    exclude_errors = st.checkbox("오류가 있는 상품은 내보내기에서 제외", value=True)
    if st.checkbox("선택한 열만 내보내기", value=False):
        export_cols = st.multiselect("내보낼 열", ALL_COLUMNS, default=REQUIRED_COLUMNS)
    else:
        export_cols = ALL_COLUMNS
    bad_indices = set((current_issues.loc[current_issues["유형"] == "오류", "행"] - 2).tolist()) if not current_issues.empty else set()
    export_df = current.drop(index=list(bad_indices), errors="ignore") if exclude_errors else current
    export_df = export_df[export_cols]
    st.write(f"내보낼 상품: **{len(export_df)}개** / 전체 {len(current)}개")
    if exclude_errors and bad_indices:
        st.warning(f"오류가 있는 상품 {len(bad_indices)}개는 제외됩니다.")
    st.download_button(
        "업로드용 엑셀 다운로드",
        data=to_excel(export_df),
        file_name=f"상품등록_업로드용_{datetime.now():%Y%m%d_%H%M}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        disabled=export_df.empty,
    )
    st.caption("플랫폼별 실제 업로드 열 구조와 값 규칙은 해당 쇼핑몰 양식에 맞게 별도 매핑해야 합니다.")

st.divider()
st.caption("포트폴리오용 MVP · 실제 쇼핑몰 업로드 전에는 해당 플랫폼의 최신 등록 양식으로 검증하세요.")
