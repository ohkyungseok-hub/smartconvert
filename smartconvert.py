from io import BytesIO
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="스마트스토어 일괄등록 변환기", layout="wide")

# =========================
# 고정 파일 (앱 폴더에 둔다고 가정)
# =========================
TEMPLATE_PATH = "ExcelSaveTemplate_250311.xlsx"
CATEGORY_MASTER_PATH = "category_20260102_012842.xls"

# =========================
# 고정값
# =========================
AS_PHONE_DEFAULT = "02-6953-4020"
AS_INFO_DEFAULT = "평일 오전 10시부터 오후 6시까지"
DEFAULT_CATEGORY_CODE = "50003307"  # 매칭 실패 시 기본 카테고리코드(원하시면 변경)

# =========================
# Utils
# =========================
TOKEN_RE = re.compile(r"[가-힣a-zA-Z0-9]+")

def tokenize(text: str) -> set[str]:
    if text is None:
        return set()
    return set(TOKEN_RE.findall(str(text).lower()))

def clean_main_image(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip().strip("[]")
    if "," in s:
        s = s.split(",")[0]
    return s.strip().strip("'").strip('"').strip()

def find_col_index_contains(labels, keyword: str):
    for j, v in enumerate(labels):
        if keyword in str(v):
            return j
    return None

def find_col_index_exact(labels, name: str):
    for j, v in enumerate(labels):
        if str(v).strip() == name:
            return j
    return None

def is_required_cell(val) -> bool:
    return str(val).replace(" ", "").strip() == "필수"

# =========================
# Load category master (cached)
# =========================
@st.cache_data
def load_category_master(path: str) -> pd.DataFrame:
    """
    category_*.xls 파일을 로드하고,
    category_code / category_name / category_path / tokens 를 만들어요.
    """
    cat = pd.read_excel(path)

    cols = cat.columns.tolist()

    # 컬럼 자동 추정 (파일마다 컬럼명이 조금 다를 수 있어 유연하게)
    code_col = next((c for c in cols if "코드" in c), None)
    name_col = next((c for c in cols if ("카테고리명" in c) or (c.endswith("명"))), None)
    path_col = next((c for c in cols if "경로" in c), None)

    if code_col is None or name_col is None:
        raise ValueError(f"카테고리 파일 컬럼을 인식하지 못했습니다. 현재 컬럼: {cols}")

    use_cols = [code_col, name_col] + ([path_col] if path_col else [])
    cat = cat[use_cols].copy()
    cat.columns = ["category_code", "category_name"] + (["category_path"] if path_col else [])

    if "category_path" not in cat.columns:
        cat["category_path"] = ""

    cat["search_text"] = (cat["category_name"].astype(str) + " " + cat["category_path"].astype(str)).str.lower()
    cat["tokens"] = cat["search_text"].map(tokenize)

    # 비어있는 토큰 제거
    cat = cat[cat["tokens"].map(len) > 0].reset_index(drop=True)
    return cat

# =========================
# Category scoring
# =========================
def top3_categories(product_text: str, cat_df: pd.DataFrame, default_code: str):
    """
    반환:
      candidates: list[dict]  (최대 3개)
      confidence: 0~100 (top1 기준)
    점수:
      overlap = |상품토큰 ∩ 카테고리토큰|
      신뢰도 = overlap / max(1, |상품토큰|) * 100
    """
    prod_tokens = tokenize(product_text)
    if not prod_tokens:
        return [], 0.0

    # 빠르게 top3만 유지
    best = []  # list of tuples(score, idx)
    for i, row in cat_df.iterrows():
        overlap = len(prod_tokens & row["tokens"])
        if overlap <= 0:
            continue
        best.append((overlap, i))

    if not best:
        return [], 0.0

    best.sort(key=lambda x: x[0], reverse=True)
    best3 = best[:3]

    candidates = []
    for score, i in best3:
        r = cat_df.iloc[i]
        candidates.append({
            "code": str(r["category_code"]),
            "name": str(r["category_name"]),
            "path": str(r.get("category_path", "")),
            "score": int(score),
        })

    top_score = candidates[0]["score"]
    confidence = (top_score / max(1, len(prod_tokens))) * 100.0
    confidence = max(0.0, min(100.0, confidence))
    return candidates, confidence

# =========================
# Convert to smartstore template (template embedded)
# =========================
def convert_to_template(
    src_df: pd.DataFrame,
    template_path: str,
    category_final_codes: list[str],
    as_phone: str,
    as_info: str
) -> BytesIO:
    tpl = pd.read_excel(template_path, header=None)
    labels = tpl.iloc[1].astype(str).tolist()

    # 컬럼 인덱스 탐색
    IDX = {
        "seller_code": find_col_index_exact(labels, "판매자 상품코드") or 0,
        "category": find_col_index_exact(labels, "카테고리코드"),
        "name": find_col_index_exact(labels, "상품명"),
        "price": find_col_index_exact(labels, "판매가"),
        "stock": find_col_index_exact(labels, "재고수량"),
        "main_img": find_col_index_exact(labels, "대표이미지"),
        "instant_value": find_col_index_contains(labels, "즉시할인 값"),
        "instant_unit": find_col_index_contains(labels, "즉시할인 단위"),
        "as_phone": find_col_index_contains(labels, "A/S 전화번호"),
        "as_info": find_col_index_contains(labels, "A/S 안내"),
    }

    required_row = tpl.iloc[2]
    example_row = tpl.iloc[3]
    required_cols = [i for i in range(tpl.shape[1]) if is_required_cell(required_row.iat[i])]

    start_row = tpl.shape[0] - 1
    out = tpl.copy()
    out = out.reindex(range(start_row + len(src_df)))

    # 1) 필수값: 예시값으로 채움
    for i in range(len(src_df)):
        r = start_row + i
        for c in required_cols:
            out.iat[r, c] = example_row.iat[c]

    # 2) 값 매핑
    for i, row in src_df.iterrows():
        r = start_row + i

        seller_code = row.get("상품연동코드(옵션)")
        if pd.isna(seller_code) or str(seller_code).strip() == "":
            seller_code = row.get("상품ID")

        cost = row.get("원가(옵션)")
        sell_sum = row.get("판매가(옵션합)")

        out.iat[r, IDX["seller_code"]] = str(seller_code)

        if IDX["category"] is not None:
            out.iat[r, IDX["category"]] = category_final_codes[i] if i < len(category_final_codes) else DEFAULT_CATEGORY_CODE

        if IDX["name"] is not None:
            out.iat[r, IDX["name"]] = row.get("상품명", "")

        # 판매가 = 원가(옵션)
        if IDX["price"] is not None:
            out.iat[r, IDX["price"]] = cost if pd.notna(cost) else ""

        # 즉시할인 값 = 원가(옵션) - 판매가(옵션합)
        if IDX["instant_value"] is not None and pd.notna(cost) and pd.notna(sell_sum):
            out.iat[r, IDX["instant_value"]] = float(cost) - float(sell_sum)

        # 즉시할인 단위 = 원 (값 있을 때만)
        if IDX["instant_unit"] is not None:
            if pd.notna(cost) and pd.notna(sell_sum):
                out.iat[r, IDX["instant_unit"]] = "원"
            else:
                out.iat[r, IDX["instant_unit"]] = ""

        if IDX["stock"] is not None:
            out.iat[r, IDX["stock"]] = row.get("재고수량(옵션)", "")

        if IDX["main_img"] is not None:
            out.iat[r, IDX["main_img"]] = clean_main_image(row.get("대표이미지"))

        # A/S 고정
        if IDX["as_phone"] is not None:
            out.iat[r, IDX["as_phone"]] = as_phone
        if IDX["as_info"] is not None:
            out.iat[r, IDX["as_info"]] = as_info

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, header=False)
    buf.seek(0)
    return buf

# =========================
# UI
# =========================
st.title("스마트스토어 일괄상품등록 변환기 (카테고리 TOP3 + 신뢰도 + 수동보정)")
st.caption("원본 상품내역 엑셀만 업로드하면, 템플릿/카테고리마스터는 기본값으로 자동 사용합니다.")

with st.expander("기본 파일 경로(앱 폴더에 존재해야 함)", expanded=False):
    st.code(
        f"TEMPLATE_PATH = {TEMPLATE_PATH}\nCATEGORY_MASTER_PATH = {CATEGORY_MASTER_PATH}",
        language="text"
    )

left, right = st.columns(2)
with left:
    src_file = st.file_uploader("📂 원본 상품내역 엑셀 업로드(.xlsx)", type=["xlsx"])
with right:
    st.text_input("A/S 전화번호(고정)", value=AS_PHONE_DEFAULT, key="as_phone")
    st.text_input("A/S 안내(고정)", value=AS_INFO_DEFAULT, key="as_info")

if not src_file:
    st.stop()

# 원본 읽기
src_df = pd.read_excel(src_file)

needed_cols = ["상품ID", "상품명", "원가(옵션)", "판매가(옵션합)", "재고수량(옵션)", "대표이미지"]
missing = [c for c in needed_cols if c not in src_df.columns]
if missing:
    st.warning(f"원본 파일에 아래 컬럼이 없습니다(일부 매핑이 공란 처리될 수 있어요): {', '.join(missing)}")

# 카테고리 마스터 로드
try:
    cat_df = load_category_master(CATEGORY_MASTER_PATH)
except Exception as e:
    st.error(f"카테고리 마스터 로드 실패: {e}")
    st.stop()

st.divider()
st.subheader("1) 카테고리 자동분류 결과 (TOP3 + 신뢰도)")

# 세션 상태: 사용자 보정 저장
if "category_overrides" not in st.session_state:
    st.session_state.category_overrides = {}  # {row_index: category_code}

# 분류용 텍스트: 상품명 + 태그(있으면)
def row_text(row):
    name = "" if pd.isna(row.get("상품명")) else str(row.get("상품명"))
    tags = "" if pd.isna(row.get("뱃지/태그")) else str(row.get("뱃지/태그"))
    return f"{name} {tags}".strip()

# 분류 결과 계산 (캐시/메모리 절약: 한 번만 계산)
@st.cache_data
def compute_suggestions(src_df: pd.DataFrame, cat_df: pd.DataFrame, default_code: str):
    rows = []
    for i, row in src_df.iterrows():
        text = row_text(row)
        candidates, conf = top3_categories(text, cat_df, default_code)
        # 후보 문자열 만들기
        def fmt(c):
            return f"{c['code']} | {c['path'] or c['name']}"
        top1 = fmt(candidates[0]) if len(candidates) > 0 else f"{default_code} | (기본값)"
        top2 = fmt(candidates[1]) if len(candidates) > 1 else ""
        top3 = fmt(candidates[2]) if len(candidates) > 2 else ""
        top1_code = candidates[0]["code"] if len(candidates) > 0 else default_code

        rows.append({
            "row_index": i,
            "상품명": row.get("상품명", ""),
            "신뢰도(%)": round(conf, 1),
            "추천1": top1,
            "추천2": top2,
            "추천3": top3,
            "추천1_코드": top1_code,
        })
    return pd.DataFrame(rows)

suggest_df = compute_suggestions(src_df, cat_df, DEFAULT_CATEGORY_CODE)

# 필터/페이징 UI
f1, f2, f3, f4 = st.columns([1.2, 1.2, 1.2, 2.4])
with f1:
    min_conf = st.slider("최소 신뢰도 필터", 0, 100, 0)
with f2:
    only_low = st.checkbox("저신뢰도만 보기", value=False)
with f3:
    page_size = st.selectbox("페이지 크기", [10, 20, 50, 100], index=1)
with f4:
    keyword = st.text_input("상품명 검색", value="")

view = suggest_df.copy()
if keyword.strip():
    view = view[view["상품명"].astype(str).str.contains(keyword.strip(), case=False, na=False)]

if only_low:
    view = view[view["신뢰도(%)"] < max(1, min_conf)]
else:
    view = view[view["신뢰도(%)"] >= min_conf]

total = len(view)
if total == 0:
    st.info("조건에 맞는 항목이 없습니다.")
    st.stop()

# 페이지
max_page = (total - 1) // page_size + 1
page = st.number_input("페이지", min_value=1, max_value=max_page, value=1, step=1)
start = (page - 1) * page_size
end = min(total, start + page_size)
page_df = view.iloc[start:end].copy()

# 미리보기 테이블
st.dataframe(
    page_df[["상품명", "신뢰도(%)", "추천1", "추천2", "추천3"]],
    use_container_width=True,
    hide_index=True
)

st.subheader("2) 수동 보정 (행별 선택)")
st.caption("각 행에서 최종 카테고리를 선택하세요. 선택하지 않으면 추천1이 적용됩니다.")

# 한 페이지 내 행별 선택 UI
for _, r in page_df.iterrows():
    i = int(r["row_index"])
    prod_name = str(r["상품명"])
    conf = r["신뢰도(%)"]

    # 후보 옵션 만들기: 추천1~3 + 기본값
    options = []
    # 추천 문자열에서 code 추출
    def extract_code(s):
        if not s:
            return ""
        return s.split("|")[0].strip()

    recs = [r["추천1"], r["추천2"], r["추천3"]]
    for s in recs:
        if s:
            options.append(s)
    default_opt = f"{DEFAULT_CATEGORY_CODE} | (기본값)"
    if default_opt not in options:
        options.append(default_opt)

    # 현재 선택값(override 있으면 우선)
    current_code = st.session_state.category_overrides.get(i, r["추천1_코드"])
    # options 중 code가 current_code인 걸 찾아 index로 세팅
    def find_option_by_code(code):
        for k, opt in enumerate(options):
            if extract_code(opt) == str(code):
                return k
        return 0

    with st.expander(f"[{i}] {prod_name}  —  신뢰도 {conf}%"):
        sel = st.selectbox(
            "최종 카테고리 선택 (TOP3 + 기본값)",
            options=options,
            index=find_option_by_code(current_code),
            key=f"cat_sel_{i}"
        )
        chosen_code = extract_code(sel)
        st.session_state.category_overrides[i] = chosen_code

        # 빠른 입력(코드 직접)
        manual_code = st.text_input("또는 카테고리코드 직접 입력", value=chosen_code, key=f"cat_manual_{i}")
        st.session_state.category_overrides[i] = manual_code.strip() if manual_code.strip() else chosen_code

st.divider()

# 최종 카테고리 코드 리스트 만들기(전체 행 기준)
final_codes = []
for i in range(len(src_df)):
    if i in st.session_state.category_overrides and st.session_state.category_overrides[i]:
        final_codes.append(st.session_state.category_overrides[i])
    else:
        # 보정 없으면 추천1 코드 적용
        final_codes.append(str(suggest_df.loc[suggest_df["row_index"] == i, "추천1_코드"].values[0]))

# 변환 실행
st.subheader("3) 변환 파일 생성")
if st.button("🚀 스마트스토어 업로드용 엑셀 생성", type="primary", use_container_width=True):
    try:
        out_buf = convert_to_template(
            src_df=src_df,
            template_path=TEMPLATE_PATH,
            category_final_codes=final_codes,
            as_phone=st.session_state.as_phone,
            as_info=st.session_state.as_info
        )
        st.success("완료! 아래 버튼으로 다운로드하세요.")
        st.download_button(
            "📥 변환된 엑셀 다운로드",
            data=out_buf.getvalue(),
            file_name="스마트스토어_일괄등록_변환.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    except Exception as e:
        st.error(f"변환 중 오류: {e}")

st.divider()
st.markdown(
    """
### 적용 규칙(고정)
- **판매가** = 원본 `원가(옵션)`
- **즉시할인 값(기본할인)** = `원가(옵션) - 판매가(옵션합)` (단위: 원)
- **A/S 전화번호** = 02-6953-4020 (기본값)
- **A/S 안내** = 평일 오전 10시부터 오후 6시까지 (기본값)
- **카테고리코드** = 자동분류 추천1 (단, 사용자가 수동 보정하면 보정값 우선)
"""
)
