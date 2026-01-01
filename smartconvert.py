from io import BytesIO
import re
import pandas as pd
import streamlit as st

# =========================
# 기본 설정
# =========================
st.set_page_config(
    page_title="스마트스토어 일괄상품등록 변환기",
    layout="wide"
)

TEMPLATE_PATH = "ExcelSaveTemplate_250311.xlsx"
CATEGORY_MASTER_PATH = "category_20260102_012842.xls"

AS_PHONE_DEFAULT = "02-6953-4020"
AS_INFO_DEFAULT = "평일 오전 10시부터 오후 6시까지"
DEFAULT_CATEGORY_CODE = "50003307"

TOKEN_RE = re.compile(r"[가-힣a-zA-Z0-9]+")

# =========================
# 공통 함수
# =========================
def tokenize(text):
    return set(TOKEN_RE.findall(str(text).lower())) if text else set()

def clean_main_image(x):
    if pd.isna(x):
        return ""
    s = str(x).strip("[]")
    if "," in s:
        s = s.split(",")[0]
    return s.strip().strip("'").strip('"')

def find_col_contains(labels, key):
    for i, v in enumerate(labels):
        if key in str(v):
            return i
    return None

def find_col_exact(labels, key):
    for i, v in enumerate(labels):
        if str(v).strip() == key:
            return i
    return None

def is_required(val):
    return str(val).replace(" ", "") == "필수"

# =========================
# 카테고리 마스터 로드 (수정본)
# =========================
@st.cache_data
def load_category_master():
    cat = pd.read_excel(CATEGORY_MASTER_PATH)

    required = ["카테고리번호", "대분류", "중분류", "소분류", "세분류"]
    for c in required:
        if c not in cat.columns:
            raise ValueError(f"카테고리 파일에 {c} 컬럼이 없습니다")

    cat = cat[required].copy()
    cat.columns = ["category_code", "lv1", "lv2", "lv3", "lv4"]
    cat.fillna("", inplace=True)

    cat["category_path"] = (
        cat["lv1"] + " " + cat["lv2"] + " " + cat["lv3"] + " " + cat["lv4"]
    ).str.strip()

    cat["tokens"] = cat["category_path"].apply(tokenize)
    cat = cat[cat["tokens"].map(len) > 0]

    return cat.reset_index(drop=True)

# =========================
# TOP3 분류 + 신뢰도
# =========================
def classify_top3(text, cat_df):
    tokens = tokenize(text)
    if not tokens:
        return [], 0.0

    scores = []
    for _, r in cat_df.iterrows():
        overlap = len(tokens & r["tokens"])
        if overlap > 0:
            scores.append((overlap, r))

    if not scores:
        return [], 0.0

    scores.sort(key=lambda x: x[0], reverse=True)
    top3 = scores[:3]

    result = []
    for s, r in top3:
        result.append({
            "code": str(r["category_code"]),
            "path": r["category_path"],
            "score": s
        })

    confidence = round(top3[0][0] / max(1, len(tokens)) * 100, 1)
    return result, confidence

# =========================
# 스마트스토어 변환
# =========================
def convert_to_smartstore(src_df, final_categories):
    tpl = pd.read_excel(TEMPLATE_PATH, header=None)
    labels = tpl.iloc[1].astype(str).tolist()

    IDX = {
        "seller": find_col_exact(labels, "판매자 상품코드") or 0,
        "category": find_col_exact(labels, "카테고리코드"),
        "name": find_col_exact(labels, "상품명"),
        "price": find_col_exact(labels, "판매가"),
        "stock": find_col_exact(labels, "재고수량"),
        "image": find_col_exact(labels, "대표이미지"),
        "discount": find_col_contains(labels, "즉시할인 값"),
        "discount_unit": find_col_contains(labels, "즉시할인 단위"),
        "as_phone": find_col_contains(labels, "A/S 전화번호"),
        "as_info": find_col_contains(labels, "A/S 안내"),
    }

    required_row = tpl.iloc[2]
    example_row = tpl.iloc[3]
    required_cols = [i for i in range(len(required_row)) if is_required(required_row[i])]

    start = tpl.shape[0] - 1
    out = tpl.copy().reindex(range(start + len(src_df)))

    for i, row in src_df.iterrows():
        r = start + i
        for c in required_cols:
            out.iat[r, c] = example_row[c]

        seller = row.get("상품연동코드(옵션)") or row.get("상품ID")
        cost = row.get("원가(옵션)")
        sell_sum = row.get("판매가(옵션합)")

        out.iat[r, IDX["seller"]] = seller
        out.iat[r, IDX["category"]] = final_categories[i]
        out.iat[r, IDX["name"]] = row.get("상품명")
        out.iat[r, IDX["price"]] = cost
        out.iat[r, IDX["stock"]] = row.get("재고수량(옵션)")
        out.iat[r, IDX["image"]] = clean_main_image(row.get("대표이미지"))

        if pd.notna(cost) and pd.notna(sell_sum):
            out.iat[r, IDX["discount"]] = cost - sell_sum
            out.iat[r, IDX["discount_unit"]] = "원"

        out.iat[r, IDX["as_phone"]] = AS_PHONE_DEFAULT
        out.iat[r, IDX["as_info"]] = AS_INFO_DEFAULT

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        out.to_excel(w, index=False, header=False)
    buf.seek(0)
    return buf

# =========================
# UI
# =========================
st.title("스마트스토어 일괄상품등록 변환기 (통합본)")
st.caption("카테고리 TOP3 · 신뢰도 · 수동보정 · 스마트스토어 양식 자동 변환")

src_file = st.file_uploader("📂 원본 상품내역 엑셀 업로드 (.xlsx)", type=["xlsx"])
if not src_file:
    st.stop()

src_df = pd.read_excel(src_file)
cat_df = load_category_master()

if "overrides" not in st.session_state:
    st.session_state.overrides = {}

rows = []
for i, row in src_df.iterrows():
    text = f"{row.get('상품명','')} {row.get('뱃지/태그','')}"
    top3, conf = classify_top3(text, cat_df)

    rows.append({
        "idx": i,
        "상품명": row.get("상품명"),
        "신뢰도(%)": conf,
        "TOP1": f"{top3[0]['code']} | {top3[0]['path']}" if top3 else "",
        "TOP2": f"{top3[1]['code']} | {top3[1]['path']}" if len(top3) > 1 else "",
        "TOP3": f"{top3[2]['code']} | {top3[2]['path']}" if len(top3) > 2 else "",
        "추천코드": top3[0]["code"] if top3 else DEFAULT_CATEGORY_CODE
    })

df_view = pd.DataFrame(rows)
st.dataframe(df_view[["상품명", "신뢰도(%)", "TOP1", "TOP2", "TOP3"]], use_container_width=True)

st.subheader("카테고리 수동 보정")
for _, r in df_view.iterrows():
    i = r["idx"]
    with st.expander(f"[{i}] {r['상품명']} (신뢰도 {r['신뢰도(%)']}%)"):
        code = st.text_input(
            "최종 카테고리코드",
            value=st.session_state.overrides.get(i, r["추천코드"]),
            key=f"cat_{i}"
        )
        st.session_state.overrides[i] = code

final_categories = [
    st.session_state.overrides.get(i, df_view.loc[i, "추천코드"])
    for i in range(len(src_df))
]

if st.button("🚀 스마트스토어 업로드 엑셀 생성", use_container_width=True):
    out = convert_to_smartstore(src_df, final_categories)
    st.download_button(
        "📥 다운로드",
        data=out.getvalue(),
        file_name="스마트스토어_일괄등록_최종.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
