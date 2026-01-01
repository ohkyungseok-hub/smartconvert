from io import BytesIO
import base64
import re
import pandas as pd
import streamlit as st

# =========================
# 기본 설정
# =========================
st.set_page_config(page_title="스마트스토어 일괄상품등록 변환기", layout="wide")

# =========================
# 내장(base64) – 비워두면 파일 폴백
# =========================
TEMPLATE_B64 = ""
CATEGORY_B64 = ""

TEMPLATE_PATH = "ExcelSaveTemplate_250311 (2).xlsx"
CATEGORY_MASTER_PATH = "category_20260102_012842.xls"

# =========================
# 고정값
# =========================
AS_PHONE_DEFAULT = "02-6953-4020"
AS_INFO_DEFAULT = "평일 오전 10시부터 오후 6시까지"
DELIVERY_TEMPLATE_CODE = "3308887"
DEFAULT_CATEGORY_CODE = "50003307"

TOKEN_RE = re.compile(r"[가-힣a-zA-Z0-9]+")

# =========================
# 공통 유틸
# =========================
def load_bytes(b64_str, path):
    if b64_str:
        return base64.b64decode(b64_str.encode())
    with open(path, "rb") as f:
        return f.read()

def tokenize(text):
    return set(TOKEN_RE.findall(str(text).lower())) if text else set()

def clean_image(x):
    if pd.isna(x):
        return ""
    s = str(x).strip("[]")
    if "," in s:
        s = s.split(",")[0]
    return s.strip().strip("'").strip('"')

def find_col(labels, key, exact=False):
    for i, v in enumerate(labels):
        if exact and str(v).strip() == key:
            return i
        if not exact and key in str(v):
            return i
    return None

def is_required(val):
    return str(val).replace(" ", "") == "필수"

# =========================
# 카테고리 마스터 로드
# =========================
@st.cache_data
def load_category_master(cat_bytes):
    cat = pd.read_excel(BytesIO(cat_bytes))
    cat = cat[["카테고리번호", "대분류", "중분류", "소분류", "세분류"]]
    cat.columns = ["code", "lv1", "lv2", "lv3", "lv4"]
    cat.fillna("", inplace=True)

    cat["path"] = (cat["lv1"] + " " + cat["lv2"] + " " + cat["lv3"] + " " + cat["lv4"]).str.strip()
    cat["tokens"] = cat["path"].apply(tokenize)
    cat["label"] = cat["code"].astype(str) + " | " + cat["path"]
    return cat[cat["tokens"].map(len) > 0].reset_index(drop=True)

# =========================
# 카테고리 분류
# =========================
def classify_top3(text, cat_df):
    tokens = tokenize(text)
    scores = []
    for _, r in cat_df.iterrows():
        overlap = len(tokens & r["tokens"])
        if overlap:
            scores.append((overlap, r))
    if not scores:
        return [], 0.0
    scores.sort(key=lambda x: x[0], reverse=True)
    top3 = scores[:3]
    conf = round(top3[0][0] / max(1, len(tokens)) * 100, 1)
    return top3, conf

def autocomplete(cat_df, query, limit=50):
    q = (query or "").lower().strip()
    if not q:
        return cat_df["label"].head(limit).tolist()
    q_tokens = tokenize(q)
    scored = []
    for _, r in cat_df.iterrows():
        overlap = len(q_tokens & r["tokens"])
        if overlap or q in r["path"].lower():
            scored.append((overlap * 10, r["label"]))
    scored.sort(reverse=True)
    return [x[1] for x in scored[:limit]]

def apply_same_name(src_df, overrides, idx, code):
    name = src_df.loc[idx, "상품명"]
    targets = src_df.index[src_df["상품명"] == name]
    for i in targets:
        overrides[i] = code
    return len(targets)

# =========================
# 스마트스토어 변환
# =========================
def convert(src_df, tpl_bytes, final_categories):
    tpl = pd.read_excel(BytesIO(tpl_bytes), header=None)
    labels = tpl.iloc[1].astype(str).tolist()

    IDX = {
        "seller": find_col(labels, "판매자 상품코드", True) or 0,
        "category": find_col(labels, "카테고리코드", True),
        "name": find_col(labels, "상품명", True),
        "price": find_col(labels, "판매가", True),
        "stock": find_col(labels, "재고수량", True),
        "image": find_col(labels, "대표이미지", True),
        "discount": find_col(labels, "즉시할인 값"),
        "discount_unit": find_col(labels, "즉시할인 단위"),
        "as_phone": find_col(labels, "A/S 전화번호"),
        "as_info": find_col(labels, "A/S 안내"),
        "delivery": find_col(labels, "배송", False),
    }

    required_cols = [i for i, v in enumerate(tpl.iloc[2]) if is_required(v)]
    example = tpl.iloc[3]
    start = tpl.shape[0] - 1
    out = tpl.copy().reindex(range(start + len(src_df)))

    for i, row in src_df.iterrows():
        r = start + i
        for c in required_cols:
            out.iat[r, c] = example[c]

        seller = row.get("상품연동코드(옵션)") or row.get("상품ID")
        cost = row.get("원가(옵션)")
        sell_sum = row.get("판매가(옵션합)")

        out.iat[r, IDX["seller"]] = seller
        out.iat[r, IDX["category"]] = final_categories[i]
        out.iat[r, IDX["name"]] = row.get("상품명")
        out.iat[r, IDX["price"]] = cost
        out.iat[r, IDX["stock"]] = row.get("재고수량(옵션)")
        out.iat[r, IDX["image"]] = clean_image(row.get("대표이미지"))

        if pd.notna(cost) and pd.notna(sell_sum):
            out.iat[r, IDX["discount"]] = cost - sell_sum
            out.iat[r, IDX["discount_unit"]] = "원"

        out.iat[r, IDX["as_phone"]] = AS_PHONE_DEFAULT
        out.iat[r, IDX["as_info"]] = AS_INFO_DEFAULT
        out.iat[r, IDX["delivery"]] = DELIVERY_TEMPLATE_CODE

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        out.to_excel(w, index=False, header=False)
    buf.seek(0)
    return buf

# =========================
# UI
# =========================
st.title("스마트스토어 일괄상품등록 변환기 (최종 통합본)")

src_file = st.file_uploader("📂 원본 상품내역 엑셀 업로드 (.xlsx)", type=["xlsx"])
if not src_file:
    st.stop()

src_df = pd.read_excel(src_file)
tpl_bytes = load_bytes(TEMPLATE_B64, TEMPLATE_PATH)
cat_bytes = load_bytes(CATEGORY_B64, CATEGORY_MASTER_PATH)
cat_df = load_category_master(cat_bytes)

rows = []
for i, row in src_df.iterrows():
    text = f"{row.get('상품명','')} {row.get('뱃지/태그','')}"
    top3, conf = classify_top3(text, cat_df)
    rows.append({
        "idx": i,
        "상품명": row.get("상품명"),
        "신뢰도": conf,
        "TOP": [r["label"] for _, r in top3],
        "추천코드": top3[0][1]["code"] if top3 else DEFAULT_CATEGORY_CODE
    })

df_view = pd.DataFrame(rows)
st.dataframe(df_view[["상품명", "신뢰도"]], use_container_width=True)

if "overrides" not in st.session_state:
    st.session_state.overrides = {}

st.subheader("카테고리 수동 보정 (자동완성 + 동일 상품명 일괄 적용)")
for _, r in df_view.iterrows():
    i = r["idx"]
    with st.expander(f"[{i}] {r['상품명']} (신뢰도 {r['신뢰도']}%)"):
        search = st.text_input("카테고리 검색", key=f"s{i}")
        options = autocomplete(cat_df, search)
        sel = st.selectbox("카테고리 선택", options, key=f"c{i}")
        code = sel.split("|")[0]
        st.session_state.overrides[i] = code

        if st.button("동일 상품명 전체 적용", key=f"a{i}"):
            n = apply_same_name(src_df, st.session_state.overrides, i, code)
            st.success(f"{n}건 적용 완료")

final_categories = [
    st.session_state.overrides.get(i, df_view.loc[df_view.idx == i, "추천코드"].values[0])
    for i in range(len(src_df))
]

if st.button("🚀 스마트스토어 업로드 엑셀 생성", type="primary"):
    out = convert(src_df, tpl_bytes, final_categories)
    st.download_button(
        "📥 다운로드",
        out.getvalue(),
        "스마트스토어_일괄등록_최종.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
