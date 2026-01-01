from io import BytesIO
import base64
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="스마트스토어 일괄상품등록 변환기", layout="wide")

# =========================
# (필수) 여기 두 개를 embed_assets.py 출력값으로 채우세요
# =========================
TEMPLATE_B64 = ""  # <-- 여기에 base64 붙여넣기
CATEGORY_B64 = ""  # <-- 여기에 base64 붙여넣기

# (선택) 폴백용 경로(리포에 파일이 있을 때만 사용)
TEMPLATE_PATH = "ExcelSaveTemplate_250311 (2).xlsx"
CATEGORY_MASTER_PATH = "category_20260102_012842.xls"

# 고정값
AS_PHONE_DEFAULT = "02-6953-4020"
AS_INFO_DEFAULT = "평일 오전 10시부터 오후 6시까지"
DEFAULT_CATEGORY_CODE = "50003307"

TOKEN_RE = re.compile(r"[가-힣a-zA-Z0-9]+")

# =========================
# Bytes 로더 (내장 우선 -> 파일 폴백)
# =========================
def load_bytes_from_b64_or_file(b64_str: str, file_path: str) -> bytes:
    if b64_str and b64_str.strip():
        return base64.b64decode(b64_str.encode("utf-8"))
    # 폴백: 파일
    with open(file_path, "rb") as f:
        return f.read()

def tokenize(text) -> set[str]:
    if text is None:
        return set()
    return set(TOKEN_RE.findall(str(text).lower()))

def clean_main_image(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip("[]")
    if "," in s:
        s = s.split(",")[0]
    return s.strip().strip("'").strip('"').strip()

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

def is_required(val) -> bool:
    return str(val).replace(" ", "") == "필수"

# =========================
# 카테고리 마스터 로드 (내장 지원)
# =========================
@st.cache_data
def load_category_master_from_bytes(cat_bytes: bytes) -> pd.DataFrame:
    cat = pd.read_excel(BytesIO(cat_bytes))

    # 사용자가 올린 파일 컬럼 구조(확정):
    # ['카테고리번호','대분류','중분류','소분류','세분류']
    required = ["카테고리번호", "대분류", "중분류", "소분류", "세분류"]
    for c in required:
        if c not in cat.columns:
            raise ValueError(f"카테고리 파일 컬럼 인식 실패. 현재 컬럼: {cat.columns.tolist()}")

    cat = cat[required].copy()
    cat.columns = ["category_code", "lv1", "lv2", "lv3", "lv4"]
    cat = cat.fillna("")

    cat["category_path"] = (cat["lv1"] + " " + cat["lv2"] + " " + cat["lv3"] + " " + cat["lv4"]).str.strip()
    cat["tokens"] = cat["category_path"].apply(tokenize)
    cat = cat[cat["tokens"].map(len) > 0].reset_index(drop=True)

    return cat

def classify_top3(text: str, cat_df: pd.DataFrame):
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

    candidates = [{
        "code": str(r["category_code"]),
        "path": str(r["category_path"]),
        "score": int(s)
    } for s, r in top3]

    confidence = round(candidates[0]["score"] / max(1, len(tokens)) * 100, 1)
    confidence = max(0.0, min(100.0, confidence))
    return candidates, confidence

# =========================
# 스마트스토어 변환 (템플릿 내장 지원)
# =========================
def convert_to_smartstore(src_df: pd.DataFrame, template_bytes: bytes, final_categories: list[str]) -> BytesIO:
    tpl = pd.read_excel(BytesIO(template_bytes), header=None)
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

        # 필수값 채우기
        for c in required_cols:
            out.iat[r, c] = example_row[c]

        # 판매자 상품코드
        seller = row.get("상품연동코드(옵션)")
        if pd.isna(seller) or str(seller).strip() == "":
            seller = row.get("상품ID")

        cost = row.get("원가(옵션)")
        sell_sum = row.get("판매가(옵션합)")

        out.iat[r, IDX["seller"]] = str(seller)

        if IDX["category"] is not None:
            out.iat[r, IDX["category"]] = final_categories[i] if i < len(final_categories) else DEFAULT_CATEGORY_CODE

        if IDX["name"] is not None:
            out.iat[r, IDX["name"]] = row.get("상품명", "")

        # 판매가 = 원가(옵션)
        if IDX["price"] is not None:
            out.iat[r, IDX["price"]] = cost if pd.notna(cost) else ""

        # 즉시할인 값 = 원가(옵션) - 판매가(옵션합)
        if IDX["discount"] is not None and pd.notna(cost) and pd.notna(sell_sum):
            out.iat[r, IDX["discount"]] = float(cost) - float(sell_sum)

        # 즉시할인 단위
        if IDX["discount_unit"] is not None:
            out.iat[r, IDX["discount_unit"]] = "원" if (pd.notna(cost) and pd.notna(sell_sum)) else ""

        # 재고/이미지
        if IDX["stock"] is not None:
            out.iat[r, IDX["stock"]] = row.get("재고수량(옵션)", "")
        if IDX["image"] is not None:
            out.iat[r, IDX["image"]] = clean_main_image(row.get("대표이미지"))

        # A/S 고정
        if IDX["as_phone"] is not None:
            out.iat[r, IDX["as_phone"]] = AS_PHONE_DEFAULT
        if IDX["as_info"] is not None:
            out.iat[r, IDX["as_info"]] = AS_INFO_DEFAULT

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        out.to_excel(w, index=False, header=False)
    buf.seek(0)
    return buf

# =========================
# UI
# =========================
st.title("스마트스토어 일괄상품등록 변환기 (TOP3 + 신뢰도 + 수동보정)")
st.caption("원본 상품내역 엑셀만 업로드하면, 템플릿/카테고리마스터는 내장값으로 자동 사용합니다.")

src_file = st.file_uploader("📂 원본 상품내역 엑셀 업로드 (.xlsx)", type=["xlsx"])
if not src_file:
    st.stop()

src_df = pd.read_excel(src_file)

# 내장/폴백 로드
try:
    template_bytes = load_bytes_from_b64_or_file(TEMPLATE_B64, TEMPLATE_PATH)
except Exception as e:
    st.error(f"템플릿 로드 실패: {e}")
    st.stop()

try:
    cat_bytes = load_bytes_from_b64_or_file(CATEGORY_B64, CATEGORY_MASTER_PATH)
    cat_df = load_category_master_from_bytes(cat_bytes)
except Exception as e:
    st.error(f"카테고리 마스터 로드 실패: {e}")
    st.stop()

# 분류 결과 계산
rows = []
for i, row in src_df.iterrows():
    text = f"{row.get('상품명','')} {row.get('뱃지/태그','')}".strip()
    top3, conf = classify_top3(text, cat_df)

    def fmt(c):
        return f"{c['code']} | {c['path']}"

    rows.append({
        "idx": i,
        "상품명": row.get("상품명", ""),
        "신뢰도(%)": conf,
        "TOP1": fmt(top3[0]) if len(top3) > 0 else f"{DEFAULT_CATEGORY_CODE} | (기본값)",
        "TOP2": fmt(top3[1]) if len(top3) > 1 else "",
        "TOP3": fmt(top3[2]) if len(top3) > 2 else "",
        "추천코드": top3[0]["code"] if len(top3) > 0 else DEFAULT_CATEGORY_CODE
    })

df_view = pd.DataFrame(rows)

st.subheader("카테고리 TOP3 + 신뢰도")
st.dataframe(df_view[["상품명", "신뢰도(%)", "TOP1", "TOP2", "TOP3"]], use_container_width=True, hide_index=True)

# 수동 보정
st.subheader("카테고리 수동 보정")
if "overrides" not in st.session_state:
    st.session_state.overrides = {}

for _, r in df_view.iterrows():
    i = int(r["idx"])
    with st.expander(f"[{i}] {r['상품명']} (신뢰도 {r['신뢰도(%)']}%)"):
        code = st.text_input(
            "최종 카테고리코드",
            value=st.session_state.overrides.get(i, r["추천코드"]),
            key=f"cat_{i}"
        )
        st.session_state.overrides[i] = code.strip() if code.strip() else r["추천코드"]

final_categories = [
    st.session_state.overrides.get(i, df_view.loc[i, "추천코드"])
    for i in range(len(src_df))
]

if st.button("🚀 스마트스토어 업로드 엑셀 생성", type="primary", use_container_width=True):
    out = convert_to_smartstore(src_df, template_bytes, final_categories)
    st.download_button(
        "📥 변환된 엑셀 다운로드",
        data=out.getvalue(),
        file_name="스마트스토어_일괄등록_최종.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
