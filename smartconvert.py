# streamlit_app.py
# ✅ 통합본: 템플릿/카테고리 내장(또는 파일 폴백) + TOP3 + 신뢰도 + 수동보정(자동완성) + 동일 상품명 일괄적용
# - 판매가 = 원가(옵션)
# - 즉시할인 값 = 원가(옵션) - 판매가(옵션합)  (단위: 원)
# - A/S 전화번호/안내 고정
#
# ⚠️ TEMPLATE_B64 / CATEGORY_B64에 base64를 넣으면 배포 환경에서도 파일 없이 동작합니다.
#    (비워두면 TEMPLATE_PATH/CATEGORY_MASTER_PATH 파일을 읽는 폴백 모드)

from io import BytesIO
import base64
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="스마트스토어 일괄상품등록 변환기", layout="wide")

# =========================
# 1) 내장(권장): embed_assets.py로 생성한 base64를 붙여 넣으세요
# =========================
TEMPLATE_B64 = ""  # ExcelSaveTemplate_250311 (2).xlsx base64
CATEGORY_B64 = ""  # category_20260102_012842.xls base64

# =========================
# 2) 폴백(선택): repo/폴더에 파일이 존재하면 읽습니다
# =========================
TEMPLATE_PATH = "ExcelSaveTemplate_250311 (2).xlsx"
CATEGORY_MASTER_PATH = "category_20260102_012842.xls"

# =========================
# 고정값
# =========================
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
    with open(file_path, "rb") as f:
        return f.read()


# =========================
# 공통 유틸
# =========================
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
# 카테고리 마스터 로드 (업로드 파일 구조에 맞춤)
# columns: ['카테고리번호','대분류','중분류','소분류','세분류']
# =========================
@st.cache_data
def load_category_master_from_bytes(cat_bytes: bytes) -> pd.DataFrame:
    cat = pd.read_excel(BytesIO(cat_bytes))

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

    # UI 표시용 라벨
    cat["label"] = cat["category_code"].astype(str) + " | " + cat["category_path"].astype(str)

    return cat


# =========================
# TOP3 분류 + 신뢰도
# =========================
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
        "score": int(s),
        "label": str(r["label"]),
    } for s, r in top3]

    confidence = round(candidates[0]["score"] / max(1, len(tokens)) * 100, 1)
    confidence = max(0.0, min(100.0, confidence))
    return candidates, confidence


# =========================
# 자동완성 추천(입력 즉시 추천)
# =========================
def autocomplete_categories(cat_df: pd.DataFrame, query: str, limit: int = 50):
    q = (query or "").strip().lower()
    if not q:
        return cat_df["label"].head(limit).tolist()

    q_tokens = tokenize(q)
    if not q_tokens:
        sub = cat_df[cat_df["category_path"].astype(str).str.lower().str.contains(q, na=False)].head(limit)
        return sub["label"].tolist()

    scores = []
    for idx, row in cat_df.iterrows():
        path = str(row["category_path"]).lower()
        overlap = len(q_tokens & row["tokens"])
        if overlap == 0 and q not in path:
            continue
        bonus = 2 if q in path else 0
        score = overlap * 10 + bonus
        scores.append((score, idx))

    if not scores:
        return []

    scores.sort(reverse=True)
    top_idx = [i for _, i in scores[:limit]]
    return cat_df.loc[top_idx, "label"].tolist()


# =========================
# 동일 상품명 일괄 적용
# =========================
def apply_category_to_same_name(src_df: pd.DataFrame, overrides: dict, target_index: int, category_code: str):
    target_name = str(src_df.loc[target_index, "상품명"])
    same_idx = src_df.index[src_df["상품명"].astype(str) == target_name].tolist()
    for j in same_idx:
        overrides[j] = category_code
    return len(same_idx)


# =========================
# 스마트스토어 변환 (템플릿 내장/폴백 지원)
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

        # 템플릿 필수값 채움(예시값)
        for c in required_cols:
            out.iat[r, c] = example_row[c]

        # 판매자 상품코드: 옵션연동코드 있으면 사용, 없으면 상품ID
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

        # 즉시할인 단위 = 원 (값 있을 때만)
        if IDX["discount_unit"] is not None:
            out.iat[r, IDX["discount_unit"]] = "원" if (pd.notna(cost) and pd.notna(sell_sum)) else ""

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
st.title("스마트스토어 일괄상품등록 변환기 (통합본)")
st.caption("카테고리 TOP3 · 신뢰도 · 자동완성 수동보정 · 동일 상품명 일괄 적용 · 스마트스토어 업로드 엑셀 생성")

with st.expander("내장/폴백 설정", expanded=False):
    st.write("✅ base64를 넣으면 배포 환경에서도 파일 없이 동작합니다.")
    st.code(f"TEMPLATE_PATH = {TEMPLATE_PATH}\nCATEGORY_MASTER_PATH = {CATEGORY_MASTER_PATH}", language="text")
    st.write("TEMPLATE_B64 / CATEGORY_B64가 비어있으면 위 파일 경로에서 읽습니다.")

src_file = st.file_uploader("📂 원본 상품내역 엑셀 업로드 (.xlsx)", type=["xlsx"])
if not src_file:
    st.stop()

src_df = pd.read_excel(src_file)

# 템플릿/카테고리 바이트 로드
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

# 원본 필수 컬럼 간단 체크
needed_cols = ["상품ID", "상품명", "원가(옵션)", "판매가(옵션합)", "재고수량(옵션)", "대표이미지"]
missing = [c for c in needed_cols if c not in src_df.columns]
if missing:
    st.warning(f"원본 파일에 아래 컬럼이 없습니다(일부 값은 공란 처리될 수 있어요): {', '.join(missing)}")

st.divider()
st.subheader("1) 카테고리 자동 분류 결과 (TOP3 + 신뢰도)")

# 추천 계산(캐시)
@st.cache_data
def compute_suggestions(df: pd.DataFrame, cat_df: pd.DataFrame):
    rows = []
    for i, row in df.iterrows():
        text = f"{row.get('상품명','')} {row.get('뱃지/태그','')}".strip()
        top3, conf = classify_top3(text, cat_df)

        rows.append({
            "idx": i,
            "상품명": row.get("상품명", ""),
            "신뢰도(%)": conf,
            "TOP1": top3[0]["label"] if len(top3) > 0 else f"{DEFAULT_CATEGORY_CODE} | (기본값)",
            "TOP2": top3[1]["label"] if len(top3) > 1 else "",
            "TOP3": top3[2]["label"] if len(top3) > 2 else "",
            "추천코드": top3[0]["code"] if len(top3) > 0 else DEFAULT_CATEGORY_CODE
        })
    return pd.DataFrame(rows)

df_view = compute_suggestions(src_df, cat_df)

# 필터/페이징
c1, c2, c3, c4 = st.columns([1.2, 1.2, 1.2, 2.4])
with c1:
    min_conf = st.slider("표시 최소 신뢰도", 0, 100, 0)
with c2:
    only_low = st.checkbox("저신뢰도만 보기", value=False)
with c3:
    page_size = st.selectbox("페이지 크기", [10, 20, 50, 100], index=1)
with c4:
    keyword = st.text_input("상품명 검색", value="")

view = df_view.copy()
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

max_page = (total - 1) // page_size + 1
page = st.number_input("페이지", min_value=1, max_value=max_page, value=1, step=1)
start = (page - 1) * page_size
end = min(total, start + page_size)
page_df = view.iloc[start:end].copy()

st.dataframe(page_df[["상품명", "신뢰도(%)", "TOP1", "TOP2", "TOP3"]], use_container_width=True, hide_index=True)

st.divider()
st.subheader("2) 카테고리 수동 보정 (자동완성 + 동일 상품명 일괄 적용)")

if "overrides" not in st.session_state:
    st.session_state.overrides = {}

# 저신뢰도만 보정(선택)
fix_threshold = st.slider("수동보정 대상(신뢰도 이하만 표시)", 0, 100, 30)
df_fix = df_view[df_view["신뢰도(%)"] <= fix_threshold].copy()
st.caption(f"현재 수동보정 표시 대상: 신뢰도 ≤ {fix_threshold}% (총 {len(df_fix)}건)")

for _, r in df_fix.iterrows():
    i = int(r["idx"])
    current_code = st.session_state.overrides.get(i, r["추천코드"])

    top_opts = [x for x in [r["TOP1"], r["TOP2"], r["TOP3"]] if x]
    top_codes = [opt.split("|")[0].strip() for opt in top_opts] if top_opts else []

    with st.expander(f"[{i}] {r['상품명']} (신뢰도 {r['신뢰도(%)']}%)"):
        # (A) TOP3 빠른 선택
        if top_opts:
            default_idx = top_codes.index(current_code) if current_code in top_codes else 0
            sel_top = st.radio(
                "TOP3에서 빠르게 선택",
                options=list(range(len(top_opts))),
                format_func=lambda k: top_opts[k],
                index=default_idx,
                key=f"top_radio_{i}"
            )
            current_code = top_opts[sel_top].split("|")[0].strip()
            st.session_state.overrides[i] = current_code

        st.divider()

        # (B) 자동완성 검색 (입력 즉시 추천)
        st.write("카테고리 자동완성 검색 (입력 즉시 추천)")
        search_q = st.text_input(
            "키워드 입력 (예: 음료, 탄산, 과자, 디저트, 생활용품...)",
            value=st.session_state.get(f"search_{i}", ""),
            key=f"search_{i}"
        )

        options = autocomplete_categories(cat_df, search_q, limit=50)

        # 현재 코드가 options에 없으면 맨 앞에 노출
        match = cat_df[cat_df["category_code"].astype(str) == str(current_code)]
        current_label = match.iloc[0]["label"] if len(match) else None
        if current_label and current_label not in options:
            options = [current_label] + options

        if not options:
            st.info("추천 결과가 없습니다. 다른 키워드를 입력해보세요.")
        else:
            sel_label = st.selectbox(
                "추천 목록에서 선택",
                options=options,
                index=0,
                key=f"auto_sel_{i}"
            )
            chosen_code = sel_label.split("|")[0].strip()
            st.session_state.overrides[i] = chosen_code
            current_code = chosen_code

        # (C) 동일 상품명 일괄 적용
        b1, b2 = st.columns([1.2, 3.8])
        with b1:
            apply_all = st.button("동일 상품명 전체 적용", key=f"apply_same_{i}")
        with b2:
            st.caption("이 상품과 같은 '상품명'을 가진 모든 행에 현재 카테고리를 한 번에 적용합니다.")

        if apply_all:
            n = apply_category_to_same_name(src_df, st.session_state.overrides, i, current_code)
            st.success(f"✅ 동일 상품명 {n}건에 카테고리({current_code})를 일괄 적용했습니다.")

        st.success(f"최종 선택 카테고리코드: {current_code}")

st.divider()
st.subheader("3) 스마트스토어 업로드 엑셀 생성")

# 최종 카테고리 리스트
final_categories = []
for i in range(len(src_df)):
    final_categories.append(st.session_state.overrides.get(i, df_view.loc[df_view["idx"] == i, "추천코드"].values[0]))

if st.button("🚀 스마트스토어 업로드 엑셀 생성", type="primary", use_container_width=True):
    out = convert_to_smartstore(src_df, template_bytes, final_categories)
    st.download_button(
        "📥 변환된 엑셀 다운로드",
        data=out.getvalue(),
        file_name="스마트스토어_일괄등록_최종.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

st.divider()
st.markdown(
    """
### 적용 규칙(고정)
- **판매가** = 원본 `원가(옵션)`
- **즉시할인 값(기본할인)** = `원가(옵션) - 판매가(옵션합)` (단위: 원)
- **A/S 전화번호** = 02-6953-4020
- **A/S 안내** = 평일 오전 10시부터 오후 6시까지
- **카테고리코드** = 자동추천 TOP1 (단, 수동 보정/일괄 적용 값이 있으면 우선)
"""
)
