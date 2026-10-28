"""
app.py — Clinical Data QC 자동 실행 대시보드 (Streamlit)

흐름:
  1) rawdata / codebook 업로드 (settings.xlsx는 선택)
  2) 설정: settings.xlsx를 올리면 그 값으로, 없으면 자동 감지값으로 드롭다운을 채움
     → 확인하고 필요하면 수정
  3) [전체 검증 실행] → qc_core.run_all_validations
  4) 대시보드: 검증별 오류 건수 + 오류 목록 + 결과 다운로드

검증 로직은 전부 qc_core.py 사용 → CLI와 동일한 결과.
"""

import warnings
warnings.filterwarnings("ignore", message="Conditional Formatting extension")
warnings.filterwarnings("ignore", message="Could not infer format")

import re
import hmac
import hashlib
import pandas as pd
import streamlit as st

import qc_core as qc
from streamlit_sortables import sort_items

st.set_page_config(page_title="Clinical Data QC", layout="wide")

def check_password():
    """공용 비밀번호 확인."""

    # 이미 비밀번호 인증을 완료한 경우
    if st.session_state.get("password_correct", False):
        return True

    st.title("Data Management 검증")
    st.caption("서비스를 이용하려면 비밀번호를 입력해 주세요.")

    password = st.text_input(
        "비밀번호",
        type="password",
        placeholder="비밀번호를 입력하세요",
        key="login_password"
    )

    if st.button("로그인", type="primary"):
        if hmac.compare_digest(
            password,
            st.secrets["APP_PASSWORD"]
        ):
            st.session_state["password_correct"] = True
            st.session_state.pop("login_password", None)
            st.rerun()
        else:
            st.error("비밀번호가 올바르지 않습니다.")

    return False


# 비밀번호가 맞지 않으면 아래 검증 화면을 실행하지 않음
if not check_password():
    st.stop()


# 로그인 이후 사이드바
with st.sidebar:
    st.success("로그인되었습니다.")

    if st.button("로그아웃"):
        st.session_state["password_correct"] = False
        st.rerun()



# settings.xlsx 구분 별칭 (cli.py와 동일하게 유지)
SETTINGS_ALIASES = {
    "연구명": "study",
    "subject_id": "subject_col", "대상자컬럼": "subject_col",
    "visit": "visit_col", "방문컬럼": "visit_col",
    "primary_sheet": "base_sheets", "기준시트": "base_sheets", "기준시트들": "base_sheets",
    "비교제외": "exclude_cols", "비교제외변수": "exclude_cols",
    "변수명컬럼": "variable_col", "변수명": "variable_col",
    "코드리스트컬럼": "codelist_col", "코드리스트": "codelist_col",
    "데이터유형컬럼": "data_type_col", "데이터유형": "data_type_col",
    "ecrf명": "form_col", "ecrf명컬럼": "form_col",
    "항목명": "res_col", "항목명컬럼": "res_col",
    "검사시트": "target_sheets",
    "방문일컬럼": "visit_date_col",
    "변수명대조제외": "skip_var_names", "변수검증제외": "skip_var_names",
    "집계임계값": "aggregate_threshold",
    "패턴임계값": "pattern_unique_threshold", "unique임계값": "pattern_unique_threshold",
    "같은날허용방문": "same_day_visits", "동일날짜방문": "same_day_visits",
}


def load_settings_xlsx(uploaded):
    """settings.xlsx(구분|값 2열)를 dict로. 없으면 빈 dict."""
    if uploaded is None:
        return {}
    df = pd.read_excel(uploaded).iloc[:, :2]
    df.columns = ["key", "value"]
    out = {}
    for _, r in df.iterrows():
        key = qc.normalize_value(r["key"]).strip().lower().replace(" ", "_")
        val = qc.normalize_value(r["value"])
        if key == "" or val == "":
            continue
        out[SETTINGS_ALIASES.get(key, key)] = val
    return out


def settings_list(value):
    if not value:
        return []
    return [x.strip() for x in re.split(r"[,|\n]+", str(value)) if x.strip()]


# ---------- 자동 감지 헬퍼 ----------
def all_columns(raw_sheets):
    seen = []
    for df in raw_sheets.values():
        for c in df.columns:
            cn = str(c).strip()
            if cn and cn not in seen:
                seen.append(cn)
    return seen


def common_columns(raw_sheets):
    common = None
    for df in raw_sheets.values():
        cols = {str(c).strip() for c in df.columns}
        common = cols if common is None else (common & cols)
    return sorted(common or set())


def guess_subject(cols):
    for c in cols:
        if any(k in str(c).upper() for k in ("SUBJ", "SUBJECT", "대상자", "SCR")):
            return c
    return cols[0] if cols else None


def guess_kw(cols, keywords):
    def _norm(s):
        return re.sub(r"\s+", "", str(s).upper())
    keys = [_norm(k) for k in keywords]
    for c in cols:
        col_norm = _norm(c)
        if any(k in col_norm for k in keys):
            return c
    return None

def pick(label, options, want, key, help=None):
    """want(설정값/자동감지값)가 옵션에 있으면 그 위치를 기본 선택."""
    want = qc.normalize_name(want) if want else None
    norm_opts = {qc.normalize_name(o): o for o in options}
    default = norm_opts[want] if want and want in norm_opts else options[0]
    # 기본값은 session_state로만 지정 (index와 동시 지정 시 Streamlit 경고 발생)
    if st.session_state.get(key) not in options:
        st.session_state[key] = default
    return st.selectbox(label, options, key=key, help=help)


def ms_state(key, options, default):
    """multiselect 기본값을 session_state로 지정. 옵션에 없는 값은 제거."""
    if key not in st.session_state:
        st.session_state[key] = [v for v in default if v in options]
    else:
        st.session_state[key] = [v for v in st.session_state[key] if v in options]


def num_state(key, default, lo, hi):
    """number_input 기본값을 session_state로 지정 (범위 밖이면 보정)."""
    try:
        v = int(float(st.session_state.get(key, default)))
    except (TypeError, ValueError):
        v = int(default)
    st.session_state[key] = min(max(v, lo), hi)

# ---------- settings.xlsx 다운로드 ----------
def build_settings_bytes():
    """현재 화면 설정을 cli.py가 읽는 (구분|값) 형식 엑셀 bytes로. 파일 경로는 넣지 않음."""
    rows = [
        ("subject_col", subject_col),
        ("visit_col", visit_col),
        ("visit_date_col", visit_date_col),
        ("base_sheets", ", ".join(base_sheets)),
        ("variable_col", variable_col),
        ("codelist_col", codelist_col),
        ("form_col", form_col),
        ("res_col", res_col),
        ("data_type_col", "" if data_type_col == "(없음)" else data_type_col),
        ("exclude_cols", ", ".join(exclude_cols)),
        ("skip_var_names", ", ".join(skip_var_names)),
        ("same_day_visits", ", ".join(same_day_visits)),
        ("aggregate_threshold", str(int(agg_threshold))),
        ("pattern_unique_threshold", str(int(pattern_threshold))),
    ]
    df = pd.DataFrame(rows, columns=["구분", "값"])
    return qc.to_excel_bytes({"settings": df})


# ================== 화면 ==================
st.title("Data Management 검증")
st.caption("rawdata와 codebook을 업로드하면 설정을 확인한 뒤 전체 검증을 한 번에 실행합니다.  \n settings.xlsx와 classification_reviewed.xlsx, date_rules.xlsx는 아래 개별 검증에서 생성·다운로드한 후 업로드하여 사용할 수 있으며, 최초 1회만 설정하면 이후에도 계속 사용할 수 있습니다.  \n 각 항목에서 오른쪽 help 를 통해 설명 및 예시를 확인할 수 있습니다")
st.header("1. 파일 업로드")
st.markdown("**필수 파일**")
r1c1, r1c2 = st.columns(2)
with r1c1:
    raw_file = st.file_uploader("rawdata.xlsx", type=["xlsx", "xls"])

with r1c2:
    codebook_file = st.file_uploader("codebook.xlsx", type=["xlsx", "xls"])

st.markdown("**선택 파일**")
st.caption("이전에 이 화면에서 해당 연구로 저장한 파일이 있으면 업로드하고 처음이면 비워두세요.")
r2c1, r2c2, r2c3 = st.columns(3)
with r2c1:
    settings_file = st.file_uploader("settings.xlsx", type=["xlsx", "xls"])
    st.caption("검증에 필요한 설정 파일-단계 2에서 생성됩니다")
with r2c2:
    classification_file = st.file_uploader("classification_reviewed.xlsx", type=["xlsx", "xls"])
    st.caption("코드리스트 분류 파일-단계 3에서 생성됩니다.")
with r2c3:
    date_rule_file = st.file_uploader("date_rules.xlsx", type=["xlsx", "xls"])
    st.caption("날짜 규칙 파일 - 단계 4에서 생성됩니다.")

if raw_file is None or codebook_file is None:
    st.info("rawdata와 codebook 파일을 업로드하면 설정 단계가 나타납니다.")
    st.stop()
cfg = load_settings_xlsx(settings_file)
raw_sheets = qc.read_excel_all(raw_file)
codebook = qc.read_first_sheet(codebook_file)
sheet_names = list(raw_sheets.keys())
raw_all = all_columns(raw_sheets)
raw_common = raw_all
cb_cols = [str(c).strip() for c in codebook.columns]

def match_option(value, options, fallback=None):
    """settings 값과 실제 옵션을 공백·대소문자 차이를 무시하고 매칭."""
    if value:
        normalized_options = {
            qc.normalize_name(option): option
            for option in options
        }

        normalized_value = qc.normalize_name(value)

        if normalized_value in normalized_options:
            return normalized_options[normalized_value]

    return fallback


# settings.xlsx의 실제 내용으로 파일 변경 여부 확인
if settings_file is not None:
    settings_signature = hashlib.sha256(
        settings_file.getvalue()
    ).hexdigest()
else:
    settings_signature = None


# settings 파일이 새로 업로드되거나 변경됐을 때만 적용
if st.session_state.get("settings_signature") != settings_signature:
    st.session_state["settings_signature"] = settings_signature

    if settings_file is not None:
        # 단일 선택값
        st.session_state["cfg_subject"] = match_option(
            cfg.get("subject_col"),
            raw_common,
            guess_subject(raw_common)
        )

        st.session_state["cfg_visit"] = match_option(
            cfg.get("visit_col"),
            raw_common,
            guess_kw(raw_common, ("VISITNM", "VISIT_NM", "방문명"))
        )

        st.session_state["cfg_visitdt"] = match_option(
            cfg.get("visit_date_col"),
            raw_common,
            guess_kw(raw_common, ("VISITDT", "VISIT_DT", "방문일", "방문일자"))
        )

        st.session_state["cfg_var"] = match_option(
            cfg.get("variable_col"),
            cb_cols,
            guess_kw(cb_cols, ("변수명", "ITEM NAME", "VARIABLE"))
        )

        st.session_state["cfg_form"] = match_option(
            cfg.get("form_col"),
            cb_cols,
            guess_kw(cb_cols, ("ECRF", "FORM", "시트"))
        )

        st.session_state["cfg_codelist"] = match_option(
            cfg.get("codelist_col"),
            cb_cols,
            guess_kw(cb_cols, ("코드리스트", "CODE"))
        )

        st.session_state["cfg_res"] = match_option(
            cfg.get("res_col"),
            cb_cols,
            guess_kw(cb_cols, ("항목명", "Label"))
        )

        data_type_options = ["(없음)"] + cb_cols
        st.session_state["cfg_datatype"] = match_option(
            cfg.get("data_type_col"),
            data_type_options,
            guess_kw(cb_cols, ("Type",)) or "(없음)"
        )

        # 복수 선택값
        sheet_map = {
            qc.normalize_name(value): value
            for value in sheet_names
        }

        st.session_state["cfg_base_sheets"] = [
            sheet_map[qc.normalize_name(value)]
            for value in settings_list(cfg.get("base_sheets"))
            if qc.normalize_name(value) in sheet_map
        ]

        st.session_state["cfg_skip_var_names"] = [
            value
            for value in settings_list(cfg.get("skip_var_names"))
            if value in raw_all
        ]

        st.session_state["cfg_exclude_cols"] = [
            value
            for value in settings_list(cfg.get("exclude_cols"))
            if value in raw_all
        ]

        # 같은 날 방문 허용: 방문명 목록이 아래에서 정해지므로 초기화 후 settings 값으로 다시 채움
        st.session_state.pop("cfg_same_day_visits", None)

        # 숫자값
        st.session_state["cfg_agg_threshold"] = int(
            float(cfg.get("aggregate_threshold", 10))
        )

        st.session_state["cfg_pattern_threshold"] = int(
            float(cfg.get("pattern_unique_threshold", 10))
        )

msg = f"rawdata 시트 {len(sheet_names)}개, codebook 컬럼 {len(cb_cols)}개"
if cfg:
    msg += f", settings에서 {len(cfg)}개 항목 로드"
st.success(msg)

# ---------- 2단계: 설정 (settings값 > 자동감지값 순으로 기본 채움) ----------
st.header("2. 설정")
st.caption("settings.xlsx 값이 있으면 그 값을, 없으면 자동 감지값을 기본으로 채웁니다. 채워진 값을 확인 후 필요 시 수정해주세요.")

with st.expander("rawdata 컬럼 매핑", expanded=True):
    cc1, cc2, cc3 = st.columns(3)
    with cc1:
        subject_col = pick("Subject ID 컬럼", raw_common,
                           cfg.get("subject_col") or guess_subject(raw_common), "cfg_subject",
        help="ex. SUBJNO, SCR_NUM")
    with cc2:
        _visit_opts = ["(없음)"] + raw_common
        visit_col = pick("방문명 컬럼", _visit_opts,
                         cfg.get("visit_col") or guess_kw(raw_common, ("VIS", "방문")) or "(없음)", "cfg_visit",
        help="ex. VISITNM, SUB_VIS_C. 방문명이 없는 연구면 '(없음)'을 선택하세요. 이 경우 방문 누락·방문일 순서 검증은 건너뜁니다.")
    with cc3:
        _vd_opts = ["(없음)"] + raw_common
        visit_date_col = pick("방문일 컬럼", _vd_opts,
                              cfg.get("visit_date_col") or guess_kw(raw_all, ("VISITDT", "방문일", "_DT")) or "(없음)",
                              "cfg_visitdt",
        help="ex. VISITDT, SUB_VIS. 방문일이 없는 연구면 '(없음)'을 선택하세요.")

    _base_default = [b for b in settings_list(cfg.get("base_sheets")) if b in sheet_names] \
        or ([sheet_names[0]] if sheet_names else [])
    # 공백차이 흡수: settings 기준시트를 정규화 매칭
    if cfg.get("base_sheets"):
        _norm_map = {qc.normalize_name(s): s for s in sheet_names}
        _base_default = [_norm_map[qc.normalize_name(b)] for b in settings_list(cfg["base_sheets"])
                         if qc.normalize_name(b) in _norm_map] or _base_default
    ms_state("cfg_base_sheets", sheet_names, _base_default)
    base_sheets = st.multiselect(
    "기준 시트",
    sheet_names,
    key="cfg_base_sheets",
    help="대상자별로 예정된 방문이 존재하는지 확인할 기준 시트를 선택하세요.   \n 선택한 시트에 있는 {Subject ID, 방문명} 조합을 기준으로 다른 시트의 방문 누락 여부를 검사합니다.  \n 방문 정보가 여러 시트에 나뉘어 있으면 해당 시트를 모두 선택하세요.  \n ex. Outcome_NRS, Screening_NRS")
    if visit_col == "(없음)":
        visit_order_map = {}
    else:
        st.markdown("**방문 순서 확인**")
        st.caption("자동으로 감지한 순서대로 나열되어 있습니다. 항목을 마우스로 끌어서 순서를 바꾸세요. "
                   "방문일 순서 검증(5번 화면)에 이 순서가 사용됩니다.")
        _vo_base = qc.build_visit_order_table(raw_sheets, visit_col, sheet_names=base_sheets or None)
        _vo_options = _vo_base["방문명"].tolist()
        _vo_key = "visit_order_sort_" + "_".join(base_sheets) + "_" + visit_col
        visit_order_list = sort_items(_vo_options, direction="vertical", key=_vo_key)
        visit_order_map = {qc.normalize_value(v): i for i, v in enumerate(visit_order_list)}


with st.expander("codebook 컬럼 매핑", expanded=True):
    mc1, mc2, mc3 = st.columns(3)
    with mc1:
        variable_col = pick("변수명 컬럼", cb_cols,
                            cfg.get("variable_col") or guess_kw(cb_cols, ("변수명", "ITEM NAME", "VARIABLE")), "cfg_var")
        _form_opts = ["(없음)"] + cb_cols
        form_col = pick("eCRF명/시트명 컬럼", _form_opts,
                        cfg.get("form_col") or guess_kw(cb_cols, ("ECRF", "CRF", "FORM", "시트")) or "(없음)", "cfg_form",
                        help="코드북에 시트/서식 구분 컬럼이 없으면 '(없음)'을 선택하세요. 이 경우 변수명 대조 결과에 어느 시트 소속인지는 표시되지 않지만, 검증 자체에는 영향이 없습니다.")
    with mc2:
        codelist_col = pick("코드리스트 컬럼", cb_cols,
                            cfg.get("codelist_col") or guess_kw(cb_cols, ("코드리스트", "CODE LIST")), "cfg_codelist")
        res_col = pick("항목명/설명 컬럼", cb_cols,
                       cfg.get("res_col") or guess_kw(cb_cols, ("항목명", "ITEM LABEL", "LABEL")), "cfg_res")
    with mc3:
        _dt_opts = ["(없음)"] + cb_cols
        data_type_col = pick("데이터 유형 컬럼 (선택)", _dt_opts,
                             cfg.get("data_type_col") or guess_kw(cb_cols, ("데이터유형", "DATA TYPE", "유형")) or "(없음)",
                             "cfg_datatype", help = "date, integer 등 변수의 데이터 유형이 기재된 컬럼을 선택하세요. 해당 컬럼이 없으면 '(없음)'을 선택하세요.")

with st.expander("검증 옵션 (선택)", expanded=False):
    oc1, oc2 = st.columns(2)
    with oc1:
        ms_state("cfg_skip_var_names", raw_all, settings_list(cfg.get("skip_var_names")))
        skip_var_names = st.multiselect(
            "코드북 대조 또는 값 검증 제외 변수", raw_all,
            key="cfg_skip_var_names",
            help="'Subject ID', '최종 확인자 성명' 등과 같이 코드리스트에 없지만 문제없는 경우 선택해주세요.  \n 여기에 분류한 변수는 반드시 3.코드리스트 분류 확인에서 'free_text'로 설정해주세요.")

        ms_state("cfg_exclude_cols", raw_all, settings_list(cfg.get("exclude_cols")))
        exclude_cols = st.multiselect(
            "공통값 비교 제외 변수", raw_all,
            key="cfg_exclude_cols",
            help = "'연구자 마지막 확인일'과 같이 동일 변수명을 사용하지만 시트별로 다른 값이 기재될 수 있는 경우 선택해주세요." )
    with oc2:
        _visit_vals = []
        for b in base_sheets:
            if visit_col in raw_sheets[b].columns:
                _visit_vals += [qc.normalize_value(v) for v in raw_sheets[b][visit_col].dropna()]
        _visit_vals = sorted(set(v for v in _visit_vals if v))
        ms_state("cfg_same_day_visits", _visit_vals, settings_list(cfg.get("same_day_visits")))
        same_day_visits = st.multiselect(
            "같은 날 방문 허용", _visit_vals,
            key="cfg_same_day_visits",  help="Screening, Visit1과 같이 방문명은 다르지만 방문일이 같을 수 있는 경우 Screening을 선택해주세요.")
        num_state("cfg_agg_threshold", cfg.get("aggregate_threshold", 10), 2, 1000)
        agg_threshold = st.number_input("집계 임계값", 2, 1000,
                                        key="cfg_agg_threshold",
        help = "같은 오류가 집계 임계값보다 많을 경우 한 행으로 축약하여 출력됩니다.")
        num_state("cfg_pattern_threshold", cfg.get("pattern_unique_threshold", 10), 2, 100)
        pattern_threshold = st.number_input("패턴 임계값", 2, 100,
                                            key="cfg_pattern_threshold",
        help = "한 변수의 unique 값이 패턴 임계값보다 많을 경우 'O'로 출력됩니다.")

st.download_button(
    "현재 설정을 settings.xlsx로 저장",
    data=build_settings_bytes(),
    file_name="settings.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    help="화면에서 고른 설정을 저장합니다. 다음에 이 파일을 올리면 설정이 자동으로 채워집니다.")


# ---------- 3단계: 코드리스트 분류 미리보기·수정 ----------
st.header("3. 코드리스트 분류 확인")
st.caption("classification_reviewed.xlsx이 있으면 그 값을, 없으면 코드리스트를 choice / range / date / free_text로 자동 분류한 결과입니다.   \n 분류가 잘못된 변수는 표에서 codelist_type과 valid_value를 직접 수정한 뒤 아래 [분류 확정]을 누르세요.   \n 구분이 모호한 경우 free_text로 구분하세요. ")


st.markdown("##### 변수 유형 안내")

st.table({
    "유형": ["choice", "range", "date", "free_text"],
    "설명": [
        "입력 가능한 값이 정해진 변수",
        "입력 가능한 숫자 범위가 정해진 변수",
        "날짜가 입력되는 변수",
        "이니셜, 사유, 비고 등 정해진 허용값 없이 자유롭게 입력하는 변수",
    ],
    "valid_value 입력 예시": [
        "0|1|2",
        "0~10",
        "",
        "",
    ],
})

def upload_signature(f):
    """업로드 파일 식별값. 같은 이름으로 다시 올려도 file_id가 바뀌어 변경을 감지한다."""
    if f is None:
        return None
    return (f.name, getattr(f, "size", None), getattr(f, "file_id", None))


# 분류 생성 (classification 파일 있으면 그걸, 없으면 자동 분류)
_cb_norm = qc.normalize_codebook_names(codebook, [variable_col, form_col])
try:
    if classification_file is not None:
        _cls = pd.read_excel(classification_file, sheet_name="classification")
        _cls = _cls.rename(columns={c: qc.normalize_name(c) for c in _cls.columns})
        base_classification = _cls
    else:
        _dt = None if data_type_col == "(없음)" else data_type_col
        base_classification, _ = qc.make_codelist_classification(
            _cb_norm, variable_col, codelist_col, form_col, res_col, data_type_col=_dt)
except Exception as e:
    st.error(f"분류 생성 실패: {e}")
    st.stop()

# 세션에 편집본 보관 (설정 바뀌면 재생성)
_sig = (variable_col, codelist_col, form_col, res_col, data_type_col,
        upload_signature(classification_file))
if st.session_state.get("cls_sig") != _sig:
    st.session_state["cls_sig"] = _sig
    _init = base_classification.copy()
    # unknown 등 4개 외 타입은 free_text로 통일 (분류는 항상 4개만)
    _valid4 = {"choice", "range", "date", "free_text"}
    if "codelist_type" in _init.columns:
        _init.loc[~_init["codelist_type"].isin(_valid4), "codelist_type"] = "free_text"
    st.session_state["cls_edited"] = _init

TYPE_ORDER = ["choice", "range", "date", "free_text"]
_show_cols = [c for c in [variable_col, res_col, codelist_col, "codelist_type", "valid_value"]
              if c in st.session_state["cls_edited"].columns]

cls_df = st.session_state["cls_edited"]
_type_counts = cls_df["codelist_type"].value_counts().to_dict()

_tabs = st.tabs([f"{t} ({_type_counts.get(t, 0)})" for t in TYPE_ORDER])
_edited_parts = {}
_editable = {"codelist_type", "valid_value"}
for _tab, _t in zip(_tabs, TYPE_ORDER):
    with _tab:
        _sub = cls_df[cls_df["codelist_type"] == _t]
        if len(_sub) == 0:
            st.caption(f"{_t}로 분류된 변수가 없습니다.")
            continue
        _sub_show = _sub[_show_cols].copy()
        _sub_show.insert(0, "_ridx", _sub.index)
        _edited_parts[_t] = st.data_editor(
            _sub_show, use_container_width=True, hide_index=True,
            key=f"editor_{_t}",
            column_config={
                "_ridx": None,
                "codelist_type": st.column_config.SelectboxColumn(
                    "codelist_type", options=TYPE_ORDER, required=True),
                "valid_value": st.column_config.TextColumn(
                    "valid_value",
                    help="choice는 '0|1|2' 형식, range는 '0~10' 형식으로 입력"),
            },
            disabled=[c for c in _show_cols if c not in _editable])

if st.button("분류 확정"):
    updated = cls_df.copy()
    _has_valid = "valid_value" in updated.columns
    for _t, _part in _edited_parts.items():
        for _, r in _part.iterrows():
            idx = r["_ridx"]
            if idx not in updated.index:
                continue
            updated.at[idx, "codelist_type"] = r["codelist_type"]
            if _has_valid and "valid_value" in r:
                updated.at[idx, "valid_value"] = "" if pd.isna(r["valid_value"]) else str(r["valid_value"])
    st.session_state["cls_edited"] = updated
    st.session_state["cls_just_confirmed"] = True
    st.rerun()

if st.session_state.pop("cls_just_confirmed", False):
    st.success("분류가 확정되었습니다. 바뀐 타입 기준으로 탭을 갱신했습니다.")

# 확정된 분류를 classification_reviewed.xlsx로 다운로드 (시트명 'classification' → 재업로드 가능)
st.download_button(
    "분류 결과를 classification_reviewed.xlsx로 저장",
    data=qc.to_excel_bytes({"classification": st.session_state["cls_edited"]}),
    file_name="classification_reviewed.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    help="확정한 분류를 저장합니다. 다음에 이 파일을 올리면 분류를 다시 하지 않고 그대로 사용합니다.")

# ---------- 4단계: 날짜 선후관계 규칙 ----------
st.header("4. 날짜 선후관계 규칙 (선택)")
st.caption("date로 분류된 변수끼리 '빠른 날짜 ≤ 늦은 날짜' 규칙을 만듭니다. "
           "예: 발현일 ≤ 소실일.  \n date_rules.xlsx를 업로드하였다면 그대로 표시됩니다.")

# date로 분류된 변수 목록 (확정된 분류 기준)
_cls = st.session_state["cls_edited"]
_date_vars = set(_cls.loc[_cls["codelist_type"] == "date", variable_col]
                 .apply(qc.normalize_name))

# 각 date 변수가 실제로 존재하는 (시트, 변수) 목록
# 변수 -> 항목명 매핑 (드롭다운에 항목명 표시용)
_labels_map = qc.build_label_map(
    qc.normalize_codebook_names(codebook, [variable_col, form_col]),
    variable_col, res_col)

_date_options = []
_date_lookup = {}
for _sname, _sdf in raw_sheets.items():
    for _c in _sdf.columns:
        _cn = qc.normalize_name(_c)
        if _cn in _date_vars:
            _res = _labels_map.get(_cn, "")
            _label = f"{_sname} | {_c}" + (f" ({_res})" if _res else "")
            _date_options.append(_label)
            _date_lookup[_label] = (_sname, str(_c).strip())
if "date_rules_list" not in st.session_state:
    st.session_state["date_rules_list"] = []

# date_rules.xlsx 업로드분을 규칙으로 반영 (파일이 새로 올라오거나 교체될 때마다 다시 읽음)
_dr_sig = upload_signature(date_rule_file)
if date_rule_file is not None and st.session_state.get("date_rules_sig") != _dr_sig:
    try:
        date_rule_file.seek(0)
        _up = pd.read_excel(date_rule_file)
        st.session_state["date_rules_list"] = _up.to_dict("records")
        st.session_state["date_rules_sig"] = _dr_sig
    except Exception as e:
        st.warning(f"date_rules 파일을 읽지 못했습니다: {e}")
elif date_rule_file is None:
    # 파일을 내리면 서명만 초기화 (화면에서 추가한 규칙은 유지, 같은 파일 재업로드 시 다시 반영)
    st.session_state.pop("date_rules_sig", None)

if not _date_options:
    st.info("date로 분류된 변수가 없거나 rawdata에서 찾지 못했습니다. "
            "3번 분류에서 날짜 변수를 date로 지정하세요.")
else:
    dc1, dc2, dc3 = st.columns([2, 2, 1])
    with dc1:
        _earlier = st.selectbox("빠른 쪽 (earlier)", _date_options, key="dr_earlier")
    with dc2:
        _later = st.selectbox("늦은 쪽 (later)", _date_options, key="dr_later")
    with dc3:
        _op = st.selectbox("관계", [1, 2, 3], key="dr_op",
                           format_func=lambda x: {1: "1: ≤ (같은날 허용)",
                                                  2: "2: < (같은날 불허)",
                                                  3: "3: = (같은날)"}[x])
    if st.button("규칙 추가"):
        es, ed = _date_lookup[_earlier]
        ls, ld = _date_lookup[_later]
        st.session_state["date_rules_list"].append({
            "earlier_sheet": es, "earlier_date": ed,
            "later_sheet": ls, "later_date": ld, "operator": _op})
        st.rerun()

# 현재 규칙 목록 표시·삭제·저장
_rules = st.session_state["date_rules_list"]
if _rules:
    st.markdown(f"**현재 규칙 {len(_rules)}개** — 삭제할 규칙의 '선택'에 체크한 뒤 [선택 규칙 삭제]를 누르세요.")

    _rules_view = pd.DataFrame(_rules)
    _rules_view.insert(0, "선택", False)
    _edited_rules = st.data_editor(
        _rules_view, use_container_width=True, hide_index=True,
        key="dr_table",
        column_config={
            "선택": st.column_config.CheckboxColumn("✓", default=False, width="small"),
            "earlier_sheet": st.column_config.TextColumn("earlier_sheet", width="medium"),
            "earlier_date": st.column_config.TextColumn("earlier_date", width="medium"),
            "later_sheet": st.column_config.TextColumn("later_sheet", width="medium"),
            "later_date": st.column_config.TextColumn("later_date", width="medium"),
            "operator": st.column_config.NumberColumn("operator", width="small"),
        },
        disabled=[c for c in _rules_view.columns if c != "선택"])

    if st.button("선택 규칙 삭제"):
        _keep = [r for i, r in enumerate(_rules) if not bool(_edited_rules.iloc[i]["선택"])]
        if len(_keep) != len(_rules):
            st.session_state["date_rules_list"] = _keep
            st.rerun()

    if st.button("전체 규칙 비우기"):
        st.session_state["date_rules_list"] = []
        st.rerun()

    st.download_button(
        "날짜 규칙을 date_rules.xlsx로 저장",
        data=qc.to_excel_bytes({"date_rules": pd.DataFrame(_rules)}),
        file_name="date_rules.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.caption("아직 규칙이 없습니다. 위에서 변수 짝을 골라 [규칙 추가]를 누르세요.")

# ---------- 5단계: 실행 ----------
st.header("5. 전체 검증 실행")
st.caption("코드리스트 분류(위 3번)에서 타입이나 valid_value를 수정했다면, "
           "먼저 [분류 확정]을 눌러야 검증에 반영됩니다.")

if st.button("전체 검증 실행", type="primary"):
    if not base_sheets:
        st.error("기준 시트(base_sheets)를 하나 이상 선택하세요.")
        st.stop()

    cb_norm = qc.normalize_codebook_names(codebook, [variable_col, form_col])

    # 분류 규칙: 화면에서 확인·수정·확정한 분류(cls_edited)를 그대로 사용한다.
    try:
        rules_df = qc.build_codelist_rules(
            st.session_state["cls_edited"], variable_col, codelist_col)
    except Exception as e:
        st.error(f"코드리스트 분류 규칙 생성 실패: {e}")
        st.stop()

    date_rules_df = None
    if st.session_state.get("date_rules_list"):
        date_rules_df = pd.DataFrame(st.session_state["date_rules_list"])
        for _c in ["earlier_sheet", "earlier_date", "later_sheet", "later_date"]:
            if _c in date_rules_df.columns:
                date_rules_df[_c] = date_rules_df[_c].apply(qc.normalize_name)

    config = {
        "subject_col": subject_col,
        "visit_col": None if visit_col == "(없음)" else visit_col,
        "base_list": base_sheets,
        "variable_col": variable_col, "codelist_col": codelist_col,
        "form_col": form_col, "res_col": res_col,
        "exclude_cols": exclude_cols, "skip_var_names": skip_var_names,
        "target_sheets": [], "same_day_visits": same_day_visits,
        "visit_date_col": None if visit_date_col == "(없음)" else visit_date_col,
        "aggregate_threshold": int(agg_threshold),
        "pattern_unique_threshold": int(pattern_threshold),
        "visit_order_map": visit_order_map,
    }

    with st.spinner("검증 실행 중..."):
        st.session_state["outcome"] = qc.run_all_validations(
            raw_sheets, cb_norm, config, rules_df=rules_df, date_rules_df=date_rules_df)
    st.session_state["ran"] = True

# ---------- 6단계: 대시보드 ----------
if st.session_state.get("ran"):
    outcome = st.session_state["outcome"]
    all_issues = outcome["all_issues"]
    results = outcome["results"]
    counts = outcome["counts"]
    total = outcome["total_issues"]

    st.header("6. 결과")

    LABELS = {
    "variable_name_issues": "변수명 대조",
    "codelist_issues": "코드값 검증",
    "visit_missing_issues": "방문 누락",
    "value_mismatch": "공통값 비교",
    "visit_date_order_issues": "방문일 순서",
    "date_order_issues": "날짜 선후관계",
}

    RESULT_HELP = {
    "variable_name_issues":
        "코드북과 rawdata의 변수명이 일치하지 않는 항목입니다. "
        "변수명 오탈자, 누락 또는 시트 매핑 오류인지 확인하세요.",

    "codelist_issues":
        "코드북에 정의된 허용값이나 숫자 범위를 벗어난 항목입니다. "
        "오입력 여부와 변수 분류 및 valid_value 설정을 확인하세요.",

    "visit_missing_issues":
        "대상자별 예상 방문 중 rawdata에서 확인되지 않은 방문입니다. "
        "실제 미방문인지, 방문명 오입력 또는 데이터 누락인지 확인하세요.",

    "value_mismatch":
        "동일 대상자의 공통 정보가 시트마다 다르게 입력된 항목입니다. "
        "각 시트의 값을 대조하고, 시트별로 달라도 되는 변수라면 제외 변수 설정을 확인하세요.",

    "visit_date_order_issues":
        "방문일이 설정된 방문 순서와 맞지 않는 항목입니다. "
        "방문명과 방문일을 확인하고, 같은 날 방문이 가능하면 관련 설정도 확인하세요.",

    "date_order_issues":
        "설정한 날짜 선후관계 규칙을 만족하지 않는 항목입니다. "
        "날짜 입력값과 설정한 날짜 규칙이 올바른지 확인하세요."}

    st.subheader(f"검증별 오류 (총 {total}건)")

    metric_cols = st.columns(len(LABELS))

    for i, (key, label) in enumerate(LABELS.items()):
        with metric_cols[i]:
            st.metric(
                label=label,
                value=counts.get(key, 0),
                help=RESULT_HELP[key],
            )

    st.subheader("전체 오류 목록 (all_issues)")
    st.caption(
    "검출된 오류를 한 번에 확인할 수 있습니다. "
    "검증 유형과 대상자, 변수명, 입력값 등의 정보를 확인하여 원본 데이터를 점검하세요."
)
    if len(all_issues) == 0:
        st.success("검출된 오류가 없습니다.")
    else:
        st.dataframe(all_issues, use_container_width=True, hide_index=True)

    st.subheader("추출 결과 (free_text_unique / 값 패턴)")

    st.caption(
        "free_text_unique는 자유입력 변수에 입력된 값을 중복 없이 정리한 결과입니다.  \n pattern_{시트명}은 각 행에서 중복되는 조합을 제거하여 입력 양상을 보여주며,  \n 값 개수가 '패턴임계값' 이상인 경우 O로 표시합니다.  \n 이를 활용하여 입력값의 조합이나 흐름에 논리적 오류가 없는지 확인할 수 있습니다."
    )

    if results:
        names = list(results.keys())

        for tab, name in zip(st.tabs(names), names):
            with tab:
                df = results[name]
                if name.startswith("pattern_") and len(df) >= 2:
                    _header_idx = set(df.index[:2])

                    def _highlight_header_rows(row):
                        if row.name in _header_idx:
                            return ["background-color: #FFF3CD"] * len(row)
                        return [""] * len(row)

                    st.dataframe(
                        df.style.apply(_highlight_header_rows, axis=1).hide(axis="index"),
                        use_container_width=True,
                    )
                else:
                    st.dataframe(
                        df,
                        use_container_width=True,
                        hide_index=True,
                    )
                
        st.subheader("결과 다운로드")
        dc1, dc2 = st.columns(2)
        with dc1:
            st.download_button("all_issues 다운로드",
                               data=qc.to_excel_bytes({"all_issues": all_issues}),
                               file_name="QC_all_issues.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with dc2:
            if results:
                st.download_button("unique/패턴 다운로드",
                                   data=qc.to_excel_bytes(results),
                                   file_name="QC_unique.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")