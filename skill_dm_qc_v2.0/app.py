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
        if any(k in str(c).upper() for k in ("SUBJ", "SUBNO", "SUBJECT", "대상자", "SCR")):
            return c
    return cols[0] if cols else None


def guess_kw(cols, keywords):
    for c in cols:
        if any(k in str(c).upper() for k in keywords):
            return c
    return None


def pick(label, options, want, key, help=None):
    """want(설정값/자동감지값)가 옵션에 있으면 그 위치를 기본 선택."""
    want = qc.normalize_name(want) if want else None
    norm_opts = {qc.normalize_name(o): o for o in options}
    idx = 0
    if want and want in norm_opts:
        idx = options.index(norm_opts[want])
    return st.selectbox(label, options, index=idx, key=key, help=help)

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
st.caption("rawdata와 codebook을 업로드하면 설정을 확인한 뒤 전체 검증을 한 번에 실행합니다.  \n settings.xlsx와 classification_reviewed.xlsx는 아래 개별 검증에서 생성·다운로드한 후 업로드하여 사용할 수 있으며, 최초 1회만 설정하면 이후에도 계속 사용할 수 있습니다")
st.header("1. 파일 업로드")
st.markdown("**필수 파일**")
r1c1, r1c2 = st.columns(2)
with r1c1:
    raw_file = st.file_uploader("rawdata.xlsx", type=["xlsx", "xls"])
    
with r1c2:
    codebook_file = st.file_uploader("codebook.xlsx", type=["xlsx", "xls"])

st.markdown("**선택 파일**")
st.caption("이전에 이 화면에서 저장한 파일이 있으면 업로드하고 처음이면 비워두세요.")
r2c1, r2c2, r2c3 = st.columns(3)
with r2c1:
    settings_file = st.file_uploader("settings.xlsx (선택)", type=["xlsx", "xls"])
    st.caption("검증에 필요한 설정 파일-단계 2에서 생성됩니다")
with r2c2:
    classification_file = st.file_uploader("classification_reviewed.xlsx (선택)", type=["xlsx", "xls"])
    st.caption("코드리스트 분류 파일-단계 3에서 생성됩니다.")
with r2c3:
    date_rule_file = st.file_uploader("date_rules.xlsx (선택)", type=["xlsx", "xls"])
    st.caption("날짜 규칙 파일 - 단계 4에서 생성됩니다.")

if raw_file is None or codebook_file is None:
    st.info("rawdata와 codebook 파일을 업로드하면 설정 단계가 나타납니다.")
    st.stop()
cfg = load_settings_xlsx(settings_file)
raw_sheets = qc.read_excel_all(raw_file)
codebook = qc.read_first_sheet(codebook_file)
sheet_names = list(raw_sheets.keys())
raw_all = all_columns(raw_sheets)
raw_common = common_columns(raw_sheets) or raw_all
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
            guess_kw(raw_common, ("VIS", "방문"))
        )

        st.session_state["cfg_visitdt"] = match_option(
            cfg.get("visit_date_col"),
            raw_common,
            guess_kw(raw_common, ("VISITDT", "방문일", "_DT"))
        )

        st.session_state["cfg_var"] = match_option(
            cfg.get("variable_col"),
            cb_cols,
            guess_kw(cb_cols, ("변수명", "ITEM NAME", "VARIABLE"))
        )

        st.session_state["cfg_form"] = match_option(
            cfg.get("form_col"),
            cb_cols,
            guess_kw(cb_cols, ("ECRF", "CRF", "FORM", "시트"))
        )

        st.session_state["cfg_codelist"] = match_option(
            cfg.get("codelist_col"),
            cb_cols,
            guess_kw(cb_cols, ("코드리스트", "CODE LIST"))
        )

        st.session_state["cfg_res"] = match_option(
            cfg.get("res_col"),
            cb_cols,
            guess_kw(cb_cols, ("항목명", "ITEM LABEL", "LABEL"))
        )

        data_type_options = ["(없음)"] + cb_cols
        st.session_state["cfg_datatype"] = match_option(
            cfg.get("data_type_col"),
            data_type_options,
            "(없음)"
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
st.caption("settings.xlsx 값이 있으면 그 값을, 없으면 자동 감지값을 기본으로 채웁니다. 확인 후 수정해주세요.")

with st.expander("rawdata 컬럼 매핑", expanded=True):
    cc1, cc2, cc3 = st.columns(3)
    with cc1:
        subject_col = pick("Subject ID 컬럼", raw_common,
                           cfg.get("subject_col") or guess_subject(raw_common), "cfg_subject",
        help="ex. SUBJNO, SCR_NUM")
    with cc2:
        visit_col = pick("방문명 컬럼", raw_common,
                         cfg.get("visit_col") or guess_kw(raw_common, ("VIS", "방문")), "cfg_visit",
        help="ex. VISITNM, SUB_VIS_C")
    with cc3:
        _vd_opts = raw_common
        visit_date_col = pick("방문일 컬럼", _vd_opts,
                              cfg.get("visit_date_col") or guess_kw(raw_all, ("VISITDT", "방문일", "_DT")) or "(없음)",
                              "cfg_visitdt",
        help="ex. VISITDT, SUB_VIS")
    _base_default = [b for b in settings_list(cfg.get("base_sheets")) if b in sheet_names] \
        or ([sheet_names[0]] if sheet_names else [])
    # 공백차이 흡수: settings 기준시트를 정규화 매칭
    if cfg.get("base_sheets"):
        _norm_map = {qc.normalize_name(s): s for s in sheet_names}
        _base_default = [_norm_map[qc.normalize_name(b)] for b in settings_list(cfg["base_sheets"])
                         if qc.normalize_name(b) in _norm_map] or _base_default
    base_sheets = st.multiselect(
    "기준 시트",
    sheet_names,
    default=_base_default,
    key="cfg_base_sheets",
    help="위에서 지정한 Subject ID, 방문명의 모든 조합이 포함되도록 설정해주세요.  \n ex. Outcome_NRS, Screening_NRS")


with st.expander("codebook 컬럼 매핑", expanded=True):
    mc1, mc2, mc3 = st.columns(3)
    with mc1:
        variable_col = pick("변수명 컬럼", cb_cols,
                            cfg.get("variable_col") or guess_kw(cb_cols, ("변수명", "ITEM NAME", "VARIABLE")), "cfg_var")
        form_col = pick("eCRF명/시트명 컬럼", cb_cols,
                        cfg.get("form_col") or guess_kw(cb_cols, ("ECRF", "CRF", "FORM", "시트")), "cfg_form")
    with mc2:
        codelist_col = pick("코드리스트 컬럼", cb_cols,
                            cfg.get("codelist_col") or guess_kw(cb_cols, ("코드리스트", "CODE LIST")), "cfg_codelist")
        res_col = pick("항목명/설명 컬럼", cb_cols,
                       cfg.get("res_col") or guess_kw(cb_cols, ("항목명", "ITEM LABEL", "LABEL")), "cfg_res")
    with mc3:
        _dt_opts = ["(없음)"] + cb_cols
        data_type_col = pick("데이터 유형 컬럼 (선택)", _dt_opts,
                             cfg.get("data_type_col") or guess_kw(cb_cols, ("데이터유형", "DATA TYPE", "유형")) or "(없음)",
                             "cfg_datatype")

with st.expander("검증 옵션 (선택)", expanded=False):
    oc1, oc2 = st.columns(2)
    with oc1:
        skip_var_names = st.multiselect(
            "변수명검증·입력값검증·pattern 출력 제외 변수", raw_all,
            default=[v for v in settings_list(cfg.get("skip_var_names")) if v in raw_all],     key="cfg_skip_var_names",
            help="'Subject ID', '최종 확인자 성명' 등과 같이 코드리스트에 없지만 문제없는 경우 선택해주세요.  \n 여기에 분류한 변수는 반드시 3.코드리스트 분류 확인에서 'free_text'로 설정해주세요.")
        
        exclude_cols = st.multiselect(
            "공통값 비교 제외 변수", raw_all,
            default=[v for v in settings_list(cfg.get("exclude_cols")) if v in raw_all], key="cfg_exclude_cols",
            help = "'연구자 마지막 확인일'과 같이 동일 변수명을 사용하지만 시트별로 다른 값이 기재될 수 있는 경우 선택해주세요." )
    with oc2:
        _visit_vals = []
        for b in base_sheets:
            if visit_col in raw_sheets[b].columns:
                _visit_vals += [qc.normalize_value(v) for v in raw_sheets[b][visit_col].dropna()]
        _visit_vals = sorted(set(v for v in _visit_vals if v))
        if (
            settings_file is not None and "cfg_same_day_visits" not in st.session_state):
                st.session_state["cfg_same_day_visits"] = [value  for value in settings_list(cfg.get("same_day_visits"))
        if value in _visit_vals
    ]


       
        same_day_visits = st.multiselect(
            "같은 날 방문 허용", _visit_vals,
            default=[v for v in settings_list(cfg.get("same_day_visits")) if v in _visit_vals],     key="cfg_same_day_visits",  help="Screening, Visit1과 같이 방문명은 다르지만 방문일이 같을 수 있는 경우 선택해주세요.")
        agg_threshold = st.number_input("집계 임계값", 2, 1000,
                                        int(float(cfg.get("aggregate_threshold", 10))),    key="cfg_agg_threshold", 
        help = "같은 오류가 집계 임계값보다 많을 경우 한 행으로 축약하여 출력됩니다.")
        pattern_threshold = st.number_input("패턴 임계값", 2, 100,
                                            int(float(cfg.get("pattern_unique_threshold", 10))), key="cfg_pattern_threshold",
        help = "한 변수의 unique 값이 패턴 임계값보다 많을 경우 'O'로 출력됩니다.")

st.download_button(
    "현재 설정을 settings.xlsx로 저장",
    data=build_settings_bytes(),
    file_name="settings.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    help="화면에서 고른 설정을 저장합니다. 다음에 이 파일을 올리면 설정이 자동으로 채워집니다.")


# ---------- 3단계: 코드리스트 분류 미리보기·수정 ----------
st.header("3. 코드리스트 분류 확인")
st.caption("코드리스트를 choice / range / date / free_text로 자동 분류한 결과입니다.   \n 타입이 틀린 변수는 표에서 codelist_type과 valid_value를 직접 수정한 뒤 아래 [분류 확정]을 누르세요.   \n 구분이 모호한 경우 free_text로 구분하세요")

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
        classification_file.name if classification_file else None)
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
           "예: 동의일 ≤ 등록일.  \n date_rules.xlsx를 업로드하였다면 그 규칙이 기본으로 들어갑니다.")

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

# date_rules.xlsx 업로드분을 초기 규칙으로 (한 번만)
if date_rule_file is not None and not st.session_state.get("date_rules_loaded"):
    try:
        _up = pd.read_excel(date_rule_file)
        st.session_state["date_rules_list"] = _up.to_dict("records")
        st.session_state["date_rules_loaded"] = True
    except Exception as e:
        st.warning(f"date_rules 파일을 읽지 못했습니다: {e}")

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

    date_rules_df = None
    if st.session_state.get("date_rules_list"):
        date_rules_df = pd.DataFrame(st.session_state["date_rules_list"])
        for _c in ["earlier_sheet", "earlier_date", "later_sheet", "later_date"]:
            if _c in date_rules_df.columns:
                date_rules_df[_c] = date_rules_df[_c].apply(qc.normalize_name)

    config = {
        "subject_col": subject_col, "visit_col": visit_col, "base_list": base_sheets,
        "variable_col": variable_col, "codelist_col": codelist_col,
        "form_col": form_col, "res_col": res_col,
        "exclude_cols": exclude_cols, "skip_var_names": skip_var_names,
        "target_sheets": [], "same_day_visits": same_day_visits,
        "visit_date_col": None if visit_date_col == "(없음)" else visit_date_col,
        "aggregate_threshold": int(agg_threshold),
        "pattern_unique_threshold": int(pattern_threshold),
    }

    with st.spinner("검증 실행 중..."):
        st.session_state["outcome"] = qc.run_all_validations(
            raw_sheets, cb_norm, config, rules_df=rules_df, date_rules_df=date_rules_df)
    st.session_state["ran"] = True

# ---------- 4단계: 대시보드 ----------
if st.session_state.get("ran"):
    outcome = st.session_state["outcome"]
    all_issues = outcome["all_issues"]
    results = outcome["results"]
    counts = outcome["counts"]
    total = outcome["total_issues"]

    st.header("4. 결과")

    LABELS = {
        "variable_name_issues": "변수명 대조",
        "codelist_issues": "코드값 검증",
        "visit_missing_issues": "방문 누락",
        "value_mismatch": "공통값 비교",
        "visit_date_order_issues": "방문일 순서",
        "date_order_issues": "날짜 선후관계",
    }
    st.subheader(f"검증별 오류 (총 {total}건)")
    metric_cols = st.columns(len(LABELS))
    for i, (key, label) in enumerate(LABELS.items()):
        with metric_cols[i]:
            st.metric(label, counts.get(key, 0))

    st.subheader("전체 오류 목록 (all_issues)")
    if len(all_issues) == 0:
        st.success("검출된 오류가 없습니다.")
    else:
        st.dataframe(all_issues, use_container_width=True, hide_index=True)

    if results:
        st.subheader("추출 결과 (free_text unique / 값 패턴)")
        names = list(results.keys())
        for tab, name in zip(st.tabs(names), names):
            with tab:
                st.dataframe(results[name], use_container_width=True, hide_index=True)

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
