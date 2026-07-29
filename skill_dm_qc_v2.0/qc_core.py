"""
qc_core.py
임상 데이터 QC 검증 엔진.
Streamlit에 의존하지 않는 순수 함수 모음 — app.py(UI)와 cli.py(배치)가 공유한다.
기존 app.py에서 검증 함수들을 그대로 옮겨온 것.
"""

import io
import re
import calendar
import pandas as pd

ISSUE_COLS = ["issue_type", "sheet", "variable", "subject_id", "visit",
              "row", "current_value", "expected", "message"]


# ---------------------------------------------------------------
# 공통 유틸
# ---------------------------------------------------------------

def fmt_num(value):
    """10.0 → '10', 0.5 → '0.5' 처럼 불필요한 소수점 제거."""
    f = float(value)
    return str(int(f)) if f.is_integer() else str(f)


def issues_to_df(issues):
    """이슈 리스트를 표준 컬럼 DataFrame으로 변환. 없는 컬럼은 빈칸 처리."""
    return pd.DataFrame(issues, columns=ISSUE_COLS).fillna("")


def build_label_map(codebook, variable_col, res_col):
    """코드북에서 변수명 → 항목명 매핑 생성. res_col이 없으면 빈 dict."""
    if res_col not in codebook.columns:
        return {}
    label_map = {}
    for _, r in codebook.iterrows():
        var = normalize_value(r[variable_col])
        label = normalize_value(r[res_col])
        if var and label:
            label_map[var] = label
    return label_map


def var_label(variable, labels=None):
    """DM 로그 표기: 변수명(항목명). 항목명이 없으면 변수명만."""
    if labels and variable in labels:
        return f"{variable}({labels[variable]})"
    return variable


def normalize_value(value):
    if isinstance(value, pd.Series):
        if len(value) == 0:
            return ""
        value = value.iloc[0]

    if pd.isna(value):
        return ""

    value = str(value).strip()

    # 날짜형 값의 자정 시간 제거
    if value.endswith(" 00:00:00"):
        value = value[:-len(" 00:00:00")]

    if value.endswith(".0"):
        number_part = value[:-2]
        if number_part.replace("-", "").isdigit():
            return number_part

    return value

def normalize_name(value):
    """시트명·변수명 전용: 모든 공백(앞뒤+중간) 제거."""
    s = normalize_value(value)
    return re.sub(r"\s+", "", s)


def normalize_codebook_names(codebook, name_cols):
    """코드북의 이름 컬럼(변수명·eCRF명) 셀값 공백 제거."""
    df = codebook.copy()
    for col in name_cols:
        if col and col in df.columns:
            df[col] = df[col].apply(normalize_name)
    return df


def normalize_display_value(value):
    if pd.isna(value):
        return ""

    if isinstance(value, pd.Timestamp):
        return value.date().isoformat()

    if hasattr(value, "date") and callable(value.date):
        try:
            return value.date().isoformat()
        except Exception:
            pass

    text = str(value).strip()

    if text.endswith(".0"):
        number_part = text[:-2]
        if number_part.replace("-", "").isdigit():
            return number_part

    try:
        num = float(text)
        if num.is_integer():
            return str(int(num))
        return str(num)
    except ValueError:
        pass

    return text


def to_excel_bytes(sheets: dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            if df is None:
                continue
            df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    return output.getvalue()


def write_excel_file(sheets: dict, path: str):
    """CLI용: 결과 dict를 바로 파일로 저장."""
    with open(path, "wb") as f:
        f.write(to_excel_bytes(sheets))


def read_excel_all(file_or_path):
    sheets = pd.read_excel(file_or_path, sheet_name=None)
    out = {}
    for sheet_name, df in sheets.items():
        df = df.rename(columns={c: normalize_name(c) for c in df.columns})
        out[normalize_name(sheet_name)] = df
    return out

def read_first_sheet(file_or_path):
    df = pd.read_excel(file_or_path)
    df = df.rename(columns={c: normalize_name(c) for c in df.columns})
    return df


def get_row_value(row, col):
    if col in row.index:
        return normalize_value(row[col])
    return ""


def sort_values_safely(values):
    # 숫자와 문자가 섞여도 안전하게: (숫자여부, 값) 튜플로 정렬
    def _key(x):
        s = str(x)
        if s.lstrip("-").isdigit():
            return (0, int(s), "")
        return (1, 0, s)
    return sorted(values, key=_key)

def extract_number(value):
    if pd.isna(value):
        return None

    text = str(value).strip()

    if text.endswith(".0"):
        text = text[:-2]

    numbers = re.findall(r"-?\d+", text)

    if not numbers:
        return None

    return int(numbers[0])


# "Screening"처럼 방문명에 숫자가 전혀 없는 임상연구 통상 초기 방문 단계 - extract_number로는
# 순서를 못 뽑아 방문일 순서 검증(validate_visit_date_order)에서 "해석 불가" 오류로 잘못 잡힌다.
# 이런 방문명은 숫자 방문(Visit1, Visit2...)보다 앞선 순서로 처리한다. 다른 스터디에서 새로운
# 비숫자 초기 방문 키워드가 나오면 이 딕셔너리에 "키워드": 순서번호만 추가하면 된다
# (작을수록 이른 방문 - 숫자 방문의 최소값인 1보다 작게 잡아야 한다).
VISIT_ORDER_KEYWORD_OVERRIDES = {
    "SCREENING": -2,
    "ENROLLMENT": -1,
}


def extract_visit_order(value):
    """방문명에서 방문 차수를 뽑는다. extract_number로 숫자를 못 찾으면 VISIT_ORDER_KEYWORD_OVERRIDES에
    등록된 방문 키워드로 순서를 대신 추정한다. 둘 다 실패하면 None(진짜 해석 불가)."""
    num = extract_number(value)
    if num is not None:
        return num
    if pd.isna(value):
        return None
    text = str(value).strip().upper()
    for keyword, order in VISIT_ORDER_KEYWORD_OVERRIDES.items():
        if keyword in text:
            return order
    return None

def parse_partial_date(value, mode):
    """불완전 날짜를 보수적으로 파싱.
    mode='earliest' → 모르는 부분 최소(월01,일01), 'latest' → 최대(월12,말일).
    UK/UNK/00/공백 등을 '모름'으로 인식. 연도조차 모르면 NaT."""
    if value is None:
        return pd.NaT
    s = str(value).strip()
    if s == "" or s.lower() in ("nan", "nat", "none"):
        return pd.NaT
    parts = re.split(r"[-/.]", s)
    if len(parts) < 1:
        return pd.NaT

    def is_unknown(p):
        p = p.strip().upper()
        if p == "": return True
        if p in ("UK","UNK","UKN","UNKN","U","NK"): return True
        if p in ("00","0000","000"): return True
        if set(p) <= {"U","K","N","?","X"}: return True
        return False

    y_raw = parts[0].strip()
    if is_unknown(y_raw) or not y_raw.isdigit():
        return pd.NaT
    year = int(y_raw)
    mon_raw = parts[1].strip() if len(parts) > 1 else ""
    day_raw = parts[2].strip() if len(parts) > 2 else ""

    if is_unknown(mon_raw) or not mon_raw.isdigit():
        month = 1 if mode == "earliest" else 12
    else:
        month = int(mon_raw)
        if not (1 <= month <= 12): return pd.NaT

    if is_unknown(day_raw) or not day_raw.isdigit():
        day = 1 if mode == "earliest" else calendar.monthrange(year, month)[1]
    else:
        day = int(day_raw)
        last = calendar.monthrange(year, month)[1]
        if not (1 <= day <= last): return pd.NaT

    try:
        return pd.Timestamp(year=year, month=month, day=day)
    except Exception:
        return pd.NaT


def format_date_for_message(value):
    if pd.isna(value):
        return ""

    date_value = pd.to_datetime(value, errors="coerce")

    if pd.isna(date_value):
        return str(value)

    return date_value.strftime("%Y-%m-%d")


def remove_empty_rows(df):
    df = df.copy()
    df = df.dropna(how="all")
    if len(df) == 0:
        return df

    # 각 셀이 "비었는지"를 프레임 전체에 벡터 연산으로 판정 (행별 apply 제거)
    as_text = df.astype(str).apply(lambda s: s.str.strip())
    empty_cell = as_text.isin(["nan", "NaN", "NaT", "None", ""])
    # 모든 셀이 빈 행만 제외
    df = df[~empty_cell.all(axis=1)]

    return df


# ---------------------------------------------------------------
# 코드리스트 분류 / 규칙 파싱
# ---------------------------------------------------------------

def classify_codelist(text):
    text = normalize_value(text)

    if text == "":
        return "free_text"

    compact = text.replace(" ", "").upper()

    if re.fullmatch(r"YYYY[-/.]?MM[-/.]?DD", compact):
        return "date"

    if "=" in text:
        return "choice"

    if "_" in text:
        return "choice"
    
    if "," in text:
        return "choice"

    if re.fullmatch(r"-?\d+(\.\d+)?\s*[-~]\s*-?\d+(\.\d+)?", text):
        return "range"

    if re.search(r"\(\s*-?\d+(\.\d+)?\s*[-~]\s*-?\d+(\.\d+)?\s*\)", text):
        return "range"

    equal_numbers = re.findall(r"(-?\d+(\.\d+)?)\s*=", text)
    if len(equal_numbers) >= 2:
        nums = [float(num[0]) for num in equal_numbers]
        if max(nums) - min(nums) >= 5:
            return "range"

    return "free_text"


# 자주 쓰이는 outcome(평가지표) 변수들 - codelist에 허용값/범위가 따로 적혀있지 않아
# free_text로 분류되는 경우가 많다. 변수명(또는 항목명)에 이 키워드가 들어있으면
# free_text 대신 range로 자동 보정하고 codelist_col 텍스트도 채워준다.
# make_codelist_classification 내부에서 자동 적용되므로 qc_core.py를 쓰는 모든 진입점
# (cli.py, app.py, dm-qc 스킬 등) 어디서 호출해도 동일하게 반영된다.
# 새로운 outcome이 추가되면 이 딕셔너리에 "키워드": (최소값, 최대값)만 추가하면 된다.
OUTCOME_RANGE_OVERRIDES = {
    "NRS": (0, 10),
    "VAS": (0, 100),
}

# 이 키워드가 변수명/항목명에 있으면 NRS/VAS가 포함돼 있어도 점수가 아니라
# "코멘트/의견진술" 자유서술 필드이므로 range 자동 보정에서 제외한다.
# (예: NRSCMT, VASCMT, CMT_EQVAS - 전부 "Comment/의견진술" 항목이었음)
OUTCOME_OVERRIDE_EXCLUDE_KEYWORDS = ["CMT", "COMMENT", "코멘트", "의견", "비고", "사유"]


def apply_outcome_overrides(df, variable_col, res_col, codelist_col):
    """codelist에 정의가 없어 free_text로 분류된 항목 중, 이름에 NRS/VAS 등
    알려진 outcome 키워드가 포함된 변수를 range(0~10, 0~100 등)로 자동 보정한다.
    단, CMT/comment/의견 등 코멘트성 필드는 점수가 아니므로 제외한다.
    codelist_col에 이미 값이 있으면 건드리지 않는다(이미 뭔가 정의돼 있던 것이므로).
    반환값: (보정된 df, 보정된 변수 목록[{"variable":..,"keyword":..,"range":".."}, ...])
    """
    df = df.copy()
    overridden = []
    is_free_text = df["codelist_type"] == "free_text"
    name_text = (df.get(variable_col, "").astype(str) + " "
                 + df.get(res_col, "").astype(str)).str.upper()
    is_comment_field = name_text.str.contains(
        "|".join(OUTCOME_OVERRIDE_EXCLUDE_KEYWORDS), na=False)

    for keyword, (lo, hi) in OUTCOME_RANGE_OVERRIDES.items():
        codelist_blank = df[codelist_col].isna() | df[codelist_col].astype(str).str.strip().isin(
            ["", "nan", "none", "None"])
        mask = (is_free_text & codelist_blank & name_text.str.contains(keyword, na=False)
                & ~is_comment_field)
        if not mask.any():
            continue
        range_text = f"{lo}~{hi}"
        df.loc[mask, "codelist_type"] = "range"
        df.loc[mask, codelist_col] = range_text
        for var in df.loc[mask, variable_col].tolist():
            overridden.append({"variable": var, "keyword": keyword, "range": range_text})

    return df, overridden


def make_codelist_classification(codebook, variable_col, codelist_col, form_col, res_col,
                                  data_type_col=None):
    df = codebook.copy()
    df["_variable"] = df[variable_col].apply(normalize_value)
    df["_codelist"] = df[codelist_col].apply(normalize_value)
    df = df[df["_variable"].ne("")].copy()
    df["codelist_type"] = df["_codelist"].apply(classify_codelist)

    if data_type_col and data_type_col in df.columns:
        codelist_blank = df["_codelist"].eq("")
        is_date_type = df[data_type_col].apply(normalize_value).str.upper().eq("DATE")
        df.loc[codelist_blank & is_date_type, "codelist_type"] = "date"

    cols = []
    if form_col in df.columns:
        cols.append(form_col)
    if res_col in df.columns:
        cols.append(res_col)
    cols += [variable_col, codelist_col, "codelist_type"]

    result = df[cols].copy()

    # NRS/VAS 등 outcome 변수는 codelist가 비어 있어도 range로 자동 보정한다.
    # res_col이 없으면(이름/항목명 기반 키워드 매칭 불가) 건너뛴다.
    outcome_overridden = []
    if res_col in result.columns:
        result, outcome_overridden = apply_outcome_overrides(
            result, variable_col, res_col, codelist_col)

    result = add_valid_value_column(result, variable_col, codelist_col)

    summary = result["codelist_type"].value_counts().reset_index()
    summary.columns = ["codelist_type", "count"]
    summary.attrs["outcome_overrides"] = outcome_overridden
    return result, summary


def parse_choice_values(text, dynamic_visit_min=2, dynamic_visit_max=99):
    text = normalize_value(text)
    allowed = {"7777", "8888", "9999"}

    if text == "":
        return sort_values_safely(allowed)

    parts = re.split(r"[|\n;]+", text)

    for part in parts:
        part = part.strip().strip('"')
        if not part:
            continue

        if "=" in part:
            for seg in part.split(","):
                seg = seg.strip()
                if "=" in seg:
                    allowed.add(normalize_value(seg.split("=", 1)[0]))
            continue

        if "," in part:
            code = part.split(",", 1)[0].strip()
            if code:
                allowed.add(normalize_value(code))
            continue

        suffix_match = re.search(r"_([0-9]+)$", part)
        if suffix_match:
            allowed.add(normalize_value(suffix_match.group(1)))
            continue

        allowed.add(normalize_value(part))

    return sort_values_safely({v for v in allowed if v != ""})


def parse_range_values(text):
    text = normalize_value(text)
    if text == "":
        return pd.Series({"min_value": "", "max_value": ""})

    match = re.search(r"(-?\d+(\.\d+)?)\s*[-~]\s*(-?\d+(\.\d+)?)", text)
    if match:
        return pd.Series({"min_value": float(match.group(1)), "max_value": float(match.group(3))})

    equal_numbers = re.findall(r"(-?\d+(\.\d+)?)\s*=", text)
    if len(equal_numbers) >= 2:
        nums = [float(num[0]) for num in equal_numbers]
        return pd.Series({"min_value": min(nums), "max_value": max(nums)})

    return pd.Series({"min_value": "", "max_value": ""})


def build_codelist_rules(classification_df, variable_col, codelist_col):
    df = classification_df.copy()
    type_col = "codelist_type"

    valid_types = ["choice", "range", "free_text", "date"]
    invalid = df[~df[type_col].isin(valid_types)]
    if len(invalid) > 0:
        raise ValueError("codelist_type은 choice/range/free_text/date 중 하나여야 합니다.")

    df[variable_col] = df[variable_col].apply(normalize_value)
    df[codelist_col] = df[codelist_col].apply(normalize_value)
    df[type_col] = df[type_col].apply(normalize_value)

    # valid_value 컬럼이 있으면 사용자가 검토·수정한 값이므로 이를 우선 사용한다.
    # valid_value 형식: choice = '0|1|7777', range = '0.0~10.0' (add_valid_value_column과 동일)
    has_valid = "valid_value" in df.columns
    if has_valid:
        df["_valid"] = df["valid_value"].apply(normalize_value)
    else:
        df["_valid"] = ""

    df["allowed_values"] = ""
    choice_mask = df[type_col] == "choice"
    for i in df.index[choice_mask]:
        v = df.at[i, "_valid"]
        if has_valid and v != "":
            df.at[i, "allowed_values"] = "|".join(
                [x.strip() for x in v.split("|") if x.strip()])
        else:
            df.at[i, "allowed_values"] = "|".join(parse_choice_values(df.at[i, codelist_col]))

    df["min_value"] = pd.Series("", index=df.index, dtype=object)
    df["max_value"] = pd.Series("", index=df.index, dtype=object)
    range_mask = df[type_col] == "range"
    for i in df.index[range_mask]:
        v = df.at[i, "_valid"]
        if has_valid and v != "" and "~" in v:
            parts = v.split("~", 1)
            df.at[i, "min_value"] = parts[0].strip()
            df.at[i, "max_value"] = parts[1].strip()
        else:
            rv = parse_range_values(df.at[i, codelist_col])
            df.at[i, "min_value"] = rv["min_value"]
            df.at[i, "max_value"] = rv["max_value"]

    df = df.drop(columns=["_valid"])
    return df

def add_valid_value_column(classification_df, variable_col, codelist_col):
    """choice/range로 분류된 항목이 실제 QC 실행 시 어떤 값을 허용하는지 valid_value
    컬럼으로 채워서 돌려준다. build_codelist_rules와 동일한 파싱 로직을 그대로 쓰므로
    여기 보이는 값은 실제 검증(validate_codelist)에 쓰이는 허용값과 항상 일치한다.
    이 함수 하나를 app.py(Streamlit)/cli.py/dm-qc 스킬이 공유해서 쓰므로 어느 진입점으로
    분류를 실행해도 classification_reviewed.xlsx의 valid_value가 동일하게 채워진다.
    """
    rules_df = build_codelist_rules(classification_df, variable_col, codelist_col)
    valid_value = pd.Series("", index=classification_df.index, dtype=object)
    choice_mask = (rules_df["codelist_type"] == "choice").values
    # 특수코드(7777/8888/9999)는 validate_codelist가 skip_values로 항상 통과시키므로
    # valid_value 표시에서는 제외한다(화면 가독성). 검증 결과에는 영향 없음.
    _special = {"7777", "8888", "9999"}
    def _drop_special(av):
        return "|".join([v for v in str(av).split("|") if v.strip() and v.strip() not in _special])
    valid_value.loc[choice_mask] = [
        _drop_special(v) for v in rules_df.loc[choice_mask, "allowed_values"].values]
    range_mask = (rules_df["codelist_type"] == "range").values

    valid_value.loc[range_mask] = (
        rules_df.loc[range_mask, "min_value"].astype(str)
        + "~" + rules_df.loc[range_mask, "max_value"].astype(str)
    ).values
    result = classification_df.copy()
    result["valid_value"] = valid_value.values
    return result


# ---------------------------------------------------------------
# 검증 함수들
# ---------------------------------------------------------------

def validate_codelist(raw_sheets, rules_df, variable_col, subject_col, visit_col,
                      skip_values=("7777", "8888", "9999"),
                      labels=None, codelist_col=None):
    issues = []
    skip_values = set(skip_values)

    choice_rules = rules_df[rules_df["codelist_type"] == "choice"].copy()
    range_rules = rules_df[rules_df["codelist_type"] == "range"].copy()

    for _, rule in choice_rules.iterrows():
        variable = normalize_value(rule[variable_col])
        codelist_text = normalize_value(rule[codelist_col]) if codelist_col else ""
        allowed_values = [
            normalize_value(v)
            for v in normalize_value(rule["allowed_values"]).split("|")
            if normalize_value(v)
        ]

        if variable == "" or not allowed_values:
            continue

        for sheet_name, df_raw in raw_sheets.items():
            if variable not in df_raw.columns:
                continue

            df = df_raw.copy().dropna(how="all")
            values_norm = df[variable].apply(normalize_value)

            invalid_rows = df[
                values_norm.ne("")
                & ~values_norm.isin(skip_values)
                & ~values_norm.isin(allowed_values)
            ]

            for idx, row in invalid_rows.iterrows():
                current_value = normalize_value(row[variable])
                allowed_display = "/".join(allowed_values)
                vl = var_label(variable, labels)
                if codelist_text:
                    msg = (f"{vl}의 코드리스트가 \"{codelist_text}\"로 정의되어있으나, "
                           f"'{current_value}'(으)로 입력되어있습니다. 확인해주세요.")
                else:
                    msg = (f"{vl} = '{current_value}'(으)로 입력되어있으나, "
                           f"허용값({allowed_display})에 없는 값입니다. 확인해주세요.")
                issues.append({
                    "issue_type": "choice_value_error",
                    "sheet": sheet_name,
                    "variable": variable,
                    "subject_id": get_row_value(row, subject_col),
                    "visit": get_row_value(row, visit_col),
                    "row": idx + 2,
                    "current_value": current_value,
                    "expected": codelist_text if codelist_text else allowed_display,
                    "message": msg
                })

    for _, rule in range_rules.iterrows():
        variable = normalize_value(rule[variable_col])
        min_value = rule["min_value"]
        max_value = rule["max_value"]

        if variable == "" or pd.isna(min_value) or pd.isna(max_value) or min_value == "" or max_value == "":
            continue

        min_value = float(min_value)
        max_value = float(max_value)

        for sheet_name, df_raw in raw_sheets.items():
            if variable not in df_raw.columns:
                continue

            df = df_raw.copy().dropna(how="all")
            values_norm = df[variable].apply(normalize_value)
            numeric_values = pd.to_numeric(df[variable], errors="coerce")

            has_value_mask = values_norm.ne("") & ~values_norm.isin(skip_values)
            range_display = f"{fmt_num(min_value)}~{fmt_num(max_value)}"

            # 1) 숫자가 아닌 값 (원인이 다르므로 별도 이슈로 분리)
            not_numeric_rows = df[has_value_mask & numeric_values.isna()]

            for idx, row in not_numeric_rows.iterrows():
                current_value = normalize_value(row[variable])
                issues.append({
                    "issue_type": "not_numeric_error",
                    "sheet": sheet_name,
                    "variable": variable,
                    "subject_id": get_row_value(row, subject_col),
                    "visit": get_row_value(row, visit_col),
                    "row": idx + 2,
                    "current_value": current_value,
                    "expected": f"숫자 ({range_display})",
                    "message": (f"{var_label(variable, labels)} = '{current_value}'(으)로 기입되어있습니다. "
                                f"숫자 항목이므로 확인 후 수정해주세요.")
                })

            # 2) 숫자이지만 범위를 벗어난 값
            out_of_range_rows = df[
                has_value_mask
                & numeric_values.notna()
                & ((numeric_values < min_value) | (numeric_values > max_value))
            ]

            for idx, row in out_of_range_rows.iterrows():
                current_value = normalize_value(row[variable])
                issues.append({
                    "issue_type": "range_value_error",
                    "sheet": sheet_name,
                    "variable": variable,
                    "subject_id": get_row_value(row, subject_col),
                    "visit": get_row_value(row, visit_col),
                    "row": idx + 2,
                    "current_value": current_value,
                    "expected": range_display,
                    "message": (f"{var_label(variable, labels)} = '{current_value}'(으)로 기재되어있으나, "
                                f"허용 범위({range_display})를 벗어난 값입니다. 확인해주세요.")
                })

    return issues_to_df(issues)


def make_free_text_unique(raw_sheets, rules_df, variable_col, labels=None):
    exclude_only_values = {"7777", "8888", "9999"}
    labels = labels or {}
    free_text_rules = rules_df[rules_df["codelist_type"] == "free_text"].copy()
    ft_vars = set(free_text_rules[variable_col].dropna().astype(str).str.strip()) - {""}

    rows = []
    # 바깥 루프를 rawdata 시트 순서로 → 결과가 시트 순서대로 정렬됨
    for sheet_name, df_raw in raw_sheets.items():
        df = df_raw.copy().dropna(how="all")
        for col in df.columns:
            variable = str(col).strip()
            if variable not in ft_vars:
                continue

            values = df[variable].apply(normalize_value)
            values = values[values != ""]
            unique_values = sort_values_safely(values.unique())

            if len(unique_values) == 0:
                continue
            if len(unique_values) == 1 and unique_values[0] in exclude_only_values:
                continue

            rows.append({
                "sheet": sheet_name,
                "variable": variable,
                "항목명": labels.get(variable, ""),
                "unique_count": len(unique_values),
                "unique_values": " | ".join(unique_values)
            })

    wide_df = pd.DataFrame(rows, columns=["sheet", "variable", "항목명", "unique_count", "unique_values"])

    if wide_df.empty:
        return wide_df

    # 처음 등장 순서 기록 (정렬 보존용)
    wide_df = wide_df.reset_index(drop=True)
    wide_df["_order"] = range(len(wide_df))

    # 같은 (variable, unique_values)는 한 줄로 묶고, 시트명은 등장 순서대로 합침
    grouped = (
        wide_df
        .groupby(["variable", "unique_values"], as_index=False, sort=False)
        .agg({
            "sheet": lambda x: " | ".join(dict.fromkeys(x)),
            "항목명": "first",
            "unique_count": "first",
            "_order": "min",
        })
    )

    grouped = grouped.sort_values("_order")

    return grouped[["sheet", "variable", "항목명", "unique_count", "unique_values"]]

def validate_variable_names(raw_sheets, codebook, variable_col, form_col,
                            skip_vars=None):
    raw_var_locations = []

    for sheet_name, df in raw_sheets.items():
        for col in df.columns:
            raw_var_locations.append({
                "sheet": sheet_name,
                "variable": str(col).strip()
            })

    raw_var_locations_df = pd.DataFrame(raw_var_locations)
    raw_vars = set(raw_var_locations_df["variable"])

    codebook_vars = (
        codebook[variable_col]
        .dropna()
        .astype(str)
        .str.strip()
    )
    codebook_vars = set(codebook_vars)

    # 검증 제외 변수 (key 변수, 시스템 변수 등 — 코드북에 없어도 정상)
    skip_compact = {normalize_value(v).replace(" ", "").upper()
                    for v in (skip_vars or []) if normalize_value(v)}

    def _skipped(var):
        return normalize_value(var).replace(" ", "").upper() in skip_compact

    missing_in_raw = sorted(v for v in (codebook_vars - raw_vars) if not _skipped(v))
    extra_in_raw = sorted(v for v in (raw_vars - codebook_vars) if not _skipped(v))

    issues = []

    for var in missing_in_raw:
        matched_rows = codebook[
            codebook[variable_col].astype(str).str.strip() == var
        ]

        if len(matched_rows) > 0:
            for _, row in matched_rows.iterrows():
                sheet_value = row.get(form_col, "") if form_col in codebook.columns else ""

                issues.append({
                    "issue_type": "missing_in_raw",
                    "sheet": sheet_value,
                    "variable": var,
                    "subject_id": "",
                    "visit": "",
                    "message": f"코드북에 정의된 변수이나 rawdata에서 확인되지 않습니다. 확인해주세요."
                })
        else:
            issues.append({
                "issue_type": "missing_in_raw",
                "sheet": "",
                "variable": var,
                "subject_id": "",
                "visit": "",
                "message": f"코드북에 정의된 변수이나 rawdata에서 확인되지 않습니다. 확인해주세요."
            })

    extra_locations = raw_var_locations_df[
        raw_var_locations_df["variable"].isin(extra_in_raw)
    ]

    for _, row in extra_locations.iterrows():
        issues.append({
            "issue_type": "extra_in_raw",
            "sheet": row["sheet"],
            "variable": row["variable"],
            "subject_id": "",
            "visit": "",
            "message": "eCRF 코드북에 없는 변수입니다."
        })

    return issues_to_df(issues)


def validate_visit_count_by_base(raw_sheets, base_sheets, subject_col, visit_col):
    """방문별 고유 대상자 수 비교. base_sheets: 시트명 1개(str) 또는 리스트."""
    if isinstance(base_sheets, str):
        base_sheets = [base_sheets]
    base_list = [b for b in base_sheets if b in raw_sheets]
    if not base_list:
        raise ValueError(f"기준 시트가 없습니다: {base_sheets}")

    base_frames = []
    for b in base_list:
        bdf = raw_sheets[b]
        if subject_col not in bdf.columns or visit_col not in bdf.columns:
            continue
        tmp = pd.DataFrame({
            "_subj": bdf[subject_col].apply(normalize_value),
            "_visit_norm": bdf[visit_col].apply(normalize_value),
        })
        tmp = tmp[(tmp["_subj"] != "") & (tmp["_visit_norm"] != "")]
        base_frames.append(tmp)
    if not base_frames:
        raise ValueError(f"기준 시트에 {subject_col}/{visit_col} 컬럼이 없습니다.")

    base_all = pd.concat(base_frames, ignore_index=True).drop_duplicates()
    base_visit_counts = (
        base_all.groupby("_visit_norm")["_subj"].nunique()
        .reset_index(name="base_count").rename(columns={"_visit_norm": "visit"})
    )

    base_set = set(base_list)
    issues = []
    for sheet_name, df_raw in raw_sheets.items():
        if sheet_name in base_set:
            continue
        if subject_col not in df_raw.columns or visit_col not in df_raw.columns:
            continue
        cmp = pd.DataFrame({
            "_subj": df_raw[subject_col].apply(normalize_value),
            "_visit_norm": df_raw[visit_col].apply(normalize_value),
        })
        cmp = cmp[(cmp["_subj"] != "") & (cmp["_visit_norm"] != "")]
        compare_counts = (
            cmp.groupby("_visit_norm")["_subj"].nunique()
            .reset_index(name="compare_count").rename(columns={"_visit_norm": "visit"})
        )
        merged = compare_counts.merge(base_visit_counts, on="visit", how="inner")
        for _, row in merged.iterrows():
            visit = row["visit"]
            compare_count = int(row["compare_count"])
            base_count = int(row["base_count"])
            if compare_count != base_count:
                base_label = "+".join(base_list)
                issues.append({
                    "issue_type": "visit_count_mismatch",
                    "sheet": sheet_name, "variable": visit_col,
                    "subject_id": "", "visit": visit,
                    "current_value": compare_count, "expected": base_count,
                    "message": (
                        f"방문({visit})의 대상자 수가 기준({base_label}, {base_count}명)과 다르게 "
                        f"{compare_count}명으로 확인됩니다. 누락 또는 중복 여부 확인해주세요."
                    )
                })
    return issues_to_df(issues)

def validate_same_values_by_base(raw_sheets, base_sheets, subject_col, visit_col,
                                 compare_cols=None, labels=None, exclude_cols=None,
                                 target_sheets=None):
    """기준 시트(들)와 다른 시트의 공통 변수 값 비교. base_sheets: 시트명 1개(str) 또는 리스트."""
    if isinstance(base_sheets, str):
        base_sheets = [base_sheets]
    base_list = [b for b in base_sheets if b in raw_sheets]
    if not base_list:
        raise ValueError(f"기준 시트가 없습니다: {base_sheets}")

    base_df = pd.concat([raw_sheets[b] for b in base_list], ignore_index=True)
    base_set = set(base_list)

    if subject_col not in base_df.columns:
        raise ValueError(f"기준 시트에 '{subject_col}' 컬럼이 없습니다.")

    issues = []

    allow = set(target_sheets) if target_sheets else None

    for sheet_name, df_raw in raw_sheets.items():
        if sheet_name in base_set:
            continue
        if allow is not None and sheet_name not in allow:
            continue

        df = df_raw.copy()

        if subject_col not in df.columns:
            continue

        if visit_col in df.columns and visit_col in base_df.columns:
            key_cols = [subject_col, visit_col]
        else:
            key_cols = [subject_col]

        exclude = {str(c).strip() for c in (exclude_cols or [])}
        exclude |= {subject_col, visit_col}

        if compare_cols:
            # 명시 지정 시 그 목록만
            candidates = compare_cols
        else:
            # 자동: 기준 시트와 대상 시트의 공통 컬럼 전부
            candidates = [
                col for col in df.columns
                if str(col).strip() != "" and str(col).strip() not in exclude
            ]

        available_compare_cols = [
            col for col in candidates
            if col in df.columns and col in base_df.columns
        ]

        if not available_compare_cols:
            continue

        needed_base_cols = key_cols + available_compare_cols
        needed_df_cols = key_cols + available_compare_cols

        missing_base_cols = [col for col in needed_base_cols if col not in base_df.columns]

        if missing_base_cols:
            continue

        base_tmp = base_df[needed_base_cols].copy()
        df_tmp = df[needed_df_cols].copy()
        df_tmp["_row"] = df_tmp.index + 2

        for col in needed_base_cols:
            base_tmp[col] = base_tmp[col].apply(normalize_display_value)

        for col in needed_df_cols:
            df_tmp[col] = df_tmp[col].apply(normalize_display_value)

        base_compare_df = (
            base_tmp
            .dropna(subset=key_cols)
            .drop_duplicates(subset=key_cols)
        )

        compare_df = (
            df_tmp
            .dropna(subset=key_cols)
            .copy()
        )

        merged = compare_df.merge(
            base_compare_df[key_cols + available_compare_cols],
            on=key_cols,
            how="inner",
            suffixes=("_current", "_base")
        )

        for _, row in merged.iterrows():
            subject_id = row[subject_col]
            visit = row[visit_col] if visit_col in key_cols else ""

            for col in available_compare_cols:
                current_value = row[f"{col}_current"]
                base_value = row[f"{col}_base"]

                if current_value != base_value:
                    issues.append({
                        "issue_type": "value_mismatch_by_key",
                        "sheet": sheet_name,
                        "variable": col,
                        "subject_id": subject_id,
                        "visit": visit,
                        "row": row["_row"],
                        "current_value": current_value,
                        "expected": base_value,
                        "message": (
                            f"기준 시트({'+'.join(base_list)})의 {var_label(col, labels)}와 값이 다르게 기입되어있습니다. "
                            f"(기준: {base_value}, 해당 시트: {current_value}) "
                        )
                    })

    return issues_to_df(issues)


AGGREGATABLE_TYPES = {
    "choice_value_error",
    "range_value_error",
    "not_numeric_error",
    "missing_value",
    "value_mismatch_by_key",
}


def aggregate_mass_issues(issues_df, threshold=10, labels=None, sheet_totals=None):
    """같은 (시트, 변수, 오류유형)이 threshold건 이상이면 한 행으로 요약.

    전수/대량 오류는 행 단위 나열보다 값 분포 요약이 실용적이다.
    (예: 코드리스트가 "1=예|2=아니오"로 정의되어있으나, '3'(118건)으로 입력되어있습니다.)

    sheet_totals: {시트명: 유효 행 수} — 결측이 전체 행이면 "모든 값" 문구 사용.
    """
    from collections import Counter

    if issues_df.empty or threshold <= 0:
        return issues_df

    sheet_totals = sheet_totals or {}
    out_frames = []

    for (itype, sheet, variable), grp in issues_df.groupby(
        ["issue_type", "sheet", "variable"], sort=False
    ):
        if itype not in AGGREGATABLE_TYPES or len(grp) < threshold:
            out_frames.append(grp)
            continue

        n = len(grp)
        subjects = {normalize_value(s) for s in grp["subject_id"] if normalize_value(s)}
        m = len(subjects)
        # 대상자 10명 이하면 전부 나열, 초과 시 인원수만 표시
        if m == 0:
            subj_disp = ""
        elif m <= 10:
            subj_disp = " ".join(sorted(subjects))
        else:
            subj_disp = f"대상자 {m}명"
        vl = var_label(variable, labels)
        expected = normalize_value(grp["expected"].iloc[0])

        values = [normalize_value(v) for v in grp["current_value"] if normalize_value(v) != ""]
        vc = Counter(values)
        top = sorted(vc.items(), key=lambda x: -x[1])
        val_disp = ", ".join(f"'{v}'({c}건)" for v, c in top[:5])
        if len(top) > 5:
            val_disp += f" 외 {len(top) - 5}종"

        if itype == "choice_value_error":
            msg = (f"{vl}의 코드리스트가 \"{expected}\"로 정의되어있으나, "
                   f"{val_disp}(으)로 입력되어있습니다 ({subj_disp}). "
                   f"전체적으로 확인해주세요.")
        elif itype == "range_value_error":
            msg = (f"{vl}의 허용 범위({expected})를 벗어난 값이 총 {n}건({subj_disp}) 있습니다. "
                   f"입력값: {val_disp}. 전체적으로 확인해주세요.")
        elif itype == "not_numeric_error":
            msg = (f"{vl}에 숫자가 아닌 값이 총 {n}건({subj_disp}) 입력되어있습니다. "
                   f"입력값: {val_disp}. 전체적으로 확인해주세요.")
        elif itype == "missing_value":
            total = sheet_totals.get(sheet)
            if total and n >= total:
                msg = f"{vl} 변수의 모든 값이 기재되어있지 않습니다. (총 {n}건)"
            else:
                msg = f"{vl} 값이 기재되지 않은 행이 총 {n}건({subj_disp}) 있습니다. 확인해주세요."
        else:  # value_mismatch_by_key
            msg = (f"{vl}가 기준 시트와 다르게 기입된 행이 총 {n}건({subj_disp}) 있습니다. "
                   f"확인 후 수정해주세요.")

        out_frames.append(pd.DataFrame([{
            "issue_type": itype,
            "sheet": sheet,
            "variable": variable,
            "subject_id": subj_disp,
            "visit": "",
            "row": f"{n}건",
            "current_value": "",
            "expected": expected,
            "message": msg,
        }], columns=ISSUE_COLS))

    return pd.concat(out_frames, ignore_index=True)


def count_valid_rows(raw_sheets, subject_col=""):
    """시트별 유효 행 수 (잔여 행 제외). 집계 시 '모든 값 결측' 판정용."""
    totals = {}
    for sheet_name, df_raw in raw_sheets.items():
        df = remove_empty_rows(df_raw)
        if subject_col and subject_col in df.columns:
            df = df[df[subject_col].apply(normalize_value) != ""]
        totals[sheet_name] = len(df)
    return totals


def merge_visit_rows(issues_df):
    """같은 대상자의 완전히 동일한 이슈(메시지까지 동일)가 방문만 다르게
    여러 행으로 나오면, visit과 row를 합쳐 한 행으로 병합한다.
    예: S4-091 방문 101/102/103 동일 오류 → visit = "101, 102, 103"
    """
    if issues_df.empty:
        return issues_df

    df = issues_df.fillna("").astype({c: str for c in issues_df.columns})

    group_cols = ["issue_type", "sheet", "variable", "subject_id",
                  "current_value", "expected", "message"]

    def _join(series):
        seen = list(dict.fromkeys(v for v in series if str(v).strip() != ""))
        return ", ".join(seen)

    merged = (
        df.groupby(group_cols, sort=False, as_index=False)
        .agg(visit=("visit", _join), row=("row", _join))
    )

    return merged[ISSUE_COLS]




def make_value_patterns(raw_sheets, subject_col, visit_col,
                        exclude_cols=None, unique_threshold=10,
                        labels=None, codelists=None):
    """시트별 행 조합 패턴 + 빈도. labels/codelists 있으면 맨 위에 설명 2줄 추가."""
    exclude_set = set(exclude_cols or []) | {subject_col, visit_col}
    special_codes = {"7777", "8888", "9999"}
    missing_text = {"", "nan", "nat", "none", "null", "na", "n/a"}
    labels = labels or {}
    codelists = codelists or {}
    out = {}
    for sheet_name, df_raw in raw_sheets.items():
        df = remove_empty_rows(df_raw)
        if subject_col in df.columns:
            df = df[df[subject_col].apply(normalize_value) != ""]
        if len(df) == 0:
            continue

        cols = [c for c in df.columns
                if str(c).strip() and str(c).strip() not in exclude_set]
        if not cols:
            continue

        pattern_df = pd.DataFrame(index=df.index)
        for c in cols:
            vals = df[c].apply(normalize_value)
            uniq = {v for v in vals if v.lower() not in missing_text}
            is_value_type = len(uniq) < unique_threshold

            def _conv(v, value_type=is_value_type):
                s = normalize_value(v)
                if s.lower() in missing_text:
                    return "Null"
                if s in special_codes:
                    return s
                return s if value_type else "O"

            pattern_df[c] = vals.apply(_conv)

        grouped = (pattern_df.groupby(list(cols), dropna=False)
                   .size().reset_index(name="count")
                   .sort_values(list(cols)))

        # 맨 위에 항목명 / 코드리스트 두 줄 추가
        if labels or codelists:
            label_row = {c: labels.get(normalize_value(c), "") for c in cols}
            label_row["count"] = "← 항목명"
            code_row = {c: codelists.get(normalize_value(c), "") for c in cols}
            code_row["count"] = "← 코드리스트"
            header = pd.DataFrame([label_row, code_row])
            grouped = pd.concat([header, grouped], ignore_index=True)

        out[sheet_name] = grouped

    return out


def validate_visit_date_order(raw_sheets, base_sheet, subject_col, visit_order_col, visit_date_col,
                              labels=None, same_day_visits=None):
    if base_sheet not in raw_sheets:
        raise ValueError(f"기준 시트가 없습니다: {base_sheet}")

    base_df = raw_sheets[base_sheet].copy()

    needed_cols = [subject_col, visit_order_col, visit_date_col]
    missing_cols = [col for col in needed_cols if col not in base_df.columns]

    if missing_cols:
        raise ValueError(f"기준 시트에 필요한 컬럼이 없습니다: {missing_cols}")

    same_day_set = {normalize_value(v) for v in (same_day_visits or [])}

    base_df["_subject_norm"] = base_df[subject_col].apply(normalize_value)
    base_df["_visit_order_num"] = base_df[visit_order_col].apply(extract_visit_order)
    base_df["_visit_date"] = pd.to_datetime(base_df[visit_date_col], errors="coerce")
    base_df["_row_number"] = base_df.index + 2

    issues = []

    invalid_rows = base_df[
        base_df["_subject_norm"].notna()
        & (
            base_df["_visit_order_num"].isna()
            | base_df["_visit_date"].isna()
        )
    ]

    for _, row in invalid_rows.iterrows():
        issues.append({
            "issue_type": "visit_date_parse_error",
            "sheet": base_sheet,
            "variable": f"{visit_order_col}, {visit_date_col}",
            "subject_id": row["_subject_norm"],
            "visit": normalize_value(row[visit_order_col]),
            "row": row["_row_number"],
            "current_value": f"{visit_order_col}={row[visit_order_col]}, {visit_date_col}={row[visit_date_col]}",
            "message": (
                f"{var_label(visit_order_col, labels)} 또는 {var_label(visit_date_col, labels)} 값을 "
                f"해석할 수 없습니다. 입력 형식 확인해주세요."
            )
        })

    check_df = base_df.dropna(
        subset=["_subject_norm", "_visit_order_num", "_visit_date"]
    ).copy()

    for subject_id, group in check_df.groupby("_subject_norm"):
        group = group.sort_values("_visit_order_num")

        previous_date = None
        previous_visit = None

        for _, row in group.iterrows():
            current_date = row["_visit_date"]
            current_visit = normalize_value(row[visit_order_col])

            if previous_date is not None:
                allow_same = previous_visit in same_day_set
                if allow_same:
                    violated = current_date < previous_date
                    order_desc = "이른"
                else:
                    violated = current_date <= previous_date
                    order_desc = "이르거나 같은"

                if violated:
                    issues.append({
                        "issue_type": "visit_date_order_mismatch",
                        "sheet": base_sheet,
                        "variable": visit_date_col,
                        "subject_id": subject_id,
                        "visit": current_visit,
                        "row": row["_row_number"],
                        "current_value": current_date.date().isoformat(),
                        "expected": f"{previous_date.date().isoformat()} 이후",
                        "message": (
                            f"방문({current_visit})의 {var_label(visit_date_col, labels)} = "
                            f"'{current_date.date()}'이 이전 방문({previous_visit}) = "
                            f"'{previous_date.date()}'보다 {order_desc} 날짜로 기재되어있습니다. 확인해주세요."
                        )
                    })

            previous_date = current_date
            previous_visit = current_visit

    return issues_to_df(issues)


JAMO_RE = re.compile(r"[ㄱ-ㅎㅏ-ㅣ]")       # 자모 단독 (ㄴ, ㅁ 등) = 키보드 오타
HANGUL_RE = re.compile(r"[가-힣]")          # 완성형 한글 (연구에 따라 정상일 수 있음)


def _suggest_sheet(name, raw_sheets):
    """시트명을 못 찾았을 때 가장 유사한 실제 시트명 제안."""
    import difflib
    # 공백 무시 완전 일치 우선 (가장 흔한 오타)
    compact = name.replace(" ", "")
    for s in raw_sheets:
        if s.replace(" ", "") == compact:
            return s
    matches = difflib.get_close_matches(name, list(raw_sheets), n=1, cutoff=0.6)
    return matches[0] if matches else None


def validate_date_order_rules(raw_sheets, date_rules_df, subject_col, visit_col, labels=None):
    required_rule_cols = [
        "earlier_sheet",
        "earlier_date",
        "later_sheet",
        "later_date",
        "operator"
    ]

    missing_rule_cols = [
        col for col in required_rule_cols
        if col not in date_rules_df.columns
    ]

    if missing_rule_cols:
        raise ValueError(f"날짜 규칙 파일에 필요한 컬럼이 없습니다: {missing_rule_cols}")

    issues = []

    for _, rule in date_rules_df.iterrows():
        earlier_sheet = normalize_value(rule["earlier_sheet"])
        earlier_col = normalize_value(rule["earlier_date"])
        later_sheet = normalize_value(rule["later_sheet"])
        later_col = normalize_value(rule["later_date"])
        operator_raw = normalize_value(rule["operator"])

        if operator_raw == "":
            operator_raw = "1"

        operator_map = {
            "1": "<=",
            "2": "<",
            "3": "=="
        }

        if operator_raw not in operator_map:
            issues.append({
                "issue_type": "invalid_date_rule_operator",
                "sheet": f"{earlier_sheet}, {later_sheet}",
                "variable": f"{earlier_col}, {later_col}",
                "subject_id": "",
                "visit": "",
                "message": (
                    f"operator는 1, 2, 3만 가능합니다. "
                    f"1=이전날짜 허용(≤), 2=이전날짜 불허(<), 3=같은 날짜(=). 현재값={operator_raw}"
                )
            })
            continue

        operator = operator_map[operator_raw]

        if earlier_sheet == "" or earlier_col == "" or later_sheet == "" or later_col == "":
            continue

        if earlier_sheet not in raw_sheets:
            issues.append({
                "issue_type": "date_rule_sheet_not_found",
                "sheet": earlier_sheet,
                "variable": earlier_col,
                "subject_id": "",
                "visit": "",
                "message": (
                    f"earlier_sheet '{earlier_sheet}'가 rawdata에 없습니다."
                    + (f" 혹시 '{_suggest_sheet(earlier_sheet, raw_sheets)}'인가요? "
                       f"규칙 파일의 시트명을 확인해주세요."
                       if _suggest_sheet(earlier_sheet, raw_sheets) else "")
                )
            })
            continue

        if later_sheet not in raw_sheets:
            issues.append({
                "issue_type": "date_rule_sheet_not_found",
                "sheet": later_sheet,
                "variable": later_col,
                "subject_id": "",
                "visit": "",
                "message": (
                    f"later_sheet '{later_sheet}'가 rawdata에 없습니다."
                    + (f" 혹시 '{_suggest_sheet(later_sheet, raw_sheets)}'인가요? "
                       f"규칙 파일의 시트명을 확인해주세요."
                       if _suggest_sheet(later_sheet, raw_sheets) else "")
                )
            })
            continue

        earlier_df = raw_sheets[earlier_sheet].copy().dropna(how="all")
        later_df = raw_sheets[later_sheet].copy().dropna(how="all")

        if earlier_col not in earlier_df.columns:
            issues.append({
                "issue_type": "date_rule_column_not_found",
                "sheet": earlier_sheet,
                "variable": earlier_col,
                "subject_id": "",
                "visit": "",
                "message": f"{earlier_sheet} 시트에 {earlier_col} 컬럼이 없습니다."
            })
            continue

        if later_col not in later_df.columns:
            issues.append({
                "issue_type": "date_rule_column_not_found",
                "sheet": later_sheet,
                "variable": later_col,
                "subject_id": "",
                "visit": "",
                "message": f"{later_sheet} 시트에 {later_col} 컬럼이 없습니다."
            })
            continue

        if subject_col not in earlier_df.columns or subject_col not in later_df.columns:
            issues.append({
                "issue_type": "date_rule_key_column_not_found",
                "sheet": f"{earlier_sheet}, {later_sheet}",
                "variable": subject_col,
                "subject_id": "",
                "visit": "",
                "message": f"두 시트 모두에 {subject_col} 컬럼이 있어야 합니다."
            })
            continue

        if visit_col in earlier_df.columns and visit_col in later_df.columns:
            key_cols = [subject_col, visit_col]
        else:
            key_cols = [subject_col]

        earlier_tmp = earlier_df[key_cols + [earlier_col]].copy()
        later_tmp = later_df[key_cols + [later_col]].copy()

        earlier_tmp["_row_earlier"] = earlier_tmp.index + 2
        later_tmp["_row_later"] = later_tmp.index + 2

        for col in key_cols:
            earlier_tmp[col] = earlier_tmp[col].apply(normalize_value)
            later_tmp[col] = later_tmp[col].apply(normalize_value)

        earlier_tmp["_earlier_date"] = earlier_tmp[earlier_col].apply(lambda v: parse_partial_date(v, "earliest"))
        later_tmp["_later_date"] = later_tmp[later_col].apply(lambda v: parse_partial_date(v, "latest"))

        earlier_tmp["_earlier_raw"] = earlier_tmp[earlier_col]
        later_tmp["_later_raw"] = later_tmp[later_col]

        if earlier_sheet == later_sheet:
            # 같은 시트: 같은 행에서 두 날짜 컬럼을 직접 비교 (merge/중복제거 안 함)
            same = earlier_df.copy()
            same["_row"] = same.index + 2
            same["_earlier_date"] = same[earlier_col].apply(lambda v: parse_partial_date(v, "earliest"))
            same["_later_date"] = same[later_col].apply(lambda v: parse_partial_date(v, "latest"))
            same = same.dropna(subset=["_earlier_date", "_later_date"])
            merged = pd.DataFrame({
                subject_col: (same[subject_col].apply(normalize_value)
                              if subject_col in same.columns else ""),
                "_earlier_date": same["_earlier_date"].values,
                "_later_date": same["_later_date"].values,
                "_earlier_raw": same[earlier_col].values,
                "_later_raw": same[later_col].values,
                "_row_earlier": same["_row"].values,
                "_row_later": same["_row"].values,
            })
            if visit_col in same.columns:
                merged[visit_col] = same[visit_col].apply(normalize_value).values
        else:
            # 다른 시트: subject(+visit) 키로 merge
            earlier_tmp = earlier_tmp.drop_duplicates(subset=key_cols)
            later_tmp = later_tmp.drop_duplicates(subset=key_cols)
            merged = earlier_tmp.merge(
                later_tmp,
                on=key_cols,
                how="inner"
            )

        if operator == "<=":
            invalid_mask = merged["_earlier_date"] > merged["_later_date"]
        elif operator == "==":
            invalid_mask = merged["_earlier_date"] != merged["_later_date"]
        else:  # "<"
            invalid_mask = merged["_earlier_date"] >= merged["_later_date"]

        invalid_rows = merged[invalid_mask]

        for _, row in invalid_rows.iterrows():
            subject_id = row[subject_col] if subject_col in merged.columns else ""
            visit = row[visit_col] if visit_col in merged.columns else ""

            # 출력은 원본 값 그대로 (UK 등 불완전 표기 유지). 비교만 변환값 사용.
            earlier_date_msg = normalize_value(row["_earlier_raw"])
            later_date_msg = normalize_value(row["_later_raw"])

            issues.append({
                "issue_type": "date_order_error",
                "sheet": f"{later_sheet}",
                "variable": f"{later_col}",
                "subject_id": subject_id,
                "visit": visit,
                "row": f"{row['_row_later']}",
                "current_value": f"{earlier_col}={earlier_date_msg}, {later_col}={later_date_msg}",
                "expected": f"{earlier_col} {operator} {later_col}",
                "message": (
                    (f"{var_label(earlier_col, labels)} = '{earlier_date_msg}'와 "
                     f"{var_label(later_col, labels)} = '{later_date_msg}'가 "
                     f"서로 다른 날짜로 기입되어있습니다. 같아야 합니다. 확인해주세요. "
                     if operator == "==" else
                     f"{var_label(earlier_col, labels)} = '{earlier_date_msg}'이 "
                     f"{var_label(later_col, labels)} = '{later_date_msg}'보다 "
                     f"{'늦은' if operator == '<=' else '늦거나 같은'} 날짜로 기입되어있습니다. "
                     f"확인해주세요. ")
                    + f"({earlier_sheet} {row['_row_earlier']}행, {later_sheet} {row['_row_later']}행)"
                )
            })

    return issues_to_df(issues)

def validate_visit_missing_by_base(raw_sheets, base_sheets, subject_col, visit_col, labels=None):
    """대상자별 방문 누락 검증 (중도탈락 고려).
    기준 명단(base_sheets)의 (대상자,방문)을 정답으로, 각 시트가 실제 쓰는
    방문과 교집합 내서 그 시트에 빠진 (대상자,방문)을 누락으로 잡는다.
    """
    if isinstance(base_sheets, str):
        base_sheets = [base_sheets]
    base_list = [b for b in base_sheets if b in raw_sheets]
    if not base_list:
        raise ValueError(f"기준 시트가 없습니다: {base_sheets}")

    base_pairs = {}
    for b in base_list:
        bdf = raw_sheets[b]
        if subject_col not in bdf.columns or visit_col not in bdf.columns:
            continue
        for subj, visit in zip(bdf[subject_col].apply(normalize_value),
                               bdf[visit_col].apply(normalize_value)):
            if subj == "" or visit == "":
                continue
            base_pairs.setdefault(subj, set()).add(visit)

    base_set = set(base_list)
    issues = []

    for sheet_name, df_raw in raw_sheets.items():
        if sheet_name in base_set:
            continue
        if subject_col not in df_raw.columns or visit_col not in df_raw.columns:
            continue

        subj_series = df_raw[subject_col].apply(normalize_value)
        visit_series = df_raw[visit_col].apply(normalize_value)

        sheet_visits = {v for v in visit_series if v != ""}
        if not sheet_visits:
            continue

        present = set(zip(subj_series, visit_series))

        for subj, base_visits in base_pairs.items():
            expected = base_visits & sheet_visits
            for visit in sorted(expected, key=lambda x: (extract_visit_order(x) if extract_visit_order(x) is not None else 9e9, x)):
                if (subj, visit) not in present:
                    issues.append({
                        "issue_type": "visit_missing",
                        "sheet": sheet_name,
                        "variable": visit_col,
                        "subject_id": subj,
                        "visit": visit,
                        "current_value": "",
                        "expected": visit,
                        "message": (
                            f"방문({visit})이 누락되어있습니다. "
                            f"기준 명단에는 있으나 이 시트에 없습니다. 확인해주세요."
                        )
                    })

    return issues_to_df(issues)


# ---------------------------------------------------------------
# 대량 동일 오류 집계
# ---------------------------------------------------------------

AGGREGATABLE_TYPES = {
    "choice_value_error",
    "range_value_error",
    "not_numeric_error",
    "missing_value",
    "value_mismatch_by_key",
}


def run_all_validations(raw_sheets, codebook, config, rules_df=None,
                        date_rules_df=None, only=None, progress=None):
    """모든 QC 검증을 순서대로 실행하고 결과를 통합해 반환.

    cli.py / app.py 공용. 검증 로직은 각 validate_* 함수 그대로 사용하며,
    이 함수는 '실행 순서 + 통합·집계·정렬'만 담당한다.

    Parameters
    ----------
    raw_sheets : dict[str, DataFrame]   read_excel_all 결과
    codebook   : DataFrame              read_first_sheet + normalize_codebook_names 완료본
    config     : dict                   아래 키 사용
        subject_col, visit_col, base_list(list), variable_col, codelist_col,
        form_col, res_col, exclude_cols(list), skip_var_names(list),
        target_sheets(list), same_day_visits(list), visit_date_col(str|None),
        aggregate_threshold(int), pattern_unique_threshold(int)
    rules_df   : DataFrame | None       코드리스트 분류 규칙(build_codelist_rules 결과).
                                        None이면 codelist/free_text/pattern 검증은 건너뜀.
    date_rules_df : DataFrame | None    날짜 선후관계 규칙. None이면 해당 검증 건너뜀.
    only       : list[str] | None       지정 시 그 검증만 실행.
    progress   : callable | None        진행 메시지 콜백 progress(name, count).

    Returns
    -------
    dict: {
        "all_issues": DataFrame,          # 통합·집계·정렬된 이슈 목록
        "results": dict[str, DataFrame],  # free_text_unique, pattern_* 등 부가 결과
        "counts": dict[str, int],         # 검증별 이슈 건수
        "total_issues": int,
    }
    """
    c = config
    subject_col = c["subject_col"]
    visit_col = c["visit_col"]
    base_list = c["base_list"]
    base_sheet = base_list[0] if base_list else None
    variable_col = c["variable_col"]
    codelist_col = c["codelist_col"]
    form_col = c["form_col"]
    res_col = c["res_col"]
    exclude_cols = c.get("exclude_cols", []) or []
    skip_var_names = c.get("skip_var_names", []) or []
    target_sheets = c.get("target_sheets", []) or []
    same_day_visits = c.get("same_day_visits", []) or []
    visit_date_col = c.get("visit_date_col")
    agg_threshold = int(c.get("aggregate_threshold", 10))
    pattern_threshold = int(c.get("pattern_unique_threshold", 10))

    labels = build_label_map(codebook, variable_col, res_col)

    if target_sheets:
        compare_targets = [x for x in target_sheets if x in raw_sheets]
    else:
        compare_targets = [name for name in raw_sheets if name not in set(base_list)]

    def should_run(name):
        return only is None or name in only

    results = {}
    issue_dfs = []
    counts = {}

    def report(name, df):
        n = len(df)
        counts[name] = n
        if progress:
            progress(name, n)
        if n > 0:
            issue_dfs.append(df)

    # 1. 변수명 대조
    if should_run("variable_names"):
        report("variable_name_issues",
               validate_variable_names(raw_sheets, codebook, variable_col, form_col,
                                       skip_vars=skip_var_names))

    # 2. 허용값·범위 검증
    # skip_var_names로 지정한 변수는 코드값(valid_value) 검증에서도 제외한다.
    # (변수명 검증·패턴은 이미 제외됨 → 이 변수는 모든 검증에서 빠지는 셈)
    if should_run("codelist") and rules_df is not None:
        if skip_var_names:
            _skip_norm = {normalize_name(v) for v in skip_var_names}
            _rules_cl = rules_df[~rules_df[variable_col].apply(normalize_name).isin(_skip_norm)].copy()
        else:
            _rules_cl = rules_df
        report("codelist_issues",
               validate_codelist(raw_sheets, _rules_cl, variable_col,
                                 subject_col, visit_col,
                                 labels=labels, codelist_col=codelist_col))

    # 3. 자유입력 고유값 목록
    if should_run("free_text_unique") and rules_df is not None:
        df = make_free_text_unique(raw_sheets, rules_df, variable_col, labels=labels)
        results["free_text_unique"] = df
        if progress:
            progress("free_text_unique", len(df))

    # 3-2. 시트별 값 패턴
    if should_run("value_patterns"):
        patterns = make_value_patterns(
            raw_sheets, subject_col, visit_col,
            exclude_cols=exclude_cols + skip_var_names,
            unique_threshold=pattern_threshold,
            labels=labels,
            codelists=build_label_map(codebook, variable_col, codelist_col))
        for sheet_name, pdf in patterns.items():
            results[f"pattern_{sheet_name}"[:31]] = pdf
        if progress:
            progress("value_patterns", len(patterns))

    # 4. 대상자별 방문 누락
    if should_run("visit_missing"):
        report("visit_missing_issues",
               validate_visit_missing_by_base(raw_sheets, base_list, subject_col, visit_col,
                                              labels=labels))

    # 5. 시트간 공통변수 값 비교
    if should_run("value_compare"):
        report("value_mismatch",
               validate_same_values_by_base(
                   raw_sheets, base_sheet, subject_col, visit_col,
                   exclude_cols=exclude_cols + skip_var_names, labels=labels,
                   target_sheets=compare_targets))

    # 6. 방문순서 ↔ 방문일 정합성
    if should_run("visit_date_order") and visit_date_col:
        report("visit_date_order_issues",
               validate_visit_date_order(
                   raw_sheets, base_sheet, subject_col,
                   visit_col, visit_date_col, labels=labels,
                   same_day_visits=same_day_visits))

    # 7. 날짜 선후관계 규칙
    if should_run("date_rules") and date_rules_df is not None:
        report("date_order_issues",
               validate_date_order_rules(
                   raw_sheets, date_rules_df, subject_col, visit_col, labels=labels))

    # --- 통합·집계·병합·정렬 ---
    FINAL_COLS = ["sheet", "variable", "subject_id", "visit", "row", "message"]

    if issue_dfs:
        all_issues = pd.concat(issue_dfs, ignore_index=True)
        all_issues = aggregate_mass_issues(
            all_issues, threshold=agg_threshold, labels=labels,
            sheet_totals=count_valid_rows(raw_sheets, subject_col))
        all_issues = merge_visit_rows(all_issues)

        sheet_order = {name: i for i, name in enumerate(raw_sheets.keys())}

        def _sheet_key(v):
            v = str(v)
            if v in sheet_order:
                return sheet_order[v]
            return sheet_order.get(v.split("->")[0].strip(), len(sheet_order))

        all_issues["_sheet_sort"] = all_issues["sheet"].map(_sheet_key)
        all_issues["_visit_sort"] = pd.to_numeric(
            all_issues["visit"].astype(str).str.split(",").str[0].str.strip(),
            errors="coerce")
        all_issues = all_issues.sort_values(
            ["_sheet_sort", "subject_id", "_visit_sort", "visit"], na_position="last")
        all_issues = all_issues[FINAL_COLS]
    else:
        all_issues = pd.DataFrame(columns=FINAL_COLS)

    return {
        "all_issues": all_issues,
        "results": results,
        "counts": counts,
        "total_issues": int(sum(counts.values())),
    }
