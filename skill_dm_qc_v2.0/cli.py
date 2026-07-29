"""
cli.py — settings.xlsx 기반 임상 데이터 QC 배치 실행 (1차 기본 검증)

1차 검증 항목:
  1. 코드리스트 자동분류·타입확정 (classification 생성)
  2. 변수명 대조 (코드북 ↔ raw)
  3. 허용값·범위 검증 (choice/range)
  4. 자유입력 고유값 목록
  5. 방문별 행수 비교
  6. 시트간 공통변수 값 비교

사용법:
    python cli.py settings.xlsx
    python cli.py settings.xlsx --round 1
    python cli.py settings.xlsx --raw 새스냅샷.xlsx
    python cli.py settings.xlsx --only codelist visit_count

settings.xlsx 형식: 첫 시트에 [구분 | 값] 2개 컬럼.
의존성: pip install pandas openpyxl
"""

import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

import qc_core as qc

# Windows 콘솔 기본 코드페이지(cp949 등)가 UTF-8이 아닌 경우가 많아 한글 출력 시
# 깨짐/에러가 발생할 수 있다. 콘솔 설정과 무관하게 항상 UTF-8로 출력하도록 강제한다.
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

# 구분(key) 별칭: 한글/영문 어느 쪽으로 적어도 인식
ALIASES = {
    "연구명": "study",
    "raw": "rawdata", "raw_data": "rawdata", "로데이터": "rawdata",
    "코드북": "codebook",
    "분류": "classification", "분류파일": "classification",
    "subject_id": "subject_col", "대상자컬럼": "subject_col",
    "visit": "visit_col", "방문컬럼": "visit_col",
    "primary_sheet": "base_sheets", "기준시트": "base_sheets",
    "기준시트들": "base_sheets",
    "비교제외": "exclude_cols", "비교제외변수": "exclude_cols",
    "변수명컬럼": "variable_col", "변수명": "variable_col",
    "코드리스트컬럼": "codelist_col", "코드리스트": "codelist_col",
    "데이터유형컬럼": "data_type_col", "데이터유형": "data_type_col",
    "ecrf명": "form_col", "ecrf명컬럼": "form_col",
    "항목명": "res_col", "항목명컬럼": "res_col",
    "검사시트": "target_sheets",
    "방문일컬럼": "visit_date_col",
    "날짜규칙": "date_rules", "날짜규칙파일": "date_rules",
    "변수명대조제외": "skip_var_names", "변수검증제외": "skip_var_names",
    "집계임계값": "aggregate_threshold",
    "패턴임계값": "pattern_unique_threshold", "unique임계값": "pattern_unique_threshold",
    "같은날허용방문": "same_day_visits", "동일날짜방문": "same_day_visits",
}

def load_settings(path: Path):
    df = pd.read_excel(path)
    df = df.iloc[:, :2]
    df.columns = ["key", "value"]

    settings = {}
    for _, r in df.iterrows():
        key = qc.normalize_value(r["key"]).strip().lower().replace(" ", "_")
        value = qc.normalize_value(r["value"])
        if key == "" or key.startswith("#") or value == "":
            continue
        key = ALIASES.get(key, key)
        settings[key] = value
    return settings


def as_list(value):
    if not value:
        return []
    return [s.strip() for s in re.split(r"[,|\n]+", str(value)) if s.strip()]


def resolve(base_dir: Path, value):
    if not value:
        return None
    p = Path(value)
    if not p.is_absolute():
        p = base_dir / p
    return p if p.exists() else None


def main():
    parser = argparse.ArgumentParser(description="임상 데이터 QC 배치 실행 (1차)")
    parser.add_argument("settings", nargs="?", default="settings.xlsx",
                        help="설정 엑셀 파일 (기본: settings.xlsx)")
    parser.add_argument("--raw", help="rawdata 경로 (설정값 덮어쓰기)")
    parser.add_argument("--codebook", help="codebook 경로 (덮어쓰기)")
    parser.add_argument("--classification", help="분류 파일 경로 (덮어쓰기)")
    parser.add_argument("--date-rules", help="날짜 규칙 파일 경로 (덮어쓰기)")
    parser.add_argument("--out", help="결과 Excel 경로")
    parser.add_argument("--round", type=int, help="QC 차수 (결과 파일명에 사용)")
    parser.add_argument("--only", nargs="*", help="지정한 검증만 실행")
    args = parser.parse_args()

    settings_path = Path(args.settings)
    if not settings_path.exists():
        sys.exit(f"오류: 설정 파일이 없습니다: {settings_path}")

    base_dir = settings_path.parent
    s = load_settings(settings_path)

    # 필수 설정 확인
    for key, name in [("subject_col", "subject_col(대상자 컬럼)"),
                      ("visit_col", "visit_col(방문 컬럼)"),
                      ("base_sheets", "base_sheets(기준 시트)")]:
        if key not in s:
            sys.exit(f"오류: settings.xlsx에 '{name}' 행이 필요합니다.")

    raw_path = args.raw or resolve(base_dir, s.get("rawdata"))
    codebook_path = args.codebook or resolve(base_dir, s.get("codebook"))
    classification_path = args.classification or resolve(base_dir, s.get("classification"))
    date_rules_path = args.date_rules or resolve(base_dir, s.get("date_rules"))

    if not raw_path or not codebook_path:
        sys.exit("오류: rawdata / codebook 파일을 찾지 못했습니다. "
                 "settings.xlsx의 파일명과 실제 파일 위치를 확인하세요.")

    subject_col = qc.normalize_name(s["subject_col"])
    visit_col = qc.normalize_name(s["visit_col"])
    exclude_cols = [qc.normalize_name(x) for x in as_list(s.get("exclude_cols"))]
    target_sheets = [qc.normalize_name(x) for x in as_list(s.get("target_sheets"))]
    skip_var_names = [qc.normalize_name(x) for x in as_list(s.get("skip_var_names"))]
    same_day_visits = [qc.normalize_value(x) for x in as_list(s.get("same_day_visits"))]
    visit_date_col = qc.normalize_name(s.get("visit_date_col")) if s.get("visit_date_col") else None
    base_list = [qc.normalize_name(x) for x in as_list(s.get("base_sheets"))]
    variable_col = qc.normalize_name(s["variable_col"])
    codelist_col = qc.normalize_name(s["codelist_col"])
    form_col = qc.normalize_name(s["form_col"])
    res_col = qc.normalize_name(s["res_col"])
    data_type_col = qc.normalize_name(s["data_type_col"]) if s.get("data_type_col") else None
    agg_threshold = int(float(s.get("aggregate_threshold", "10")))
    pattern_threshold = int(float(s.get("pattern_unique_threshold", 10)))

    print(f"[load] settings: {settings_path}")
    print(f"[load] rawdata: {raw_path}")
    raw_sheets = qc.read_excel_all(raw_path)
    print(f"[load] 시트 {len(raw_sheets)}개")

    print(f"[load] codebook: {codebook_path}")
    codebook = qc.read_first_sheet(codebook_path)
    codebook = qc.normalize_codebook_names(codebook, [variable_col, form_col])
    print(f"[load] 항목명 매핑: {len(qc.build_label_map(codebook, variable_col, res_col))}개 변수")

    # classification: 있으면 로드, 없으면 코드북에서 자동 생성
    rules_df = None
    if classification_path:
        print(f"[load] classification: {classification_path}")
        classification_edited = pd.read_excel(classification_path, sheet_name="classification")
        classification_edited = classification_edited.rename(
            columns={c: qc.normalize_name(c) for c in classification_edited.columns})
        rules_df = qc.build_codelist_rules(classification_edited, variable_col, codelist_col)
    else:
        print("[info] classification 파일이 없어 코드북에서 자동 분류합니다.")
        auto_cls, _ = qc.make_codelist_classification(
            codebook, variable_col, codelist_col, form_col, res_col,
            data_type_col=data_type_col)
        rules_df = qc.build_codelist_rules(auto_cls, variable_col, codelist_col)

    # 날짜 규칙 파일 로드 (있을 때만)
    date_rules_df = None
    if date_rules_path:
        date_rules_df = pd.read_excel(date_rules_path)
        for _c in ["earlier_sheet", "earlier_date", "later_sheet", "later_date"]:
            if _c in date_rules_df.columns:
                date_rules_df[_c] = date_rules_df[_c].apply(qc.normalize_name)

    # --- 전체 검증 실행 (cli/app 공용 함수) ---
    config = {
        "subject_col": subject_col, "visit_col": visit_col, "base_list": base_list,
        "variable_col": variable_col, "codelist_col": codelist_col,
        "form_col": form_col, "res_col": res_col,
        "exclude_cols": exclude_cols, "skip_var_names": skip_var_names,
        "target_sheets": target_sheets, "same_day_visits": same_day_visits,
        "visit_date_col": visit_date_col,
        "aggregate_threshold": agg_threshold,
        "pattern_unique_threshold": pattern_threshold,
    }

    def _progress(name, n):
        print(f"[check] {name}: {n}")

    only = args.only if args.only is not None else None
    outcome = qc.run_all_validations(
        raw_sheets, codebook, config,
        rules_df=rules_df, date_rules_df=date_rules_df,
        only=only, progress=_progress)

    all_issues = outcome["all_issues"]
    results = outcome["results"]
    total_issues = outcome["total_issues"]

    # --- 저장: settings 폴더 아래 results/, 시분초 포함 ---
    # all_issues(집계된 이슈 목록)와 나머지 결과(free_text_unique, 값 패턴 등)를
    # 별도 파일로 저장한다.
    study = s.get("study", "study")
    round_part = f"_{args.round}차" if args.round else ""
    out_dir = base_dir / "results"
    out_dir.mkdir(exist_ok=True)
    base_name = f"{study}_QC{round_part}_{datetime.now():%y%m%d_%H%M%S}"
    if args.out:
        out_base = Path(args.out)
        out_stem = out_base.with_suffix("")
        all_issues_path = str(out_stem) + "_all_issues" + out_base.suffix
        unique_path = str(out_stem) + "_unique" + out_base.suffix
    else:
        all_issues_path = str(out_dir / f"{base_name}_all_issues.xlsx")
        unique_path = str(out_dir / f"{base_name}_unique.xlsx")

    qc.write_excel_file({"all_issues": all_issues}, all_issues_path)
    qc.write_excel_file(results if results else {"unique": pd.DataFrame()}, unique_path)
    print(f"\n[done] 총 이슈 {total_issues}건 → {all_issues_path}")
    print(f"[done] unique/패턴 결과 -> {unique_path}")
    sys.exit(1 if total_issues > 0 else 0)


if __name__ == "__main__":
    main()
