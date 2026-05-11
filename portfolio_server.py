"""
Portfolio MCP Server v14
v13 대비 변경사항:
- search_portfolio, get_statistics, get_company_history 반환에 _query_meta 추가
  → 실제 DB 조회 기반 해석 확인 정보 (적용된 필터, 매칭 건수, 주의사항)
  → Custom Instructions에서 이 블록을 [검색 해석]으로 출력
"""

import os
import json
from pathlib import Path

import pandas as pd
from mcp.server.fastmcp import FastMCP

# ── 설정 ─────────────────────────────────────────────────────────────
EXCEL_PATH = os.environ.get(
    "PORTFOLIO_EXCEL_PATH",
    str(Path(__file__).parent / "sample_portfolio.xlsx")
)

# ── 허용값 상수 ───────────────────────────────────────────────────────
VALID_METRICS = {
    "count_companies", "latest_total_amount", "sum_round_amount",
    "avg_amount", "median_amount", "avg_equity", "count_rows", "count_exited",
}

VALID_GROUP_BY = {
    "분야", "지역", "투자유형", "당시 투자 Round 정보", "Exit 여부"
}

VALID_SORT_BY = {
    "합계 투자금액(M$)", "당시 투자 Round 금액(M$)", "지분율(%)", "투자년도", "기업명"
}

METRIC_ALIASES: dict[str, str] = {
    # latest_total_amount (구 sum_amount 포함 — 기업별 최신 누적액 합계)
    "sum_amount": "latest_total_amount",
    "투자금액합계": "latest_total_amount", "누적투자금액": "latest_total_amount",
    "총투자금액": "latest_total_amount", "투자규모": "latest_total_amount",
    "total_investment": "latest_total_amount", "total_invested": "latest_total_amount",
    "investment_sum": "latest_total_amount", "amount_sum": "latest_total_amount",
    # sum_round_amount (기간 내 집행 금액 합계)
    "기간투자금액": "sum_round_amount", "집행금액합계": "sum_round_amount",
    "투자금액": "sum_round_amount", "합계": "sum_round_amount",
    "round_amount_sum": "sum_round_amount", "invested_amount_sum": "sum_round_amount",
    # avg_amount
    "평균투자금액": "avg_amount", "평균금액": "avg_amount", "평균": "avg_amount",
    "평균라운드금액": "avg_amount", "avg_investment": "avg_amount",
    "avg_invested": "avg_amount", "average_investment": "avg_amount",
    "avg_round_amount": "avg_amount", "invested_amount_avg": "avg_amount",
    # median_amount
    "중앙값": "median_amount", "중앙투자금액": "median_amount",
    "median_investment": "median_amount", "median_round": "median_amount",
    # avg_equity
    "평균지분율": "avg_equity", "지분율": "avg_equity", "avg_equity_rate": "avg_equity",
    # count_companies
    "포트폴리오사수": "count_companies", "회사수": "count_companies",
    "기업수": "count_companies", "포트폴리오수": "count_companies",
    "company_count": "count_companies", "count_company": "count_companies",
    "num_companies": "count_companies",
    # count_rows
    "투자건수": "count_rows", "건수": "count_rows", "라운드수": "count_rows",
    "row_count": "count_rows", "investment_count": "count_rows",
    # count_exited
    "exit수": "count_exited", "엑싯수": "count_exited", "exit기업수": "count_exited",
    "num_exited": "count_exited", "exited_count": "count_exited",
}

GROUP_BY_ALIASES: dict[str, str] = {
    "sector": "분야", "섹터": "분야", "industry": "분야",
    "region": "지역", "국가": "지역", "나라": "지역", "country": "지역", "location": "지역",
    "type": "투자유형", "유형": "투자유형", "investment_type": "투자유형",
    "round": "당시 투자 Round 정보", "라운드": "당시 투자 Round 정보",
    "round_info": "당시 투자 Round 정보", "stage": "당시 투자 Round 정보",
    "exit": "Exit 여부", "엑싯": "Exit 여부", "exit_status": "Exit 여부",
    "company": "기업명",
}


# ── 데이터 로더 ───────────────────────────────────────────────────────
def load_data() -> pd.DataFrame:
    if not Path(EXCEL_PATH).exists():
        raise FileNotFoundError(f"엑셀 파일을 찾을 수 없습니다: {EXCEL_PATH}")
    try:
        df = pd.read_excel(EXCEL_PATH)
    except Exception as e:
        raise RuntimeError(f"엑셀 파일 읽기 실패: {e}")
    for col in ["투자년도", "투자월", "합계 투자금액(M$)", "당시 투자 Round 금액(M$)", "지분율(%)"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


def get_db_meta() -> dict:
    df = load_data()
    return {
        "sectors":          sorted(df["분야"].dropna().unique().tolist()),
        "regions":          sorted(df["지역"].dropna().unique().tolist()),
        "rounds":           sorted(df["당시 투자 Round 정보"].dropna().unique().tolist()),
        "investment_types": sorted(df["투자유형"].dropna().unique().tolist()),
        "exit_statuses":    sorted(df["Exit 여부"].dropna().unique().tolist()),
        "companies":        sorted(df["기업명"].dropna().unique().tolist()),
    }


def df_to_records(df: pd.DataFrame) -> list:
    records = []
    for _, row in df.iterrows():
        record = {}
        for k, v in row.items():
            if pd.isna(v):
                record[k] = None
            elif isinstance(v, float):
                record[k] = int(v) if v == int(v) and abs(v) < 1e15 else round(v, 4)
            else:
                record[k] = v
        records.append(record)
    return records


def apply_list_filter(df: pd.DataFrame, col: str, val) -> pd.DataFrame:
    if isinstance(val, list):
        return df[df[col].isin(val)]
    return df[df[col] == val]


def normalize_metric(m: str) -> str:
    if m in VALID_METRICS:
        return m
    return METRIC_ALIASES.get(m.strip(), m)


def normalize_group_by(g: str) -> str:
    if g in VALID_GROUP_BY:
        return g
    return GROUP_BY_ALIASES.get(g.strip(), g)


def error_response(msg: str) -> str:
    return json.dumps({"error": msg, "action_required": "파라미터를 수정하여 다시 호출하세요."}, ensure_ascii=False)


# ── App Description ───────────────────────────────────────────────────
def _build_app_description() -> str:
    try:
        meta = get_db_meta()
        sectors_str     = ", ".join(meta["sectors"])
        regions_str     = ", ".join(meta["regions"])
        rounds_str      = ", ".join(meta["rounds"])
        inv_types_str   = ", ".join(meta["investment_types"])
        exit_status_str = ", ".join(meta["exit_statuses"])
    except Exception:
        sectors_str = regions_str = rounds_str = inv_types_str = exit_status_str = "(엑셀 로드 실패)"

    return f"""포트폴리오 투자 DB. 모든 데이터는 이 서버에서만 조회. Claude 자체 지식으로 답변 금지.

■ DB 실제값 (파라미터 매핑 시 반드시 아래 값 중 하나를 사용)
  분야:     {sectors_str}
  지역:     {regions_str}
  라운드:   {rounds_str}
  투자유형: {inv_types_str}
  Exit 여부: {exit_status_str}

■ 매핑 규칙
  - 사용자 표현을 위 DB 실제값으로 변환하여 파라미터에 전달
  - 복수 대상이 명확한 경우 배열로 전달 (OR 조건)
  - 필드가 다른 조건 간에는 AND 조건 적용

■ get_statistics metrics 허용값 (영어만, 한국어 사용 금지)
  count_companies     = 고유 포트폴리오사 수 (기간 필터 적용)
  latest_total_amount = 기업별 최신 누적 투자액 합계 ※ 전체 기간 기준, year 필터 무관
  sum_round_amount    = 기간 내 집행된 라운드 금액 합계 ※ year_from/to 필터와 반드시 함께 사용
  avg_amount          = 평균 라운드 투자금액 (기간 필터 적용)
  median_amount       = 라운드 투자금액 중앙값 (기간 필터 적용)
  avg_equity          = 평균 지분율 (전환채권 제외, 기간 필터 적용)
  count_rows          = 투자 건수·라운드 수 (기간 필터 적용)
  count_exited        = Exit 완료 기업 수 ※ 전체 기간 기준

■ metric 선택 기준
  - 기간 한정 투자 흐름 분석 → sum_round_amount + year_from/year_to 필수
  - 전체 포트폴리오 규모 파악 → latest_total_amount
  - 분포 파악 (평균 왜곡 우려 시) → median_amount 병행

■ get_statistics group_by 허용값
  분야 | 지역 | 투자유형 | 당시 투자 Round 정보 | Exit 여부"""


# ── FastMCP 앱 초기화 ─────────────────────────────────────────────────
app = FastMCP("portfolio-mcp", instructions=_build_app_description())


# ══════════════════════════════════════════════════════════════════════
# TOOLS
# ══════════════════════════════════════════════════════════════════════

@app.tool()
def search_portfolio(
    company_name: str | None = None,
    sector: str | list[str] | None = None,
    region: str | list[str] | None = None,
    investment_type: str | list[str] | None = None,
    exit_status: str | list[str] | None = None,
    round_info: str | list[str] | None = None,
    year_from: int | None = None,
    year_to: int | None = None,
    min_total_amount: float | None = None,
    max_total_amount: float | None = None,
    min_equity: float | None = None,
    max_equity: float | None = None,
    latest_only: bool = False,
    sort_by: str = "투자년도",
    sort_asc: bool = False,
    limit: int = 200,
) -> str:
    """
    포트폴리오 다중 조건 검색.

    Args:
        company_name: 기업명 부분 일치
        sector: 투자 분야. 배열 가능 (OR)
        region: 투자 지역. 배열 가능 (OR)
        investment_type: 지분투자 | 전환채권
        exit_status: Y=Exit완료, N=미완료
        round_info: 투자 라운드. 배열 가능 (OR)
        year_from: 투자 시작 연도 이후 (포함)
        year_to: 투자 종료 연도 이전 (포함)
        min_total_amount: 누적 투자금액 최솟값 (M$)
        max_total_amount: 누적 투자금액 최댓값 (M$)
        min_equity: 지분율 최솟값 (%)
        max_equity: 지분율 최댓값 (%)
        latest_only: True=기업별 최신 라운드만
        sort_by: 정렬 기준. 허용값: 합계 투자금액(M$) | 당시 투자 Round 금액(M$) | 지분율(%) | 투자년도 | 기업명
        sort_asc: True=오름차순, False=내림차순
        limit: 최대 반환 행 수 (기본 200)
    """
    try:
        if sort_by not in VALID_SORT_BY:
            sort_by = "투자년도"

        df = load_data()

        # 필터 적용 + 실제 적용된 조건 추적
        applied = {}
        warnings = []

        if company_name:
            df = df[df["기업명"].str.contains(company_name, na=False, case=False)]
            applied["기업명"] = f"'{company_name}' 포함"
        if sector:
            df = apply_list_filter(df, "분야", sector)
            applied["분야"] = sector
        if region:
            df = apply_list_filter(df, "지역", region)
            applied["지역"] = region
        if investment_type:
            valid_types = df["투자유형"].dropna().unique().tolist()
            types = investment_type if isinstance(investment_type, list) else [investment_type]
            invalid_types = [t for t in types if t not in valid_types]
            if invalid_types:
                warnings.append(f"투자유형 {invalid_types}은 DB에 없는 값입니다. 허용값: {valid_types}")
            valid_input = [t for t in types if t in valid_types]
            if valid_input:
                df = apply_list_filter(df, "투자유형", valid_input if len(valid_input) > 1 else valid_input[0])
                applied["투자유형"] = valid_input
        if exit_status:
            valid_exits = df["Exit 여부"].dropna().unique().tolist()
            exits = exit_status if isinstance(exit_status, list) else [exit_status]
            invalid_exits = [e for e in exits if e not in valid_exits]
            if invalid_exits:
                warnings.append(f"Exit 여부 {invalid_exits}은 DB에 없는 값입니다. 허용값: {valid_exits}")
            valid_input = [e for e in exits if e in valid_exits]
            if valid_input:
                df = apply_list_filter(df, "Exit 여부", valid_input if len(valid_input) > 1 else valid_input[0])
                applied["Exit 여부"] = valid_input
        if round_info:
            if isinstance(round_info, list):
                df = df[df["당시 투자 Round 정보"].isin(round_info)]
            else:
                df = df[df["당시 투자 Round 정보"] == round_info]
            applied["라운드"] = round_info
        if year_from is not None or year_to is not None:
            if year_from is not None:
                df = df[df["투자년도"] >= year_from]
            if year_to is not None:
                df = df[df["투자년도"] <= year_to]
            applied["투자년도"] = f"{year_from or ''}~{year_to or ''}"
        if min_total_amount is not None or max_total_amount is not None:
            if min_total_amount is not None:
                df = df[df["합계 투자금액(M$)"] >= min_total_amount]
            if max_total_amount is not None:
                df = df[df["합계 투자금액(M$)"] <= max_total_amount]
            applied["합계 투자금액(M$)"] = f"{min_total_amount or ''}~{max_total_amount or ''}"
        if min_equity is not None or max_equity is not None:
            if min_equity is not None:
                df = df[df["지분율(%)"] >= min_equity]
            if max_equity is not None:
                df = df[df["지분율(%)"] <= max_equity]
            applied["지분율(%)"] = f"{min_equity or ''}~{max_equity or ''}"
        if latest_only:
            df = df.sort_values(["투자년도", "투자월"]).groupby("기업명").last().reset_index()
            applied["latest_only"] = "기업별 최신 라운드만"

        # 조건 적용 후 0건인 필드 경고
        if len(df) == 0 and applied:
            warnings.append("적용된 조건으로 매칭된 결과가 없습니다. 조건을 확인하세요.")

        if sort_by in df.columns:
            df = df.sort_values(sort_by, ascending=sort_asc)

        total = len(df)
        truncated = total > limit
        if truncated:
            warnings.append(f"결과 {total}건 중 {limit}건만 반환. 조건을 좁히거나 limit을 조정하세요.")
        df = df.head(limit)

        result = {
            "_query_meta": {
                "적용된 조건": applied if applied else "없음 (전체 조회)",
                "매칭된 기업 수": int(df["기업명"].nunique()) if "기업명" in df.columns else None,
                "매칭된 총 건수": total,
                "주의": warnings if warnings else None,
            },
            "returned": len(df),
            "records": df_to_records(df),
        }
        return json.dumps(result, ensure_ascii=False, indent=2)

    except Exception as e:
        return error_response(f"{type(e).__name__}: {e}")


@app.tool()
def get_company_history(company_name: str) -> str:
    """
    특정 기업 전체 투자 히스토리 조회.

    Args:
        company_name: 기업명 (부분 일치. 정확할수록 오매칭 방지)
    """
    try:
        df = load_data()
        matched = df[df["기업명"].str.contains(company_name, na=False, case=False)]

        if matched.empty:
            return error_response(f"'{company_name}'에 해당하는 기업을 찾을 수 없습니다. 기업명을 확인하세요.")

        found_companies = matched["기업명"].unique().tolist()
        warnings = []
        if len(found_companies) > 1:
            warnings.append(f"'{company_name}' 검색에 {len(found_companies)}개 기업 매칭됨. 더 정확한 기업명 사용을 권장합니다.")

        companies_data = {}
        for co in found_companies:
            co_df = matched[matched["기업명"] == co].sort_values(["투자년도", "투자월"])
            companies_data[co] = {
                "총_라운드수": len(co_df),
                "첫_투자년도": int(co_df["투자년도"].min()) if not co_df["투자년도"].isna().all() else None,
                "최신_라운드": co_df.iloc[-1]["당시 투자 Round 정보"] if len(co_df) > 0 else None,
                "최신_누적투자금액(M$)": co_df.iloc[-1]["합계 투자금액(M$)"] if len(co_df) > 0 else None,
                "history": df_to_records(co_df),
            }

        result = {
            "_query_meta": {
                "검색어": company_name,
                "매칭된 기업": found_companies,
                "주의": warnings if warnings else None,
            },
            "companies": companies_data,
        }
        return json.dumps(result, ensure_ascii=False, indent=2)

    except Exception as e:
        return error_response(f"{type(e).__name__}: {e}")


@app.tool()
def get_statistics(
    group_by: str,
    metrics: list[str],
    sector: str | list[str] | None = None,
    region: str | list[str] | None = None,
    investment_type: str | list[str] | None = None,
    year_from: int | None = None,
    year_to: int | None = None,
) -> str:
    """
    포트폴리오 통계/집계. Claude가 직접 계산하지 말고 반드시 이 Tool 사용.

    Args:
        group_by: 집계 기준. 허용값: 분야 | 지역 | 투자유형 | 당시 투자 Round 정보 | Exit 여부
        metrics: 집계 방식 목록 (복수 지정 가능, 영어만, 한국어 금지).
                 허용값:
                   count_companies     = 고유 포트폴리오사 수
                   latest_total_amount = 기업별 최신 누적 투자액 합계 (전체 기간 기준, year 필터 무관)
                   sum_round_amount    = 기간 내 집행된 라운드 금액 합계 (year_from/to 필터와 함께 사용)
                   avg_amount          = 평균 라운드 투자금액
                   median_amount       = 라운드 투자금액 중앙값
                   avg_equity          = 평균 지분율 (전환채권 제외)
                   count_rows          = 투자 건수 (라운드 수)
                   count_exited        = Exit 완료 기업 수
        sector: 섹터 사전 필터 (OR 지원)
        region: 지역 사전 필터 (OR 지원)
        investment_type: 투자유형 사전 필터 (OR 지원). 예: 지분투자 | 전환채권
        year_from: 투자 시작 연도 이상 (포함). sum_round_amount 사용 시 필수
        year_to: 투자 종료 연도 이하 (포함). sum_round_amount 사용 시 필수
    """
    try:
        original_group_by = group_by
        group_by = normalize_group_by(group_by)
        if group_by not in VALID_GROUP_BY:
            return error_response(
                f"group_by 허용값: {', '.join(VALID_GROUP_BY)}. "
                f"입력값 '{group_by}'는 허용되지 않습니다."
            )

        original_metrics = list(metrics)
        normalized_metrics = [normalize_metric(m) for m in metrics]
        invalid = [m for m in normalized_metrics if m not in VALID_METRICS]
        if invalid:
            return error_response(
                f"metrics 허용값: {', '.join(VALID_METRICS)}. "
                f"허용되지 않는 값: {invalid}. 한국어 사용 금지."
            )
        metrics = normalized_metrics

        normalizations = []
        if original_group_by != group_by:
            normalizations.append(f"group_by: '{original_group_by}' → '{group_by}'")
        for orig, norm in zip(original_metrics, normalized_metrics):
            if orig != norm:
                normalizations.append(f"metric: '{orig}' → '{norm}'")

        df_all = load_data()
        applied_filters = {}

        if sector:
            df_all = apply_list_filter(df_all, "분야", sector)
            applied_filters["분야"] = sector
        if region:
            df_all = apply_list_filter(df_all, "지역", region)
            applied_filters["지역"] = region
        if investment_type:
            valid_types = df_all["투자유형"].dropna().unique().tolist()
            types = investment_type if isinstance(investment_type, list) else [investment_type]
            valid_input = [t for t in types if t in valid_types]
            if valid_input:
                df_all = apply_list_filter(df_all, "투자유형", valid_input if len(valid_input) > 1 else valid_input[0])
                applied_filters["투자유형"] = valid_input

        # 기간 필터: sum_round_amount / avg_amount / median_amount / count_rows / count_companies에 적용
        df_period = df_all.copy()
        if year_from is not None:
            df_period = df_period[df_period["투자년도"] >= year_from]
            applied_filters["투자년도 from"] = year_from
        if year_to is not None:
            df_period = df_period[df_period["투자년도"] <= year_to]
            applied_filters["투자년도 to"] = year_to

        if group_by == "Exit 여부":
            df_all = df_all.copy()
            df_all["Exit 여부"] = df_all["Exit 여부"].map({"Y": "Exit 완료", "N": "보유 중"}).fillna(df_all["Exit 여부"])
            df_period = df_period.copy()
            df_period["Exit 여부"] = df_period["Exit 여부"].map({"Y": "Exit 완료", "N": "보유 중"}).fillna(df_period["Exit 여부"])

        # latest_total_amount는 항상 전체 기간 기준 (year 필터 무관)
        latest_all = df_all.sort_values(["투자년도", "투자월"]).groupby("기업명").last().reset_index()

        agg_period = df_period.groupby(group_by)
        agg_all    = df_all.groupby(group_by)
        result_rows: dict = {}
        notes = []

        for metric in metrics:
            if metric == "count_companies":
                # 기간 필터 적용
                for k, v in agg_period["기업명"].nunique().items():
                    result_rows.setdefault(k, {group_by: k})["포트폴리오사 수"] = int(v)

            elif metric == "count_rows":
                for k, v in agg_period.size().items():
                    result_rows.setdefault(k, {group_by: k})["투자 건수"] = int(v)

            elif metric == "latest_total_amount":
                # 전체 기간 기준 — year 필터 무관
                group_map = df_all[["기업명", group_by]].drop_duplicates("기업명")
                merged    = latest_all.merge(group_map, on="기업명", suffixes=("", "_grp"))
                target    = group_by + "_grp" if group_by + "_grp" in merged.columns else group_by
                for k, v in merged.groupby(target)["합계 투자금액(M$)"].sum().round(1).items():
                    result_rows.setdefault(k, {group_by: k})["누적 투자액 합계(M$)"] = float(v)
                if year_from or year_to:
                    notes.append("latest_total_amount는 기업별 전체 기간 누적액 기준이므로 year 필터와 무관합니다.")

            elif metric == "sum_round_amount":
                # 기간 필터 적용된 라운드 금액 합계
                for k, v in agg_period["당시 투자 Round 금액(M$)"].sum().round(1).items():
                    result_rows.setdefault(k, {group_by: k})["기간 집행 금액 합계(M$)"] = float(v)
                if not year_from and not year_to:
                    notes.append("sum_round_amount: year_from/year_to 미설정으로 전체 기간 집계됩니다.")

            elif metric == "avg_amount":
                for k, v in agg_period["당시 투자 Round 금액(M$)"].mean().round(2).items():
                    result_rows.setdefault(k, {group_by: k})["평균 라운드 금액(M$)"] = float(v)

            elif metric == "median_amount":
                for k, v in agg_period["당시 투자 Round 금액(M$)"].median().round(2).items():
                    result_rows.setdefault(k, {group_by: k})["라운드 금액 중앙값(M$)"] = float(v)

            elif metric == "avg_equity":
                df_eq = df_period[df_period["지분율(%)"] > 0]
                for k, v in df_eq.groupby(group_by)["지분율(%)"].mean().round(2).items():
                    result_rows.setdefault(k, {group_by: k})["평균 지분율(%)"] = float(v)
                notes.append("avg_equity: 전환채권(지분율 0%) 제외 후 계산")

            elif metric == "count_exited":
                # Exit은 전체 기간 기준
                exited_df = df_all[df_all["Exit 여부"].isin(["Y", "Exit 완료"])]
                for k, v in exited_df.groupby(group_by)["기업명"].nunique().items():
                    result_rows.setdefault(k, {group_by: k})["Exit 완료 기업 수"] = int(v)

        rows = sorted(
            result_rows.values(),
            key=lambda x: list(x.values())[1] if len(x) > 1 else 0,
            reverse=True,
        )

        result = {
            "_query_meta": {
                "집계 기준": group_by,
                "사용된 metrics": metrics,
                "사전 필터": applied_filters if applied_filters else "없음",
                "정규화": normalizations if normalizations else None,
                "주의": notes if notes else None,
            },
            "rows": rows,
        }
        return json.dumps(result, ensure_ascii=False, indent=2)

    except Exception as e:
        return error_response(f"{type(e).__name__}: {e}")


@app.tool()
def get_portfolio_summary() -> str:
    """
    포트폴리오 전체 현황/요약.
    '포트폴리오 보여줘', '전체 현황', '몇 개사야' 등에 사용.
    """
    try:
        df = load_data()
        companies    = df["기업명"].nunique()
        exited       = df[df["Exit 여부"] == "Y"]["기업명"].nunique()
        latest       = df.sort_values(["투자년도", "투자월"]).groupby("기업명").last()
        total_amount = latest["합계 투자금액(M$)"].sum().round(1)
        first_inv    = df.groupby("기업명")["투자년도"].min()
        yearly_new   = {
            int(k): int(v)
            for k, v in first_inv.value_counts().sort_index().items()
            if pd.notna(k)
        }
        result = {
            "총 포트폴리오사":        int(companies),
            "투자 분야 수":           int(df["분야"].nunique()),
            "투자 지역 수":           int(df["지역"].nunique()),
            "총 투자 건수(라운드)":   int(len(df)),
            "Exit 완료":              int(exited),
            "Exit 미완료(보유 중)":   int(companies - exited),
            "누적 투자금액(M$)":      float(total_amount),
            "분야별 포트폴리오사 수": {k: int(v) for k, v in df.groupby("분야")["기업명"].nunique().sort_values(ascending=False).items()},
            "지역별 포트폴리오사 수": {k: int(v) for k, v in df.groupby("지역")["기업명"].nunique().sort_values(ascending=False).items()},
            "연도별 신규 투자사 수":  yearly_new,
        }
        return json.dumps(result, ensure_ascii=False, indent=2)

    except Exception as e:
        return error_response(f"{type(e).__name__}: {e}")


# ── 진입점 ────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app.run()