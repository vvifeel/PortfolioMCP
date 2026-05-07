"""
Portfolio MCP Server v9
- explain_query Tool 추가: 검색 전 해석 출력을 별도 Tool로 강제
- build_interpretation_header, format_filters 제거 (explain_query로 통합)
- 모든 Tool에서 query_interpretation, mapping_notes 파라미터 제거
- 모든 Tool의 header 생성 코드 제거
- Tool description 필수 출력 순서 지시 → explain_query 호출 지시로 교체
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

app = FastMCP("portfolio-mcp")

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
        "sectors":   sorted(df["분야"].dropna().unique().tolist()),
        "regions":   sorted(df["지역"].dropna().unique().tolist()),
        "rounds":    sorted(df["당시 투자 Round 정보"].dropna().unique().tolist()),
        "companies": sorted(df["기업명"].dropna().unique().tolist()),
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


# ── Tool description 동적 생성 ────────────────────────────────────────
def _search_desc() -> str:
    try:
        meta = get_db_meta()
        sectors_str = ", ".join(meta["sectors"])
        regions_str = ", ".join(meta["regions"])
        rounds_str  = ", ".join(meta["rounds"])
    except Exception:
        sectors_str = regions_str = rounds_str = "(엑셀 로드 실패)"

    return f"""【필수】포트폴리오 투자 데이터 검색.
Claude 자체 지식으로 절대 답변하지 마세요.

【필수 호출 순서】
1. 반드시 explain_query를 먼저 호출하세요.
2. explain_query 반환값 전체를 사용자에게 그대로 출력하세요. 절대 생략하지 마세요.
3. 출력 완료 후 이 Tool을 호출하세요.

【파라미터 매핑 규칙】
사용자 표현을 아래 DB 실제값으로 반드시 변환하세요.
복수 선택은 배열로 전달 (OR 조건). 다른 필드 간은 AND 조건.

▸ sector DB 실제값: {sectors_str}
▸ region DB 실제값: {regions_str}
▸ round_info DB 실제값: {rounds_str}

매핑 예시:
  '바이오' → 'Bio/Healthcare'
  '인공지능' → 'AI/SaaS'
  '미국이랑 한국' → ["미국", "한국"] (OR)
  'Series B 이상' → ["Series B", "Series C", "Series D", "Pre-IPO CB"] (OR)
  '기업별 최신 현황' → latest_only=True
  '투자금액 큰 순' → sort_by="합계 투자금액(M$)", sort_asc=False

【시각화 규칙】
- 결과 3건 이상이면 반드시 표로 출력
- 차트는 반드시 표 다음에 출력
- 단일 값이면 차트 생략"""


def _stats_desc() -> str:
    try:
        meta = get_db_meta()
        sectors_str = ", ".join(meta["sectors"])
        regions_str = ", ".join(meta["regions"])
    except Exception:
        sectors_str = regions_str = "(엑셀 로드 실패)"

    return f"""【필수】포트폴리오 통계/집계. Claude가 직접 계산하지 마세요.
복수 metric 동시 집계 지원.

【필수 호출 순서】
1. 반드시 explain_query를 먼저 호출하세요.
2. explain_query 반환값 전체를 사용자에게 그대로 출력하세요. 절대 생략하지 마세요.
3. 출력 완료 후 이 Tool을 호출하세요.

▸ sector 필터 DB 실제값: {sectors_str}
▸ region 필터 DB 실제값: {regions_str}

【시각화 규칙】
- 집계 결과는 반드시 표로 출력
- 비교 의미 있을 때만 차트 추가 (표 다음에)"""


# ══════════════════════════════════════════════════════════════════════
# TOOLS
# ══════════════════════════════════════════════════════════════════════

# ── Tool 0: explain_query ─────────────────────────────────────────────
@app.tool()
def explain_query(
    interpretation: str,
    applied_filters: dict,
    mapping_notes: list[str] | None = None,
    condition_rule: str = "같은 필드 내 복수값 = OR  /  다른 필드 간 = AND",
) -> str:
    """
    【필수】search_portfolio, get_statistics, get_company_history 호출 전에
    반드시 이 Tool을 먼저 호출하세요. DB 연산 없이 즉시 반환합니다.

    【중요】이 Tool의 반환값을 절대 수정하거나 요약하지 마세요.
    반환된 텍스트 전체를 사용자에게 그대로 출력하세요.
    출력을 생략하거나 건너뛰고 검색 Tool을 호출하지 마세요.

    Args:
        interpretation: 사용자 질문을 어떻게 이해했는지 한 줄 설명
        applied_filters: 적용할 필터 딕셔너리
                         예: {"분야": ["Bio/Healthcare"], "지역": "한국"}
        mapping_notes: 사용자 표현 → DB 실제값 변환 내용 목록
                       예: ["'바이오' → 'Bio/Healthcare'", "'한국이랑 미국' → ['한국', '미국']"]
        condition_rule: 조건 적용 규칙 설명
    """
    lines = [
        "━" * 40,
        "[검색 해석]",
        f"• 이해한 내용: {interpretation}",
    ]

    if mapping_notes:
        lines.append("• DB 매핑:")
        for note in mapping_notes:
            lines.append(f"    {note}")

    if applied_filters:
        lines.append("• 적용할 조건:")
        for k, v in applied_filters.items():
            if isinstance(v, list):
                lines.append(f"    - {k}: {' 또는 '.join(str(i) for i in v)}")
            else:
                lines.append(f"    - {k}: {v}")
    else:
        lines.append("• 적용할 조건: 없음 (전체 조회)")

    lines.append(f"• 조건 규칙: {condition_rule}")
    lines.append("━" * 40)

    return "\n".join(lines)


# ── Tool 1: search_portfolio ──────────────────────────────────────────
@app.tool(description=_search_desc())
def search_portfolio(
    company_name: str | None = None,
    sector: str | list[str] | None = None,
    region: str | list[str] | None = None,
    investment_type: str | None = None,
    exit_status: str | None = None,
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
    Args:
        company_name: 기업명 부분 일치 검색
        sector: 투자 분야. OR 지원 (배열 가능)
        region: 투자 지역. OR 지원 (배열 가능)
        investment_type: 지분투자 또는 전환채권
        exit_status: Y=Exit완료, N=미완료
        round_info: 투자 라운드. OR 지원 (배열 가능)
        year_from: 투자 시작 연도 이후 (포함)
        year_to: 투자 종료 연도 이전 (포함)
        min_total_amount: 누적 투자금액 최솟값 (M$)
        max_total_amount: 누적 투자금액 최댓값 (M$)
        min_equity: 지분율 최솟값 (%)
        max_equity: 지분율 최댓값 (%)
        latest_only: True면 기업별 최신 라운드 행만 반환
        sort_by: 정렬 기준 컬럼
        sort_asc: True=오름차순, False=내림차순
        limit: 최대 반환 행 수 (기본 200)
    """
    try:
        df = load_data()

        if company_name:
            df = df[df["기업명"].str.contains(company_name, na=False, case=False)]
        if sector:
            df = apply_list_filter(df, "분야", sector)
        if region:
            df = apply_list_filter(df, "지역", region)
        if investment_type:
            df = df[df["투자유형"] == investment_type]
        if exit_status:
            df = df[df["Exit 여부"] == exit_status]
        if round_info:
            if isinstance(round_info, list):
                df = df[df["당시 투자 Round 정보"].isin(round_info)]
            else:
                df = df[df["당시 투자 Round 정보"].str.contains(round_info, na=False)]
        if year_from is not None:
            df = df[df["투자년도"] >= year_from]
        if year_to is not None:
            df = df[df["투자년도"] <= year_to]
        if min_total_amount is not None:
            df = df[df["합계 투자금액(M$)"] >= min_total_amount]
        if max_total_amount is not None:
            df = df[df["합계 투자금액(M$)"] <= max_total_amount]
        if min_equity is not None:
            df = df[df["지분율(%)"] >= min_equity]
        if max_equity is not None:
            df = df[df["지분율(%)"] <= max_equity]
        if latest_only:
            df = df.sort_values(["투자년도", "투자월"]).groupby("기업명").last().reset_index()
        if sort_by in df.columns:
            df = df.sort_values(sort_by, ascending=sort_asc)

        total = len(df)
        truncated = total > limit
        df = df.head(limit)

        result = {
            "total_matched": total,
            "returned": len(df),
            "truncated": truncated,
            "truncated_note": f"결과 {total}건 중 {limit}건만 반환. 조건을 좁히거나 limit을 조정하세요." if truncated else None,
            "records": df_to_records(df),
        }
        return json.dumps(result, ensure_ascii=False, indent=2)

    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {e}"}, ensure_ascii=False)


# ── Tool 2: get_company_history ───────────────────────────────────────
@app.tool()
def get_company_history(company_name: str) -> str:
    """
    【필수】특정 기업의 투자 히스토리 조회.
    Claude 자체 지식으로 특정 기업 데이터를 절대 답변하지 마세요.
    복수 기업 매칭 시 기업별로 분리하여 반환합니다.

    【필수 호출 순서】
    1. 반드시 explain_query를 먼저 호출하세요.
    2. explain_query 반환값 전체를 사용자에게 그대로 출력하세요. 절대 생략하지 마세요.
    3. 출력 완료 후 이 Tool을 호출하세요.

    Args:
        company_name: 기업명 (부분 일치. 정확할수록 오매칭 방지)
    """
    try:
        df = load_data()
        matched = df[df["기업명"].str.contains(company_name, na=False, case=False)]

        if matched.empty:
            return json.dumps({
                "error": f"'{company_name}'에 해당하는 기업을 찾을 수 없습니다."
            }, ensure_ascii=False)

        found_companies = matched["기업명"].unique().tolist()
        warning = None
        if len(found_companies) > 1:
            warning = f"'{company_name}' 검색에 {len(found_companies)}개 기업 매칭: {found_companies}. 더 정확한 기업명으로 재검색을 권장합니다."

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
            "매칭된_기업수": len(found_companies),
            "매칭된_기업": found_companies,
            "warning": warning,
            "companies": companies_data,
        }
        return json.dumps(result, ensure_ascii=False, indent=2)

    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {e}"}, ensure_ascii=False)


# ── Tool 3: get_statistics ────────────────────────────────────────────
@app.tool(description=_stats_desc())
def get_statistics(
    group_by: str,
    metrics: list[str],
    sector: str | list[str] | None = None,
    region: str | list[str] | None = None,
) -> str:
    """
    Args:
        group_by: 집계 기준. 분야 | 지역 | 투자유형 | 당시 투자 Round 정보 | Exit 여부
        metrics: 집계 방식 목록 (복수 동시 지정 가능).
                 count_companies=포트폴리오사 수,
                 sum_amount=누적 투자금액 합계,
                 avg_amount=평균 라운드 투자금액,
                 avg_equity=평균 지분율(전환채권 제외),
                 count_rows=투자 건수
        sector: 섹터 사전 필터 (OR 지원)
        region: 지역 사전 필터 (OR 지원)
    """
    try:
        df = load_data()

        if sector:
            df = apply_list_filter(df, "분야", sector)
        if region:
            df = apply_list_filter(df, "지역", region)
        if group_by == "Exit 여부":
            df = df.copy()
            df["Exit 여부"] = df["Exit 여부"].map({"Y": "Exit 완료", "N": "보유 중"}).fillna(df["Exit 여부"])

        latest   = df.sort_values(["투자년도", "투자월"]).groupby("기업명").last().reset_index()
        agg_base = df.groupby(group_by)
        result_rows: dict = {}

        for metric in metrics:
            if metric == "count_companies":
                for k, v in agg_base["기업명"].nunique().items():
                    result_rows.setdefault(k, {group_by: k})["포트폴리오사 수"] = int(v)
            elif metric == "count_rows":
                for k, v in agg_base.size().items():
                    result_rows.setdefault(k, {group_by: k})["투자 건수"] = int(v)
            elif metric == "sum_amount":
                group_map = df[["기업명", group_by]].drop_duplicates("기업명")
                merged    = latest.merge(group_map, on="기업명", suffixes=("", "_grp"))
                target    = group_by + "_grp" if group_by + "_grp" in merged.columns else group_by
                for k, v in merged.groupby(target)["합계 투자금액(M$)"].sum().round(1).items():
                    result_rows.setdefault(k, {group_by: k})["총 투자금액(M$)"] = float(v)
            elif metric == "avg_amount":
                for k, v in agg_base["당시 투자 Round 금액(M$)"].mean().round(2).items():
                    result_rows.setdefault(k, {group_by: k})["평균 라운드 투자금액(M$)"] = float(v)
            elif metric == "avg_equity":
                df_eq = df[df["지분율(%)"] > 0]
                for k, v in df_eq.groupby(group_by)["지분율(%)"].mean().round(2).items():
                    result_rows.setdefault(k, {group_by: k})["평균 지분율(%)"] = float(v)

        rows = sorted(
            result_rows.values(),
            key=lambda x: list(x.values())[1] if len(x) > 1 else 0,
            reverse=True,
        )
        result = {
            "주의사항": "avg_equity는 전환채권(지분율 0%) 제외 후 계산된 값입니다." if "avg_equity" in metrics else None,
            "rows": rows,
        }
        return json.dumps(result, ensure_ascii=False, indent=2)

    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {e}"}, ensure_ascii=False)


# ── Tool 4: get_portfolio_summary ─────────────────────────────────────
@app.tool()
def get_portfolio_summary() -> str:
    """
    【필수】포트폴리오 전체 현황/요약.
    '포트폴리오 보여줘', '전체 현황', '몇 개사야' 등 모두 포함.
    Claude 자체 지식으로 절대 답변하지 마세요.

    【시각화 규칙】
    - 분야별/지역별 분포는 반드시 표로 출력
    - 비교 의미 있을 때만 차트 추가 (표 다음에)
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
        return json.dumps({"error": f"{type(e).__name__}: {e}"}, ensure_ascii=False)


# ── Tool 5: get_db_schema ─────────────────────────────────────────────
@app.tool()
def get_db_schema() -> str:
    """
    포트폴리오 DB의 현재 고유값 목록 반환.
    엑셀 업데이트 후 최신 값 확인 시 사용.
    어떤 섹터/지역/라운드가 있는지 물어볼 때 사용.
    """
    try:
        meta = get_db_meta()
        result = {
            "description":      "포트폴리오 DB 현재 고유값 목록 (엑셀 원본 기준 실시간 반영)",
            "sectors":          meta["sectors"],
            "regions":          meta["regions"],
            "rounds":           meta["rounds"],
            "total_companies":  len(meta["companies"]),
            "sample_companies": meta["companies"][:10],
        }
        return json.dumps(result, ensure_ascii=False, indent=2)

    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {e}"}, ensure_ascii=False)


# ── 진입점 ────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app.run()