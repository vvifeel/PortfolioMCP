# Portfolio MCP 서버 — Windows 셋업 가이드

## 폴더 구조
```
C:\Users\YourName\portfolio_mcp\
├── portfolio_server.py          ← MCP 서버 코드
├── requirements.txt             ← 패키지 목록
├── sample_portfolio.xlsx        ← 포트폴리오 엑셀 (여기에 복사)
└── claude_desktop_config_example.json
```

---

## Step 1 — 파일 배치

1. 이 폴더 전체를 `C:\Users\(본인계정)\portfolio_mcp\` 에 복사
2. 포트폴리오 엑셀 파일을 같은 폴더에 복사하고 이름을 `sample_portfolio.xlsx` 로 변경
   (또는 config의 PORTFOLIO_EXCEL_PATH 경로를 실제 파일 경로로 수정)

---

## Step 2 — 패키지 설치

PowerShell 또는 명령 프롬프트에서:

```powershell
cd C:\Users\YourName\portfolio_mcp
pip install -r requirements.txt
```

설치 확인:
```powershell
python -c "import mcp; import pandas; print('OK')"
```

---

## Step 3 — Claude Desktop config 수정

Claude Desktop config 파일 위치:
```
C:\Users\(본인계정)\AppData\Roaming\Claude\claude_desktop_config.json
```

파일이 없으면 새로 만들고, 있으면 mcpServers 항목을 추가:

```json
{
  "mcpServers": {
    "portfolio": {
      "command": "python",
      "args": ["C:\\Users\\YourName\\portfolio_mcp\\portfolio_server.py"],
      "env": {
        "PORTFOLIO_EXCEL_PATH": "C:\\Users\\YourName\\portfolio_mcp\\sample_portfolio.xlsx"
      }
    }
  }
}
```

⚠️ 경로에서 역슬래시(\)는 반드시 두 번(\\) 써야 합니다
⚠️ YourName 을 본인 Windows 계정명으로 교체하세요

---

## Step 4 — Claude Desktop 재시작

config 수정 후 Claude Desktop을 완전히 종료했다가 다시 시작합니다.
(트레이 아이콘 우클릭 → 종료 → 재시작)

---

## Step 5 — 연결 확인

Claude Desktop에서 새 대화를 시작하고 입력창 근처에
🔧 아이콘(도구) 또는 MCP 연결 표시가 보이면 성공입니다.

테스트 질문:
- "포트폴리오 전체 요약 보여줘"
- "한국 Bio 섹터 포트폴리오 검색해줘"
- "오션벤처스 투자 히스토리 보여줘"

---

## 문제 해결

**"python을 찾을 수 없다" 오류**
→ `command`를 `"python"` 대신 전체 경로로 변경:
```json
"command": "C:\\Python311\\python.exe"
```
Python 경로 확인: `where python` (PowerShell에서)

**MCP 아이콘이 안 보인다**
→ config 파일 경로 오타 확인
→ JSON 문법 오류 확인 (역슬래시 두 번인지)
→ Claude Desktop 완전 재시작

**엑셀 파일 못 찾는다는 오류**
→ PORTFOLIO_EXCEL_PATH 경로 확인
→ 파일명 대소문자 확인

---

## 엑셀 파일 업데이트 시

엑셀 파일만 교체하면 됩니다.
MCP 서버는 매 쿼리마다 파일을 새로 읽으므로
Claude Desktop 재시작 없이 바로 반영됩니다.

---

## 사용 가능한 Tool 목록

| Tool | 설명 |
|------|------|
| search_portfolio | 다중 조건 필터링 검색 |
| get_company_history | 특정 기업 전체 라운드 히스토리 |
| get_statistics | 분야/지역별 집계 통계 |
| get_portfolio_summary | 전체 포트폴리오 현황 요약 |
