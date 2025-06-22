# HWP MCP (한글 문서 처리 MCP 서버)

한글 HWP 문서 작성을 위한 Model Context Protocol (MCP) 서버입니다.

## 기능

- 한글(HWP) 문서 생성 및 편집
- 텍스트 삽입 및 서식 설정
- 테이블 생성 및 관리
- Claude Desktop과 통합하여 자연어로 문서 작성

## 설치 방법

### 1. 저장소 클론

```bash
git clone https://github.com/BAKJONGCHAN/hwp-mcp.git
cd hwp-mcp
```

### 2. 의존성 설치

```bash
pip install -r requirements.txt
```

### 3. Claude Desktop 설정

Claude Desktop의 설정 파일을 수정합니다:

**Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
**macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "hwp-mcp": {
      "command": "python",
      "args": ["C:\\Users\\[사용자명]\\Documents\\hwp-mcp\\hwp_mcp_stdio_server_fixed.py"],
      "cwd": "C:\\Users\\[사용자명]\\Documents\\hwp-mcp"
    }
  }
}
```

> **주의**: `[사용자명]` 부분을 실제 사용자명으로 변경하고, 클론한 경로에 맞게 수정하세요.

### 4. 가상환경 사용 (권장)

```bash
# 가상환경 생성
python -m venv venv

# 가상환경 활성화 (Windows)
venv\Scripts\activate

# 가상환경 활성화 (macOS/Linux)
source venv/bin/activate

# 의존성 설치
pip install -r requirements.txt
```

가상환경을 사용하는 경우 Claude Desktop 설정:

```json
{
  "mcpServers": {
    "hwp-mcp": {
      "command": "C:\\Users\\[사용자명]\\Documents\\hwp-mcp\\venv\\Scripts\\python.exe",
      "args": ["hwp_mcp_stdio_server_fixed.py"],
      "cwd": "C:\\Users\\[사용자명]\\Documents\\hwp-mcp"
    }
  }
}
```

## 사용 방법

1. 한글(HWP) 프로그램이 설치되어 있어야 합니다
2. Claude Desktop을 재시작합니다
3. Claude와 대화하면서 HWP 문서를 작성할 수 있습니다

### 예시 명령어

- "새로운 HWP 문서를 만들어 주세요"
- "제목을 '보고서'로 입력해 주세요"
- "3x3 표를 만들어 주세요"
- "문서를 저장해 주세요"

## 파일 구조

- `hwp_mcp_stdio_server_fixed.py`: 메인 MCP 서버 파일
- `requirements.txt`: Python 의존성 패키지 목록
- `INSTALL_MANUAL.md`: 상세 설치 가이드
- `install_hwp_mcp_fixed.bat`: Windows 자동 설치 스크립트

## 문제 해결

상세한 문제 해결 방법은 [INSTALL_MANUAL.md](INSTALL_MANUAL.md) 파일을 참조하세요.

## 라이선스

MIT License

## 기여

이슈나 풀 리퀘스트는 언제든 환영합니다!
