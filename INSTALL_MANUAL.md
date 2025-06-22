# HWP MCP 수정된 버전 - 수동 설치 가이드

## 1단계: 필수 패키지 설치

명령 프롬프트를 열고 다음 명령어를 실행하세요:

```cmd
pip install mcp pywin32>=305
python -m pywin32_postinstall -install
```

## 2단계: 설치 확인

```cmd
python -c "import mcp; print('MCP OK')"
python -c "import win32com.client; print('pywin32 OK')"
```

## 3단계: HWP 연결 테스트

```cmd
python -c "import win32com.client; hwp = win32com.client.Dispatch('HWPFrame.HwpObject'); print('HWP 연결 성공'); hwp = None"
```

## 4단계: HWP MCP 테스트

```cmd
cd C:\Users\USER\Downloads\hwp-mcp-main\hwp-mcp-main
python test_hwp_mcp_fixed.py
```

## 5단계: Claude Desktop 설정

Claude Desktop 설정 파일을 편집하세요:

**파일 위치**: `%APPDATA%\Claude\claude_desktop_config.json`

**추가할 내용**:
```json
{
  "mcpServers": {
    "hwp-mcp-fixed": {
      "command": "python",
      "args": ["C:\\Users\\USER\\Downloads\\hwp-mcp-main\\hwp-mcp-main\\hwp_mcp_stdio_server_fixed.py"],
      "cwd": "C:\\Users\\USER\\Downloads\\hwp-mcp-main\\hwp-mcp-main"
    }
  }
}
```

## 문제해결

### pywin32 설치 문제
```cmd
pip uninstall pywin32
pip install pywin32>=305
python -m pywin32_postinstall -install
```

### HWP 연결 문제
- 한글 프로그램이 설치되어 있는지 확인
- 한글 프로그램 버전이 2014 이상인지 확인
- 관리자 권한으로 명령 프롬프트 실행

### MCP 설치 문제
```cmd
pip install --upgrade pip
pip install mcp
```

---

이 방법으로 설치가 완료되면 Claude에서 HWP MCP 도구들을 사용할 수 있습니다.
