# HWP MCP - 수정된 버전 (Enhanced & Fixed)

한글 워드프로세서(HWP)를 Claude와 함께 사용할 수 있게 해주는 MCP(Model Context Protocol) 서버입니다.

## 🚀 주요 개선사항

### ✨ 안정성 향상
- **연결 리셋 기능**: 문제 발생 시 자동으로 한글 프로그램 연결을 재설정
- **향상된 오류 처리**: 더 상세한 오류 메시지와 해결 방법 제공
- **안전한 리소스 관리**: 프로그램 종료 시 리소스 정리

### 🔧 기능 완성
- **누락된 메서드 추가**: `fill_table_cell`, `get_table_cell_text`, `merge_table_cells` 등
- **문서 생성 안정화**: 여러 방법을 시도하여 문서 생성 성공률 향상
- **텍스트 삽입 개선**: 다양한 텍스트 형식 지원

### 📊 디버깅 지원
- **상세한 로깅**: 문제 진단을 위한 자세한 로그 기록
- **테스트 스크립트**: 설치 및 연결 상태를 쉽게 확인
- **시스템 정보**: 환경 설정 확인 도구

## 📁 파일 구조

```
hwp-mcp-main/
├── src/tools/
│   ├── hwp_controller_fixed.py      # 수정된 HWP 컨트롤러 (메인)
│   ├── hwp_controller.py            # 원본 HWP 컨트롤러
│   └── hwp_table_tools.py           # 표 관련 도구
├── hwp_mcp_stdio_server_fixed.py   # 수정된 MCP 서버 (메인)
├── hwp_mcp_stdio_server.py         # 원본 MCP 서버
├── install_hwp_mcp_fixed.bat       # 수정된 설치 스크립트
├── test_hwp_mcp_fixed.py           # 테스트 스크립트
├── claude_desktop_config_example.json  # Claude Desktop 설정 예시
├── HWP_MCP_수정버전_가이드.md       # 상세 가이드
└── README_FIXED.md                  # 이 파일
```

## 🛠️ 설치 방법

### 1. 자동 설치 (권장)

```cmd
cd C:\Users\USER\Downloads\hwp-mcp-main\hwp-mcp-main\
install_hwp_mcp_fixed.bat
```

### 2. 수동 설치

```cmd
# 필수 패키지 설치
pip install mcp pywin32>=305
python -m pywin32_postinstall -install

# 설치 확인
python test_hwp_mcp_fixed.py
```

## ⚙️ Claude Desktop 설정

Claude Desktop의 설정 파일에 다음을 추가하세요:

**위치**: `%APPDATA%\Claude\claude_desktop_config.json`

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

## 🎯 사용 가능한 도구들

### 기본 문서 조작
- `hwp_create()` - 새 문서 생성
- `hwp_insert_text(text)` - 텍스트 삽입
- `hwp_save_document(path)` - 문서 저장

### 표 관련 기능
- `hwp_insert_table(rows, cols)` - 표 삽입
- 표 데이터 채우기 (JSON 형식 지원)
- 셀 병합 및 개별 셀 편집

### 시스템 관리
- `hwp_system_info()` - 시스템 정보 확인
- `hwp_reset_connection()` - 연결 리셋
- `hwp_ping_pong()` - 연결 테스트

## 🔍 문제해결

### 자주 발생하는 문제

#### 1. "새 문서 생성에 실패했습니다"
```
해결방법: hwp_reset_connection() 실행 후 다시 시도
```

#### 2. "한글 프로그램에 연결할 수 없습니다"
```
확인사항:
- 한글 프로그램 설치 여부 (2014 이상 권장)
- COM 자동화 지원 버전인지 확인
- 관리자 권한으로 실행
```

#### 3. "pywin32 관련 오류"
```cmd
pip uninstall pywin32
pip install pywin32>=305
python -m pywin32_postinstall -install
```

### 디버깅

1. **로그 확인**: `hwp_mcp_stdio_server_fixed.log`
2. **테스트 실행**: `python test_hwp_mcp_fixed.py`
3. **시스템 정보**: Claude에서 `hwp_system_info()` 실행

## 📋 요구사항

- **Python**: 3.8 이상
- **한글 프로그램**: 2014 이상 (COM 자동화 지원)
- **운영체제**: Windows
- **라이브러리**: mcp, pywin32>=305

## 🆚 원본 버전과의 차이점

| 항목 | 원본 버전 | 수정된 버전 |
|------|-----------|-------------|
| 연결 안정성 | 기본 | 향상된 재연결 로직 |
| 오류 처리 | 최소한 | 상세한 오류 메시지 |
| 누락 메서드 | 일부 누락 | 모든 메서드 구현 |
| 디버깅 | 기본 로그 | 상세한 로깅 |
| 테스트 | 없음 | 종합 테스트 스크립트 |
| 설치 | 수동 | 자동 설치 스크립트 |

## 📝 사용 예시

```python
# Claude에서 다음과 같이 사용할 수 있습니다:

# 1. 새 문서 생성
hwp_create()

# 2. 텍스트 삽입
hwp_insert_text("안녕하세요, HWP MCP 수정된 버전입니다!")

# 3. 표 삽입
hwp_insert_table(3, 4)

# 4. 문서 저장
hwp_save_document("C:/Documents/test.hwp")

# 5. 문제 발생 시 연결 리셋
hwp_reset_connection()
```

## 🤝 기여

문제 발견이나 개선 제안이 있으시면 이슈를 등록해주세요.

## 📄 라이선스

MIT License

---

**주의**: 이 수정된 버전은 원본 HWP MCP의 모든 기능을 포함하며, 안정성과 사용성을 크게 개선했습니다. 원본 버전에서 문제가 있었다면 이 수정된 버전을 사용해보세요.
