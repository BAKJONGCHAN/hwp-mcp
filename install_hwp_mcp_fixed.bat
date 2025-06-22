@echo off
chcp 65001 > nul
echo ========================================
echo HWP MCP 수정된 버전 설치 스크립트
echo ========================================
echo.

echo [1/5] Python 및 pip 확인 중...
python --version
if %errorlevel% neq 0 (
    echo 오류: Python이 설치되지 않았거나 PATH에 추가되지 않았습니다.
    echo Python 3.8 이상을 설치하고 PATH에 추가해주세요.
    pause
    exit /b 1
)

echo.
echo [2/5] 필수 패키지 설치 중...
echo MCP 라이브러리 설치...
pip install mcp
if %errorlevel% neq 0 (
    echo 오류: MCP 설치 실패
    pause
    exit /b 1
)

echo pywin32 라이브러리 설치...
pip install pywin32>=305
if %errorlevel% neq 0 (
    echo 오류: pywin32 설치 실패
    pause
    exit /b 1
)

echo.
echo [3/5] pywin32 후처리 설정...
python -m pywin32_postinstall -install
if %errorlevel% neq 0 (
    echo 경고: pywin32 후처리 설정 실패 (무시하고 계속)
)

echo.
echo [4/5] 한글 프로그램 설치 확인...
echo 한글 프로그램 COM 인터페이스 테스트 중...
python -c "import win32com.client; hwp = win32com.client.Dispatch('HWPFrame.HwpObject'); print('한글 프로그램 연결 성공'); hwp = None"
if %errorlevel% neq 0 (
    echo 경고: 한글 프로그램 연결 테스트 실패
    echo 다음을 확인해주세요:
    echo - 한글 프로그램이 설치되어 있는지
    echo - 한글 프로그램 버전이 2014 이상인지
    echo - COM 자동화가 지원되는 버전인지
    echo.
    echo 계속 진행하려면 아무 키나 누르세요...
    pause > nul
)

echo.
echo [5/5] 수정된 HWP MCP 서버 테스트...
echo 서버 시작 테스트 중 (5초 후 자동 종료)...
timeout /t 2 /nobreak > nul
python hwp_mcp_stdio_server_fixed.py --test 2>&1 | findstr /C:"정상 작동"
if %errorlevel% neq 0 (
    echo 경고: 서버 테스트에서 일부 문제가 발생했습니다.
    echo 하지만 설치는 완료되었습니다.
)

echo.
echo ========================================
echo 수정된 HWP MCP 설치 완료!
echo ========================================
echo.
echo 사용 방법:
echo 1. Claude Desktop의 MCP 설정에서 이 서버를 추가하세요
echo 2. 서버 경로: %CD%\hwp_mcp_stdio_server_fixed.py
echo 3. 로그 파일: %CD%\hwp_mcp_stdio_server_fixed.log
echo.
echo 주요 개선사항:
echo - 안정된 한글 프로그램 연결
echo - 향상된 오류 처리
echo - 연결 리셋 기능
echo - 상세한 로깅
echo.
echo 설치가 완료되었습니다!
pause
