#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import json
import traceback
import logging
import ssl
import subprocess
from threading import Thread
import time

# Configure enhanced logging
log_filename = "hwp_mcp_stdio_server_fixed.log"
logging.basicConfig(
    level=logging.INFO,
    filename=log_filename,
    filemode="a",
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    encoding='utf-8'
)

# 추가 스트림 핸들러 설정
stderr_handler = logging.StreamHandler(sys.stderr)
stderr_handler.setFormatter(logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s"))
logger = logging.getLogger("hwp-mcp-stdio-server-fixed")
logger.addHandler(stderr_handler)

# Optional: Disable SSL certificate validation for development
ssl._create_default_https_context = ssl._create_unverified_context

def check_dependencies():
    """필수 의존성을 체크하고 설치 가이드를 제공합니다."""
    logger.info("의존성 체크 시작...")
    
    missing_deps = []
    install_commands = []
    
    # Python 버전 체크
    python_version = sys.version_info
    logger.info(f"Python 버전: {python_version.major}.{python_version.minor}.{python_version.micro}")
    
    if python_version < (3, 8):
        logger.error("Python 3.8 이상이 필요합니다.")
        print("Error: Python 3.8 이상이 필요합니다.", file=sys.stderr)
        return False
    
    # FastMCP 체크
    try:
        import mcp
        logger.info("FastMCP 라이브러리 확인됨")
    except ImportError:
        missing_deps.append("mcp")
        install_commands.append("pip install mcp")
        logger.error("FastMCP 라이브러리가 설치되지 않음")
    
    # pywin32 체크
    try:
        import win32com.client
        logger.info("pywin32 라이브러리 확인됨")
    except ImportError:
        missing_deps.append("pywin32")
        install_commands.append("pip install pywin32>=305")
        install_commands.append("python -m pywin32_postinstall -install")
        logger.error("pywin32 라이브러리가 설치되지 않음")
    
    if missing_deps:
        error_msg = f"""
오류: 필수 라이브러리가 설치되지 않았습니다.

누락된 라이브러리: {', '.join(missing_deps)}

해결 방법:
1. 관리자 권한으로 명령 프롬프트를 엽니다.
2. 다음 명령어들을 순서대로 실행하세요:

{chr(10).join(f'   {cmd}' for cmd in install_commands)}

3. 설치 완료 후 이 서버를 다시 시작하세요.
        """
        logger.error(error_msg)
        print(error_msg, file=sys.stderr)
        return False
    
    return True

def test_hwp_connection():
    """한글 프로그램 연결을 테스트합니다."""
    try:
        import win32com.client
        logger.info("한글 프로그램 연결 테스트 중...")
        
        # 더 안전한 연결 테스트
        try:
            hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
            if hwp:
                logger.info("한글 프로그램 연결 성공")
                # 연결 해제
                hwp = None
                return True
            else:
                logger.error("한글 오브젝트 생성 실패")
                return False
        except Exception as inner_e:
            logger.error(f"한글 프로그램 객체 생성 실패: {str(inner_e)}")
            return False
        
    except Exception as e:
        logger.error(f"한글 프로그램 연결 실패: {str(e)}")
        error_msg = f"""
한글 프로그램 연결 오류:
{str(e)}

해결 방법:
1. 한글 프로그램이 설치되어 있는지 확인하세요.
2. 한글 프로그램이 COM 자동화를 지원하는 버전인지 확인하세요.
3. 한글 프로그램을 한 번 실행한 후 종료해보세요.
4. 관리자 권한으로 실행해보세요.
5. 한글 프로그램 버전이 2014 이상인지 확인하세요.
        """
        print(error_msg, file=sys.stderr)
        return False

# Set up paths
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append(current_dir)

logger.info(f"현재 디렉토리: {current_dir}")
logger.info(f"Python 경로: {sys.path}")

# 의존성 체크
if not check_dependencies():
    sys.exit(1)

# FastMCP import
try:
    from mcp.server.fastmcp import FastMCP
    logger.info("FastMCP 성공적으로 import됨")
except ImportError as e:
    logger.error(f"FastMCP import 실패: {str(e)}")
    print(f"Error: FastMCP import 실패. 'pip install mcp'를 실행하세요.", file=sys.stderr)
    sys.exit(1)

# HwpController import 시도 (수정된 버전 우선)
hwp_controller_imported = False
HwpController = None

try:
    # 먼저 수정된 버전 시도
    from src.tools.hwp_controller_fixed import HwpController
    logger.info("수정된 HwpController import 성공")
    hwp_controller_imported = True
except ImportError as e:
    logger.warning(f"수정된 HwpController import 실패, 원본 시도: {str(e)}")
    try:
        from src.tools.hwp_controller import HwpController
        logger.info("원본 HwpController import 성공")
        hwp_controller_imported = True
    except ImportError as e2:
        logger.error(f"원본 HwpController import도 실패: {str(e2)}")
        # 대체 경로 시도
        try:
            sys.path.append(os.path.join(current_dir, "src"))
            sys.path.append(os.path.join(current_dir, "src", "tools"))
            from hwp_controller_fixed import HwpController
            logger.info("HwpController 대체 경로에서 import 성공 (수정된 버전)")
            hwp_controller_imported = True
        except ImportError as e3:
            try:
                from hwp_controller import HwpController
                logger.info("HwpController 대체 경로에서 import 성공 (원본)")
                hwp_controller_imported = True
            except ImportError as e4:
                logger.error(f"모든 경로에서 HwpController import 실패: {str(e4)}")

# HwpTableTools import 시도
hwp_table_tools_imported = False
try:
    from src.tools.hwp_table_tools import HwpTableTools
    logger.info("HwpTableTools import 성공")
    hwp_table_tools_imported = True
except ImportError as e:
    logger.error(f"HwpTableTools import 실패: {str(e)}")
    try:
        from hwp_table_tools import HwpTableTools
        logger.info("HwpTableTools 대체 경로에서 import 성공")
        hwp_table_tools_imported = True
    except ImportError as e2:
        logger.error(f"모든 경로에서 HwpTableTools import 실패: {str(e2)}")

if not hwp_controller_imported:
    error_msg = """
Error: HwpController 모듈을 찾을 수 없습니다.

해결 방법:
1. src/tools/hwp_controller.py 파일이 존재하는지 확인하세요.
2. 파일 경로와 권한을 확인하세요.
3. 스크립트를 올바른 디렉토리에서 실행하고 있는지 확인하세요.
    """
    logger.error(error_msg)
    print(error_msg, file=sys.stderr)
    sys.exit(1)

if not hwp_table_tools_imported:
    logger.warning("HwpTableTools를 import할 수 없지만 계속 진행합니다.")

# 한글 프로그램 연결 테스트
hwp_connection_ok = test_hwp_connection()
if not hwp_connection_ok:
    logger.warning("한글 프로그램 연결 테스트 실패, 하지만 서버는 시작합니다.")

# Initialize FastMCP server
mcp = FastMCP(
    "hwp-mcp-fixed",
    version="0.2.0",
    description="HWP MCP Server for controlling Hangul Word Processor (Fixed Version)",
    dependencies=["pywin32>=305"],
    env_vars={}
)

# Global HWP controller instance
hwp_controller = None
hwp_table_tools = None

def get_hwp_controller():
    """Get or create HwpController instance with enhanced error handling."""
    global hwp_controller, hwp_table_tools
    if hwp_controller is None:
        logger.info("HwpController 인스턴스 생성 중...")
        try:
            hwp_controller = HwpController()
            # 더 안전한 연결 시도
            if not hwp_controller.connect(visible=True, register_security_module=False):
                logger.error("한글 프로그램 연결 실패")
                hwp_controller = None
                return None
            
            # 테이블 도구 인스턴스도 초기화
            if hwp_table_tools_imported:
                hwp_table_tools = HwpTableTools(hwp_controller)
            
            logger.info("한글 프로그램 연결 성공")
            return hwp_controller
        except Exception as e:
            logger.error(f"HwpController 생성 오류: {str(e)}", exc_info=True)
            hwp_controller = None
            return None
    return hwp_controller

def get_hwp_table_tools():
    """Get or create HwpTableTools instance."""
    global hwp_table_tools, hwp_controller
    if hwp_table_tools is None and hwp_table_tools_imported:
        hwp_controller = get_hwp_controller()
        if hwp_controller:
            hwp_table_tools = HwpTableTools(hwp_controller)
    return hwp_table_tools

def reset_hwp_controller():
    """Reset HwpController instance."""
    global hwp_controller, hwp_table_tools
    if hwp_controller:
        try:
            hwp_controller.disconnect()
        except:
            pass
    hwp_controller = None
    hwp_table_tools = None
    logger.info("HwpController 인스턴스 리셋됨")

@mcp.tool()
def hwp_system_info() -> str:
    """시스템 정보와 의존성 상태를 확인합니다."""
    try:
        info = {
            "python_version": f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}",
            "hwp_controller_imported": hwp_controller_imported,
            "hwp_table_tools_imported": hwp_table_tools_imported,
            "current_directory": current_dir,
            "log_file": log_filename,
            "server_version": "0.2.0 (Fixed)"
        }
        
        # 한글 프로그램 상태 확인
        try:
            hwp = get_hwp_controller()
            info["hwp_connection"] = "성공" if hwp else "실패"
        except Exception as e:
            info["hwp_connection"] = f"오류: {str(e)}"
        
        return json.dumps(info, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.error(f"시스템 정보 확인 오류: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_ping_pong(message: str = "핑") -> str:
    """
    연결 테스트용 핑퐁 함수입니다.
    
    Args:
        message: 테스트 메시지 (기본값: "핑")
        
    Returns:
        str: 응답 메시지
    """
    try:
        logger.info(f"핑퐁 테스트 함수 호출됨: 메시지 - {message}")
        
        if message == "핑":
            response = "퐁"
        elif message == "퐁":
            response = "핑"
        else:
            response = f"알 수 없는 메시지: {message} (핑 또는 퐁을 보내주세요)"
        
        current_time = time.strftime("%Y-%m-%d %H:%M:%S")
        
        result = {
            "response": response,
            "original_message": message,
            "timestamp": current_time,
            "server_status": "정상 작동 (Fixed Version)"
        }
        
        return json.dumps(result, ensure_ascii=False)
    except Exception as e:
        logger.error(f"핑퐁 테스트 함수 오류: {str(e)}", exc_info=True)
        return f"테스트 오류 발생: {str(e)}"

@mcp.tool()
def hwp_create() -> str:
    """Create a new HWP document with enhanced error handling."""
    try:
        logger.info("새 문서 생성 시도...")
        
        # HwpController 인스턴스 가져오기
        hwp = get_hwp_controller()
        if not hwp:
            logger.error("한글 프로그램 연결 실패")
            # 다시 시도
            reset_hwp_controller()
            hwp = get_hwp_controller()
            if not hwp:
                return "Error: 한글 프로그램에 연결할 수 없습니다. 한글 프로그램이 설치되어 있고 실행 가능한지 확인하세요."
        
        # 새 문서 생성
        if hwp.create_new_document():
            logger.info("새 문서 생성 성공")
            return "새 문서가 성공적으로 생성되었습니다"
        else:
            logger.error("새 문서 생성 실패")
            # 연결을 다시 시도
            reset_hwp_controller()
            return "Error: 새 문서 생성에 실패했습니다. 한글 프로그램 상태를 확인하세요."
    except Exception as e:
        logger.error(f"문서 생성 오류: {str(e)}", exc_info=True)
        # 오류 발생시 연결 리셋
        reset_hwp_controller()
        return f"Error: 문서 생성 중 오류가 발생했습니다: {str(e)}"

@mcp.tool()
def hwp_insert_text(text: str, preserve_linebreaks: bool = True) -> str:
    """Insert text with enhanced error handling."""
    try:
        if not text:
            return "Error: 텍스트가 필요합니다"
        
        logger.info(f"텍스트 삽입 시도: {text[:50]}...")
        
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: 한글 프로그램에 연결할 수 없습니다"
        
        if hwp.insert_text(text, preserve_linebreaks):
            logger.info("텍스트 삽입 성공")
            return "텍스트가 성공적으로 삽입되었습니다"
        else:
            logger.error("텍스트 삽입 실패")
            return "Error: 텍스트 삽입에 실패했습니다"
    except Exception as e:
        logger.error(f"텍스트 삽입 오류: {str(e)}", exc_info=True)
        return f"Error: 텍스트 삽입 중 오류가 발생했습니다: {str(e)}"

@mcp.tool()
def hwp_save_document(file_path: str = None) -> str:
    """Save the current document."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: 한글 프로그램에 연결할 수 없습니다"
        
        if hwp.save_document(file_path):
            if file_path:
                logger.info(f"문서 저장 성공: {file_path}")
                return f"문서가 성공적으로 저장되었습니다: {file_path}"
            else:
                logger.info("문서 저장 성공")
                return "문서가 성공적으로 저장되었습니다"
        else:
            return "Error: 문서 저장에 실패했습니다"
    except Exception as e:
        logger.error(f"문서 저장 오류: {str(e)}", exc_info=True)
        return f"Error: 문서 저장 중 오류가 발생했습니다: {str(e)}"

@mcp.tool()
def hwp_insert_table(rows: int, cols: int) -> str:
    """Insert a table into the document."""
    try:
        if rows <= 0 or cols <= 0:
            return "Error: 행과 열은 1 이상이어야 합니다"
        
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: 한글 프로그램에 연결할 수 없습니다"
        
        if hwp.insert_table(rows, cols):
            logger.info(f"표 삽입 성공: {rows}x{cols}")
            return f"표가 성공적으로 삽입되었습니다 ({rows}행 x {cols}열)"
        else:
            return "Error: 표 삽입에 실패했습니다"
    except Exception as e:
        logger.error(f"표 삽입 오류: {str(e)}", exc_info=True)
        return f"Error: 표 삽입 중 오류가 발생했습니다: {str(e)}"

@mcp.tool()
def hwp_reset_connection() -> str:
    """Reset HWP connection."""
    try:
        reset_hwp_controller()
        return "한글 프로그램 연결이 리셋되었습니다"
    except Exception as e:
        logger.error(f"연결 리셋 오류: {str(e)}", exc_info=True)
        return f"Error: 연결 리셋 중 오류가 발생했습니다: {str(e)}"

if __name__ == "__main__":
    logger.info("수정된 HWP MCP stdio 서버 시작")
    logger.info(f"로그 파일: {log_filename}")
    logger.info(f"한글 프로그램 연결 테스트: {'성공' if hwp_connection_ok else '실패'}")
    
    try:
        # FastMCP 서버를 stdio transport로 실행
        logger.info("서버 실행 중...")
        mcp.run(transport="stdio")
    except KeyboardInterrupt:
        logger.info("서버가 사용자에 의해 중단되었습니다")
        reset_hwp_controller()
    except Exception as e:
        logger.error(f"서버 실행 오류: {str(e)}", exc_info=True)
        print(f"Error: {str(e)}", file=sys.stderr)
        reset_hwp_controller()
        sys.exit(1)
