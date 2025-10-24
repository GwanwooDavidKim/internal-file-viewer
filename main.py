#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
사내 파일 뷰어 (Internal File Viewer)

메인 애플리케이션 진입점입니다.

이 프로그램은 팀 내에 흩어져 있는 다양한 포맷의 업무 자료(PPT, PDF, Excel 등)를
하나의 애플리케이션에서 신속하게 탐색하고 내용을 확인할 수 있도록 도와줍니다.
"""
import sys
import os
import config
from core.auth import AuthenticationManager
from utils.logger import LoggerManager, get_logger


def console_login(auth_manager):
    """
    콘솔 모드에서 로그인을 수행합니다.
    
    Args:
        auth_manager: AuthenticationManager 인스턴스
        
    Returns:
        bool: 로그인 성공 여부
    """
    logger = get_logger()
    
    # 비대화형 환경 감지
    if not sys.stdin.isatty():
        logger.warning("비대화형 환경에서 실행 중 - 콘솔 로그인 불가")
        print("[오류] 비대화형 환경에서는 GUI 모드를 사용하세요: python main.py --gui")
        return False
    
    print(f"\n=== {config.APP_SETTINGS['app_name']} v{config.APP_SETTINGS['app_version']} ===")
    print("콘솔 모드로 실행 중입니다.")
    print("\n데모용 계정:")
    print("• 관리자 계정과 팀원 계정을 사용할 수 있습니다")
    print("• 실제 운영 시에는 설정에서 계정을 관리하세요")
    print("-" * 50)
    
    max_attempts = 3
    for attempt in range(max_attempts):
        print(f"\n로그인 시도 {attempt + 1}/{max_attempts}")
        
        try:
            username = input("사용자명: ").strip()
            password = input("비밀번호: ").strip()
            
            if not username or not password:
                print("[오류] 사용자명과 비밀번호를 모두 입력해주세요.")
                continue
            
            success, message = auth_manager.authenticate(username, password)
            
            if success:
                print(f"[완료] {message}")
                return True
            else:
                print(f"[오류] {message}")
                
        except KeyboardInterrupt:
            print("\n\n프로그램을 종료합니다.")
            return False
        except Exception as e:
            print(f"[오류] 로그인 중 오류 발생: {e}")
    
    print("\n[오류] 로그인 시도 횟수를 초과했습니다.")
    return False


def console_menu(auth_manager):
    """
    콘솔 메뉴를 표시하고 사용자 입력을 처리합니다.
    
    Args:
        auth_manager: AuthenticationManager 인스턴스
    """
    while True:
        user_info = auth_manager.get_user_info()
        if not user_info:
            break
        
        print("\n" + "=" * 60)
        print(f"사용자: {user_info['username']}")
        if user_info['is_admin']:
            print("권한: 관리자")
        else:
            remaining_days = user_info.get('remaining_days', 0)
            print(f"권한: 일반 사용자 (남은 일수: {remaining_days}일)")
        
        print("\n[정보] 메뉴:")
        print("1. 파일 탐색 (개발 중)")
        print("2. 파일 검색 (개발 중)")
        print("3. 사용자 정보 보기")
        if user_info['is_admin']:
            print("4. 관리자 메뉴 (개발 중)")
        print("9. 로그아웃")
        print("0. 종료")
        print("-" * 60)
        
        try:
            choice = input("선택하세요: ").strip()
            
            if choice == "1":
                print("\n[폴더] 파일 탐색 기능은 개발 중입니다.")
                print("지원 예정 형식: PDF, PPT/PPTX, Excel, Word, 이미지")
                
            elif choice == "2":
                print("\n[확인] 파일 검색 기능은 개발 중입니다.")
                print("파일명 및 내용 검색 기능을 제공할 예정입니다.")
                
            elif choice == "3":
                show_user_info(user_info)
                
            elif choice == "4" and user_info['is_admin']:
                print("\n[사용자] 관리자 메뉴는 개발 중입니다.")
                print("사용자 계정 관리 기능을 제공할 예정입니다.")
                
            elif choice == "9":
                auth_manager.logout()
                print("[완료] 로그아웃되었습니다.")
                break
                
            elif choice == "0":
                auth_manager.logout()
                print("[완료] 프로그램을 종료합니다.")
                break
                
            else:
                print("[오류] 올바른 메뉴 번호를 선택해주세요.")
                
        except KeyboardInterrupt:
            print("\n\n[완료] 프로그램을 종료합니다.")
            auth_manager.logout()
            break
        except Exception as e:
            print(f"[오류] 오류 발생: {e}")


def show_user_info(user_info):
    """
    사용자 정보를 표시합니다.
    
    Args:
        user_info: 사용자 정보 딕셔너리
    """
    print("\n" + "=" * 40)
    print("[사용자] 사용자 정보")
    print("=" * 40)
    print(f"사용자명: {user_info['username']}")
    print(f"유형: {'관리자' if user_info['is_admin'] else '일반 사용자'}")
    print(f"로그인 시간: {user_info['login_time'].strftime('%Y-%m-%d %H:%M:%S')}")
    
    if not user_info['is_admin']:
        expiration_date = user_info.get('expiration_date')
        if expiration_date:
            print(f"계정 만료일: {expiration_date.strftime('%Y-%m-%d')}")
            print(f"남은 사용일: {user_info.get('remaining_days', 0)}일")
        else:
            print("계정 만료일: 설정되지 않음")
    
    print("=" * 40)


def setup_application(gui_mode=False):
    """
    애플리케이션 초기 설정을 수행합니다.
    """
    print(f"[시작] {config.APP_SETTINGS['app_name']} 시작 중...")
    print(f"[정보] 버전: {config.APP_SETTINGS['app_version']}")
    if gui_mode:
        print("[모드] GUI 모드로 실행됩니다.")
    else:
        print("[설정] 콘솔 모드로 실행됩니다.")
    return True


def check_dependencies():
    """
    필수 의존성을 확인합니다.
    
    Returns:
        bool: 모든 의존성이 충족되면 True
    """
    required_modules = [
        'pandas',
        'openpyxl', 
        'fitz',  # PyMuPDF
        'pptx',  # python-pptx
        'docx',  # python-docx
        'PIL',   # Pillow
    ]
    
    missing_modules = []
    print("[확인] 의존성 확인 중...")
    
    for module in required_modules:
        try:
            if module == 'fitz':
                import fitz
                print("  [완료] PyMuPDF (PDF 처리)")
            elif module == 'pptx':
                import pptx
                print("  [완료] python-pptx (PowerPoint 처리)")
            elif module == 'docx':
                import docx
                print("  [완료] python-docx (Word 처리)")
            elif module == 'PIL':
                import PIL
                print("  [완료] Pillow (이미지 처리)")
            elif module == 'pandas':
                import pandas
                print("  [완료] pandas (Excel 처리)")
            elif module == 'openpyxl':
                import openpyxl
                print("  [완료] openpyxl (Excel 처리)")
            else:
                __import__(module)
        except ImportError:
            missing_modules.append(module)
            print(f"  [오류] {module} - 설치되지 않음")
    
    if missing_modules:
        print(f"\n[오류] 다음 모듈이 설치되지 않았습니다: {', '.join(missing_modules)}")
        print("pip install pandas openpyxl PyMuPDF python-pptx python-docx Pillow 명령어로 설치해주세요.")
        return False
    
    print("[완료] 모든 의존성이 확인되었습니다.")
    return True


def launch_gui():
    """
    GUI 모드로 애플리케이션을 실행합니다.
    """
    try:
        from PyQt6.QtWidgets import QApplication
        from ui.main_window import MainWindow
        
        app = QApplication(sys.argv)
        
        # 메인 윈도우 생성 및 표시
        window = MainWindow()
        window.show()
        
        # 이벤트 루프 실행
        sys.exit(app.exec())
        
    except ImportError as e:
        print(f"[오류] PyQt6를 사용할 수 없습니다: {e}")
        print("GUI 모드를 사용하려면 PyQt6를 설치해주세요: pip install PyQt6")
        sys.exit(1)
    except Exception as e:
        print(f"[오류] GUI 모드 실행 중 오류가 발생했습니다: {str(e)}", file=sys.stderr)
        sys.exit(1)


def main():
    """
    메인 함수 - 애플리케이션의 진입점입니다.
    """
    try:
        # 로깅 시스템 초기화
        logger_manager = LoggerManager()
        logger = logger_manager.get_app_logger()
        logger.info("애플리케이션 시작")
        
        # 의존성 확인
        if not check_dependencies():
            logger.error("의존성 확인 실패")
            sys.exit(1)
        
        # 실행 모드 확인 (GUI 또는 콘솔)
        if len(sys.argv) > 1 and sys.argv[1] == "--gui":
            # GUI 모드
            setup_application(gui_mode=True)
            launch_gui()
        else:
            # 콘솔 모드 (기본)
            setup_application(gui_mode=False)
            auth_manager = AuthenticationManager()
            
            # 로그인 수행
            if console_login(auth_manager):
                # 메인 메뉴 실행
                console_menu(auth_manager)
            else:
                print("[오류] 로그인에 실패했습니다.")
                sys.exit(1)
        
    except Exception as e:
        print(f"[오류] 애플리케이션 실행 중 오류가 발생했습니다: {str(e)}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()