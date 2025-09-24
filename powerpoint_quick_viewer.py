#!/usr/bin/env python3
"""
PowerPoint 빠른 접근 뷰어
딜레이 없이 PowerPoint 파일에 즉시 접근할 수 있는 유틸리티
"""

import os
import subprocess
import sys
from pathlib import Path

def open_powerpoint_file(file_path: str):
    """PowerPoint 파일을 기본 프로그램으로 엽니다."""
    try:
        if sys.platform == 'win32':
            # Windows: PowerPoint로 직접 열기
            os.startfile(file_path)
        elif sys.platform == 'darwin':
            # macOS: 기본 앱으로 열기
            subprocess.run(['open', file_path])
        else:
            # Linux: 기본 앱으로 열기
            subprocess.run(['xdg-open', file_path])
        return True
    except Exception as e:
        print(f"파일 열기 오류: {e}")
        return False

def open_file_location(file_path: str):
    """파일이 있는 폴더를 엽니다."""
    try:
        folder_path = os.path.dirname(file_path)
        if sys.platform == 'win32':
            # Windows Explorer에서 파일 선택
            subprocess.run(['explorer', '/select,', file_path])
        elif sys.platform == 'darwin':
            # Finder에서 파일 선택
            subprocess.run(['open', '-R', file_path])
        else:
            # 폴더 열기
            subprocess.run(['xdg-open', folder_path])
        return True
    except Exception as e:
        print(f"폴더 열기 오류: {e}")
        return False

def copy_file_path(file_path: str):
    """파일 경로를 클립보드에 복사합니다."""
    try:
        if sys.platform == 'win32':
            # Windows 클립보드
            import pyperclip
            pyperclip.copy(file_path)
        return True
    except Exception as e:
        print(f"경로 복사 오류: {e}")
        return False

if __name__ == "__main__":
    # 테스트용
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        if os.path.exists(file_path):
            print(f"🚀 PowerPoint 파일 열기: {file_path}")
            open_powerpoint_file(file_path)
        else:
            print(f"❌ 파일을 찾을 수 없습니다: {file_path}")
    else:
        print("사용법: python powerpoint_quick_viewer.py <PowerPoint 파일 경로>")