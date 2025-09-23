# -*- coding: utf-8 -*-
"""
사내 파일 뷰어 설정 파일 (Internal File Viewer Configuration)

사용자 계정 정보와 애플리케이션 전반의 설정값을 관리합니다.
"""
from datetime import datetime, timedelta

# 관리자 계정 정보 (Admin Accounts)
ADMIN_ACCOUNTS = {
    "admin": "admin123",  # 관리자 계정
}

# 배포용 팀원 계정 정보 (Deployable Team Member Accounts)
DEPLOYABLE_ACCOUNTS = {
    "user1": "password1",
    "user2": "password2",
    "user3": "password3",
}

# 계정 만료일 설정 (Account Expiration Dates)
# 관리자는 이 설정을 수정하여 팀원 계정의 사용 기한을 관리할 수 있습니다.
ACCOUNT_EXPIRATION = {
    "user1": datetime(2025, 12, 31),  # 2025년 12월 31일까지
    "user2": datetime(2025, 11, 30),  # 2025년 11월 30일까지
    "user3": datetime(2025, 10, 31),  # 2025년 10월 31일까지
}

# 애플리케이션 설정 (Application Settings)
APP_SETTINGS = {
    "app_name": "사내 파일 뷰어",
    "app_version": "1.0.0",
    "default_folder_path": "",  # 기본 탐색 폴더 경로 (빈 문자열이면 사용자가 선택)
    "window_width": 1200,
    "window_height": 800,
    "splitter_ratio": 0.3,  # 좌측 파일 탐색기 : 우측 뷰어 비율
}

# 파일 지원 형식 (Supported File Formats)
SUPPORTED_FILE_EXTENSIONS = {
    "pdf": [".pdf"],
    "powerpoint": [".ppt", ".pptx"],
    "excel": [".xls", ".xlsx", ".xlsm"],
    "word": [".doc", ".docx"],
    "image": [".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".svg"],
    "text": [".txt", ".md", ".log"],
}

# 검색 설정 (Search Settings)
SEARCH_SETTINGS = {
    "max_file_size_mb": 50,  # 검색 대상 최대 파일 크기 (MB)
    "index_update_interval": 300,  # 인덱스 업데이트 주기 (초)
    "search_timeout": 3,  # 검색 타임아웃 (초)
}

# UI 색상 팔레트 (Color Palette)
# 디자인 원칙 문서에서 정의된 색상을 사용
UI_COLORS = {
    "primary": "#34495E",     # 주조색 - 짙은 회색/남색 계열
    "secondary": "#BDC3C7",   # 보조색 - 밝은 회색
    "accent": "#3498DB",      # 강조색 - 파란색 계열
    "text": "#2C3E50",        # 텍스트 색상
    "background": "#FFFFFF",  # 배경색
    "hover": "#5DADE2",       # 마우스 오버 시 색상
}

# 폰트 설정 (Font Settings)
UI_FONTS = {
    "title_size": 16,      # 제목 폰트 크기
    "subtitle_size": 12,   # 부제목 폰트 크기
    "body_size": 10,       # 본문 폰트 크기
    "font_family": "맑은 고딕",  # 기본 폰트
}


def get_all_supported_extensions():
    """
    모든 지원되는 파일 확장자 목록을 반환합니다.
    
    Returns:
        list: 지원되는 모든 파일 확장자 리스트
    """
    extensions = []
    for ext_list in SUPPORTED_FILE_EXTENSIONS.values():
        extensions.extend(ext_list)
    return extensions


def is_account_expired(username):
    """
    사용자 계정의 만료 여부를 확인합니다.
    
    Args:
        username (str): 확인할 사용자명
        
    Returns:
        bool: 만료되었으면 True, 아니면 False
    """
    if username in ADMIN_ACCOUNTS:
        return False  # 관리자 계정은 만료되지 않음
    
    if username not in ACCOUNT_EXPIRATION:
        return True  # 만료일이 설정되지 않은 계정은 만료로 처리
    
    return datetime.now() > ACCOUNT_EXPIRATION[username]


def get_remaining_days(username):
    """
    사용자 계정의 남은 사용 일수를 반환합니다.
    
    Args:
        username (str): 확인할 사용자명
        
    Returns:
        int: 남은 사용 일수 (만료된 경우 음수)
    """
    if username in ADMIN_ACCOUNTS:
        return float('inf')  # 관리자는 무제한
    
    if username not in ACCOUNT_EXPIRATION:
        return -1  # 만료일 미설정
    
    remaining = ACCOUNT_EXPIRATION[username] - datetime.now()
    return remaining.days