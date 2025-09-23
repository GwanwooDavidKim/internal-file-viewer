# -*- coding: utf-8 -*-
"""
인증 및 계정 관리 모듈 (Authentication and Account Management Module)

사용자 로그인, 계정 유효성 검사, 권한 관리 등의 기능을 제공합니다.
"""
from datetime import datetime
from typing import Optional, Tuple, Dict, Any
import bcrypt
import config


class AuthenticationManager:
    """
    사용자 인증 및 계정 관리를 담당하는 클래스입니다.
    
    이 클래스는 다음과 같은 기능을 제공합니다:
    - 사용자 로그인 검증
    - 계정 만료 여부 확인
    - 관리자 권한 확인
    - 사용자 세션 관리
    """
    
    def __init__(self):
        """AuthenticationManager 인스턴스를 초기화합니다."""
        self.current_user = None
        self.is_admin = False
        self.login_time = None
        
    def authenticate(self, username: str, password: str) -> Tuple[bool, str]:
        """
        사용자 인증을 수행합니다.
        
        Args:
            username (str): 사용자명
            password (str): 비밀번호
            
        Returns:
            Tuple[bool, str]: (인증 성공 여부, 메시지)
        """
        try:
            # 입력값 유효성 검사
            if not username or not password:
                return False, "사용자명과 비밀번호를 모두 입력해주세요."
            
            # 관리자 계정 확인
            if username in config.ADMIN_ACCOUNTS:
                if self._verify_password(password, config.ADMIN_ACCOUNTS[username]):
                    self._set_user_session(username, is_admin=True)
                    return True, f"관리자로 로그인되었습니다. 환영합니다, {username}님!"
                else:
                    return False, "비밀번호가 올바르지 않습니다."
            
            # 일반 사용자 계정 확인
            if username in config.DEPLOYABLE_ACCOUNTS:
                if self._verify_password(password, config.DEPLOYABLE_ACCOUNTS[username]):
                    # 계정 만료 여부 확인
                    if config.is_account_expired(username):
                        remaining_days = config.get_remaining_days(username)
                        return False, f"계정이 만료되었습니다. (만료: {abs(remaining_days)}일 전)"
                    
                    self._set_user_session(username, is_admin=False)
                    remaining_days = config.get_remaining_days(username)
                    return True, f"로그인 성공! 환영합니다, {username}님! (남은 사용일: {remaining_days}일)"
                else:
                    return False, "비밀번호가 올바르지 않습니다."
            
            # 존재하지 않는 사용자
            return False, "존재하지 않는 사용자입니다."
            
        except Exception as e:
            return False, f"로그인 중 오류가 발생했습니다: {str(e)}"
    
    def _verify_password(self, password: str, hashed_password: str) -> bool:
        """
        비밀번호를 해시와 비교하여 검증합니다.
        
        Args:
            password (str): 평문 비밀번호
            hashed_password (str): 해시된 비밀번호
            
        Returns:
            bool: 비밀번호 일치 여부
        """
        try:
            return bcrypt.checkpw(password.encode('utf-8'), hashed_password.encode('utf-8'))
        except Exception:
            return False
    
    def _set_user_session(self, username: str, is_admin: bool = False) -> None:
        """
        사용자 세션을 설정합니다.
        
        Args:
            username (str): 사용자명
            is_admin (bool): 관리자 여부
        """
        self.current_user = username
        self.is_admin = is_admin
        self.login_time = datetime.now()
    
    def logout(self) -> None:
        """현재 사용자를 로그아웃시킵니다."""
        self.current_user = None
        self.is_admin = False
        self.login_time = None
    
    def is_logged_in(self) -> bool:
        """
        사용자가 로그인되어 있는지 확인합니다.
        
        Returns:
            bool: 로그인 상태 여부
        """
        return self.current_user is not None
    
    def check_session_validity(self) -> Tuple[bool, str]:
        """
        현재 세션의 유효성을 확인합니다.
        
        Returns:
            Tuple[bool, str]: (세션 유효 여부, 메시지)
        """
        if not self.is_logged_in():
            return False, "로그인이 필요합니다."
        
        # 관리자는 항상 유효
        if self.is_admin:
            return True, "세션이 유효합니다."
        
        # 일반 사용자의 계정 만료 여부 재확인
        if config.is_account_expired(self.current_user):
            self.logout()
            return False, "계정이 만료되어 자동 로그아웃되었습니다."
        
        return True, "세션이 유효합니다."
    
    def get_user_info(self) -> Optional[Dict[str, Any]]:
        """
        현재 로그인된 사용자의 정보를 반환합니다.
        
        Returns:
            Optional[Dict[str, Any]]: 사용자 정보 딕셔너리 또는 None
        """
        if not self.is_logged_in():
            return None
        
        user_info = {
            "username": self.current_user,
            "is_admin": self.is_admin,
            "login_time": self.login_time,
        }
        
        if not self.is_admin and self.current_user:
            user_info.update({
                "expiration_date": config.ACCOUNT_EXPIRATION.get(self.current_user),
                "remaining_days": config.get_remaining_days(self.current_user),
                "is_expired": config.is_account_expired(self.current_user),
            })
        
        return user_info
    
    def has_admin_permission(self) -> bool:
        """
        현재 사용자가 관리자 권한을 가지고 있는지 확인합니다.
        
        Returns:
            bool: 관리자 권한 여부
        """
        return self.is_admin and self.is_logged_in()
    
    def get_all_users_info(self) -> Dict[str, Dict[str, Any]]:
        """
        모든 사용자의 정보를 반환합니다. (관리자 전용)
        
        Returns:
            Dict[str, Dict[str, Any]]: 모든 사용자 정보
        """
        if not self.has_admin_permission():
            return {}
        
        users_info = {}
        
        # 관리자 계정 정보
        for admin_username in config.ADMIN_ACCOUNTS.keys():
            users_info[admin_username] = {
                "type": "admin",
                "is_expired": False,
                "expiration_date": None,
                "remaining_days": float('inf'),
            }
        
        # 일반 사용자 계정 정보
        for username in config.DEPLOYABLE_ACCOUNTS.keys():
            users_info[username] = {
                "type": "user",
                "is_expired": config.is_account_expired(username),
                "expiration_date": config.ACCOUNT_EXPIRATION.get(username),
                "remaining_days": config.get_remaining_days(username),
            }
        
        return users_info
    
    def update_user_expiration(self, username: str, new_expiration: datetime) -> Tuple[bool, str]:
        """
        사용자 계정의 만료일을 업데이트합니다. (관리자 전용)
        
        Args:
            username (str): 대상 사용자명
            new_expiration (datetime): 새로운 만료일
            
        Returns:
            Tuple[bool, str]: (성공 여부, 메시지)
        """
        if not self.has_admin_permission():
            return False, "관리자 권한이 필요합니다."
        
        if username not in config.DEPLOYABLE_ACCOUNTS:
            return False, "존재하지 않는 사용자입니다."
        
        try:
            # 설정 업데이트 (실제 구현에서는 파일이나 데이터베이스에 저장)
            config.ACCOUNT_EXPIRATION[username] = new_expiration
            return True, f"{username}의 계정 만료일이 {new_expiration.strftime('%Y-%m-%d')}로 업데이트되었습니다."
        except Exception as e:
            return False, f"계정 만료일 업데이트 중 오류가 발생했습니다: {str(e)}"