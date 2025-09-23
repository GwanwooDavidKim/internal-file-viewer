# -*- coding: utf-8 -*-
"""
로깅 시스템 (Logging System)

애플리케이션 전체의 로깅을 관리합니다.
"""
import logging
import os
import sys
import time
from datetime import datetime
from logging.handlers import RotatingFileHandler
from typing import Optional
import config


class ApplicationLogger:
    """
    애플리케이션 로거 클래스입니다.
    
    파일과 콘솔 출력을 지원하며, 로그 파일 로테이션 기능을 제공합니다.
    """
    
    def __init__(self, name: str = "FileViewer", log_dir: str = "logs"):
        """
        ApplicationLogger 인스턴스를 초기화합니다.
        
        Args:
            name (str): 로거 이름
            log_dir (str): 로그 파일 디렉토리
        """
        self.name = name
        self.log_dir = log_dir
        self.logger = None
        self.setup_logger()
    
    def setup_logger(self):
        """로거를 설정합니다."""
        # 로그 디렉토리 생성
        if not os.path.exists(self.log_dir):
            os.makedirs(self.log_dir)
        
        # 로거 생성
        self.logger = logging.getLogger(self.name)
        self.logger.setLevel(logging.DEBUG)
        
        # 기존 핸들러 제거 (중복 방지)
        if self.logger.handlers:
            self.logger.handlers.clear()
        
        # 포매터 설정
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # 파일 핸들러 설정 (로테이션)
        log_file = os.path.join(self.log_dir, f"{self.name.lower()}.log")
        file_handler = RotatingFileHandler(
            log_file, 
            maxBytes=10*1024*1024,  # 10MB
            backupCount=5
        )
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(formatter)
        self.logger.addHandler(file_handler)
        
        # 콘솔 핸들러 설정
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.INFO)
        
        # 콘솔용 간단한 포매터
        console_formatter = logging.Formatter(
            '%(levelname)s - %(message)s'
        )
        console_handler.setFormatter(console_formatter)
        self.logger.addHandler(console_handler)
        
        # 에러 로그 별도 파일
        error_file = os.path.join(self.log_dir, f"{self.name.lower()}_errors.log")
        error_handler = RotatingFileHandler(
            error_file,
            maxBytes=5*1024*1024,  # 5MB
            backupCount=3
        )
        error_handler.setLevel(logging.ERROR)
        error_handler.setFormatter(formatter)
        self.logger.addHandler(error_handler)
    
    def debug(self, message: str, **kwargs):
        """디버그 로그를 기록합니다."""
        self.logger.debug(message, **kwargs)
    
    def info(self, message: str, **kwargs):
        """정보 로그를 기록합니다."""
        self.logger.info(message, **kwargs)
    
    def warning(self, message: str, **kwargs):
        """경고 로그를 기록합니다."""
        self.logger.warning(message, **kwargs)
    
    def error(self, message: str, exception: Optional[Exception] = None, **kwargs):
        """오류 로그를 기록합니다."""
        if exception:
            self.logger.error(f"{message}: {str(exception)}", exc_info=True, **kwargs)
        else:
            self.logger.error(message, **kwargs)
    
    def critical(self, message: str, exception: Optional[Exception] = None, **kwargs):
        """심각한 오류 로그를 기록합니다."""
        if exception:
            self.logger.critical(f"{message}: {str(exception)}", exc_info=True, **kwargs)
        else:
            self.logger.critical(message, **kwargs)


class PerformanceLogger:
    """
    성능 측정을 위한 로거 클래스입니다.
    """
    
    def __init__(self, logger: ApplicationLogger):
        """
        PerformanceLogger 인스턴스를 초기화합니다.
        
        Args:
            logger (ApplicationLogger): 기본 로거
        """
        self.logger = logger
        self.start_times = {}
    
    def start_timer(self, operation_name: str):
        """
        성능 측정을 시작합니다.
        
        Args:
            operation_name (str): 작업 이름
        """
        self.start_times[operation_name] = time.time()
        self.logger.debug(f"성능 측정 시작: {operation_name}")
    
    def end_timer(self, operation_name: str, log_level: str = "info"):
        """
        성능 측정을 종료하고 결과를 로깅합니다.
        
        Args:
            operation_name (str): 작업 이름
            log_level (str): 로그 레벨
        """
        if operation_name not in self.start_times:
            self.logger.warning(f"성능 측정이 시작되지 않은 작업: {operation_name}")
            return
        
        elapsed_time = time.time() - self.start_times[operation_name]
        del self.start_times[operation_name]
        
        message = f"성능 측정 완료: {operation_name} - {elapsed_time:.3f}초"
        
        if log_level == "debug":
            self.logger.debug(message)
        elif log_level == "info":
            self.logger.info(message)
        elif log_level == "warning":
            self.logger.warning(message)
    
    def measure_function(self, func_name: str = None):
        """
        함수 실행 시간을 측정하는 데코레이터입니다.
        
        Args:
            func_name (str): 함수 이름 (지정하지 않으면 실제 함수명 사용)
        """
        def decorator(func):
            def wrapper(*args, **kwargs):
                operation_name = func_name or f"{func.__module__}.{func.__name__}"
                self.start_timer(operation_name)
                
                try:
                    result = func(*args, **kwargs)
                    self.end_timer(operation_name)
                    return result
                except Exception as e:
                    self.end_timer(operation_name)
                    self.logger.error(f"함수 실행 중 오류 발생: {operation_name}", exception=e)
                    raise
            
            return wrapper
        return decorator


class FileOperationLogger:
    """
    파일 작업 전용 로거 클래스입니다.
    """
    
    def __init__(self, logger: ApplicationLogger):
        """
        FileOperationLogger 인스턴스를 초기화합니다.
        
        Args:
            logger (ApplicationLogger): 기본 로거
        """
        self.logger = logger
    
    def log_file_access(self, file_path: str, operation: str, success: bool = True, 
                       error: Optional[Exception] = None):
        """
        파일 접근 로그를 기록합니다.
        
        Args:
            file_path (str): 파일 경로
            operation (str): 작업 종류 (read, write, delete 등)
            success (bool): 성공 여부
            error (Optional[Exception]): 오류 정보
        """
        filename = os.path.basename(file_path)
        
        if success:
            self.logger.info(f"파일 {operation} 성공: {filename}")
        else:
            self.logger.error(f"파일 {operation} 실패: {filename}", exception=error)
    
    def log_file_processing(self, file_path: str, file_type: str, 
                          processing_time: float, success: bool = True,
                          error: Optional[Exception] = None):
        """
        파일 처리 로그를 기록합니다.
        
        Args:
            file_path (str): 파일 경로
            file_type (str): 파일 타입
            processing_time (float): 처리 시간 (초)
            success (bool): 성공 여부
            error (Optional[Exception]): 오류 정보
        """
        filename = os.path.basename(file_path)
        
        if success:
            self.logger.info(f"파일 처리 완료: {filename} ({file_type}, {processing_time:.3f}초)")
        else:
            self.logger.error(f"파일 처리 실패: {filename} ({file_type})", exception=error)
    
    def log_search_operation(self, query: str, result_count: int, search_time: float):
        """
        검색 작업 로그를 기록합니다.
        
        Args:
            query (str): 검색 쿼리
            result_count (int): 결과 수
            search_time (float): 검색 시간 (초)
        """
        self.logger.info(f"검색 완료: '{query}' - {result_count}개 결과, {search_time:.3f}초")
    
    def log_indexing_operation(self, directory: str, file_count: int, 
                             indexing_time: float, success: bool = True):
        """
        인덱싱 작업 로그를 기록합니다.
        
        Args:
            directory (str): 인덱싱 디렉토리
            file_count (int): 처리된 파일 수
            indexing_time (float): 인덱싱 시간 (초)
            success (bool): 성공 여부
        """
        if success:
            self.logger.info(f"인덱싱 완료: {directory} - {file_count}개 파일, {indexing_time:.3f}초")
        else:
            self.logger.error(f"인덱싱 실패: {directory}")


class AuthenticationLogger:
    """
    인증 관련 로거 클래스입니다.
    """
    
    def __init__(self, logger: ApplicationLogger):
        """
        AuthenticationLogger 인스턴스를 초기화합니다.
        
        Args:
            logger (ApplicationLogger): 기본 로거
        """
        self.logger = logger
    
    def log_login_attempt(self, username: str, success: bool, ip_address: str = "local"):
        """
        로그인 시도를 기록합니다.
        
        Args:
            username (str): 사용자명
            success (bool): 성공 여부
            ip_address (str): IP 주소
        """
        if success:
            self.logger.info(f"로그인 성공: {username} ({ip_address})")
        else:
            self.logger.warning(f"로그인 실패: {username} ({ip_address})")
    
    def log_logout(self, username: str):
        """
        로그아웃을 기록합니다.
        
        Args:
            username (str): 사용자명
        """
        self.logger.info(f"로그아웃: {username}")
    
    def log_session_expired(self, username: str):
        """
        세션 만료를 기록합니다.
        
        Args:
            username (str): 사용자명
        """
        self.logger.info(f"세션 만료: {username}")
    
    def log_permission_denied(self, username: str, operation: str):
        """
        권한 거부를 기록합니다.
        
        Args:
            username (str): 사용자명
            operation (str): 시도한 작업
        """
        self.logger.warning(f"권한 거부: {username} - {operation}")


class LoggerManager:
    """
    전체 로깅 시스템을 관리하는 클래스입니다.
    """
    
    _instance = None
    _initialized = False
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self):
        if not self._initialized:
            self.app_logger = ApplicationLogger()
            self.performance_logger = PerformanceLogger(self.app_logger)
            self.file_logger = FileOperationLogger(self.app_logger)
            self.auth_logger = AuthenticationLogger(self.app_logger)
            
            self._initialized = True
            self.app_logger.info("로깅 시스템 초기화 완료")
    
    @classmethod
    def get_instance(cls):
        """싱글톤 인스턴스를 반환합니다."""
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance
    
    def get_app_logger(self) -> ApplicationLogger:
        """애플리케이션 로거를 반환합니다."""
        return self.app_logger
    
    def get_performance_logger(self) -> PerformanceLogger:
        """성능 로거를 반환합니다."""
        return self.performance_logger
    
    def get_file_logger(self) -> FileOperationLogger:
        """파일 작업 로거를 반환합니다."""
        return self.file_logger
    
    def get_auth_logger(self) -> AuthenticationLogger:
        """인증 로거를 반환합니다."""
        return self.auth_logger
    
    def shutdown(self):
        """로깅 시스템을 종료합니다."""
        self.app_logger.info("로깅 시스템 종료")
        # 모든 핸들러 정리
        for handler in self.app_logger.logger.handlers:
            handler.close()
        
        logging.shutdown()


# 편의 함수들
def get_logger() -> ApplicationLogger:
    """전역 애플리케이션 로거를 반환합니다."""
    return LoggerManager.get_instance().get_app_logger()


def get_performance_logger() -> PerformanceLogger:
    """전역 성능 로거를 반환합니다."""
    return LoggerManager.get_instance().get_performance_logger()


def get_file_logger() -> FileOperationLogger:
    """전역 파일 로거를 반환합니다."""
    return LoggerManager.get_instance().get_file_logger()


def get_auth_logger() -> AuthenticationLogger:
    """전역 인증 로거를 반환합니다."""
    return LoggerManager.get_instance().get_auth_logger()


def measure_performance(operation_name: str = None):
    """성능 측정 데코레이터입니다."""
    return get_performance_logger().measure_function(operation_name)