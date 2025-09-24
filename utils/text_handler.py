# -*- coding: utf-8 -*-
"""
텍스트 파일 처리 모듈 (Text File Handler)

텍스트 파일(.txt, .md, .log)의 내용 읽기 및 미리보기 기능을 제공합니다.
"""
import os
from typing import Dict, Any, Optional


class TextHandler:
    """
    텍스트 파일 처리를 위한 클래스입니다.
    
    주요 기능:
    - 텍스트 파일 내용 읽기
    - 파일 정보 및 메타데이터 조회
    - 미리보기 텍스트 제공
    """
    
    def __init__(self):
        """TextHandler 인스턴스를 초기화합니다."""
        self.supported_extensions = ['.txt', '.md', '.log']
        self.encoding_fallbacks = ['utf-8', 'cp949', 'latin-1', 'utf-16']
    
    def can_handle(self, file_path: str) -> bool:
        """
        파일이 이 핸들러가 처리할 수 있는 형식인지 확인합니다.
        
        Args:
            file_path (str): 파일 경로
            
        Returns:
            bool: 처리 가능 여부
        """
        return any(file_path.lower().endswith(ext) for ext in self.supported_extensions)
    
    def read_file_content(self, file_path: str, max_size_mb: int = 10) -> str:
        """
        텍스트 파일의 내용을 읽어 반환합니다.
        
        Args:
            file_path (str): 파일 경로
            max_size_mb (int): 읽을 최대 파일 크기 (MB)
            
        Returns:
            str: 파일 내용
        """
        try:
            # 파일 크기 확인
            file_size = os.path.getsize(file_path)
            max_size_bytes = max_size_mb * 1024 * 1024
            
            if file_size > max_size_bytes:
                return f"파일이 너무 큽니다 ({file_size / (1024*1024):.1f}MB). 최대 {max_size_mb}MB까지 지원됩니다."
            
            # 여러 인코딩 시도
            for encoding in self.encoding_fallbacks:
                try:
                    with open(file_path, 'r', encoding=encoding) as file:
                        content = file.read()
                        return content
                except UnicodeDecodeError:
                    continue
            
            return "텍스트 파일을 읽을 수 없습니다. 지원되지 않는 인코딩입니다."
            
        except Exception as e:
            return f"파일 읽기 오류: {e}"
    
    def get_preview(self, file_path: str, max_lines: int = 20) -> str:
        """
        파일의 미리보기 텍스트를 반환합니다.
        
        Args:
            file_path (str): 파일 경로
            max_lines (int): 표시할 최대 줄 수
            
        Returns:
            str: 미리보기 텍스트
        """
        try:
            for encoding in self.encoding_fallbacks:
                try:
                    with open(file_path, 'r', encoding=encoding) as file:
                        lines = []
                        for i, line in enumerate(file):
                            if i >= max_lines:
                                lines.append(f"... (총 {self.get_line_count(file_path)}줄 중 {max_lines}줄 표시)")
                                break
                            lines.append(line.rstrip())
                        return '\n'.join(lines)
                except UnicodeDecodeError:
                    continue
            
            return "미리보기를 생성할 수 없습니다."
            
        except Exception as e:
            return f"미리보기 오류: {e}"
    
    def get_line_count(self, file_path: str) -> int:
        """
        파일의 총 줄 수를 반환합니다.
        
        Args:
            file_path (str): 파일 경로
            
        Returns:
            int: 줄 수
        """
        try:
            for encoding in self.encoding_fallbacks:
                try:
                    with open(file_path, 'r', encoding=encoding) as file:
                        return sum(1 for _ in file)
                except UnicodeDecodeError:
                    continue
            return 0
        except Exception:
            return 0
    
    def get_word_count(self, file_path: str) -> int:
        """
        파일의 단어 수를 반환합니다.
        
        Args:
            file_path (str): 파일 경로
            
        Returns:
            int: 단어 수
        """
        try:
            content = self.read_file_content(file_path)
            if content.startswith("파일") or content.startswith("텍스트"):
                return 0  # 오류 메시지인 경우
            return len(content.split())
        except Exception:
            return 0
    
    def get_metadata(self, file_path: str) -> Dict[str, Any]:
        """
        텍스트 파일의 메타데이터를 반환합니다.
        
        Args:
            file_path (str): 파일 경로
            
        Returns:
            Dict[str, Any]: 메타데이터 정보
        """
        try:
            file_stats = os.stat(file_path)
            file_size = file_stats.st_size
            
            return {
                'filename': os.path.basename(file_path),
                'file_size': file_size,
                'file_size_mb': round(file_size / (1024 * 1024), 2),
                'line_count': self.get_line_count(file_path),
                'word_count': self.get_word_count(file_path),
                'file_type': self.get_file_type(file_path),
                'encoding': self.detect_encoding(file_path),
                'creation_time': file_stats.st_ctime,
                'modification_time': file_stats.st_mtime,
            }
            
        except Exception as e:
            return {
                'error': f"메타데이터 조회 오류: {e}",
                'filename': os.path.basename(file_path),
                'file_size': 0,
            }
    
    def get_file_type(self, file_path: str) -> str:
        """
        파일 확장자를 기반으로 파일 타입을 반환합니다.
        
        Args:
            file_path (str): 파일 경로
            
        Returns:
            str: 파일 타입
        """
        ext = os.path.splitext(file_path)[1].lower()
        type_mapping = {
            '.txt': 'Plain Text',
            '.md': 'Markdown',
            '.log': 'Log File',
        }
        return type_mapping.get(ext, 'Text File')
    
    def detect_encoding(self, file_path: str) -> str:
        """
        파일의 인코딩을 감지합니다.
        
        Args:
            file_path (str): 파일 경로
            
        Returns:
            str: 인코딩 이름
        """
        for encoding in self.encoding_fallbacks:
            try:
                with open(file_path, 'r', encoding=encoding) as file:
                    file.read(1024)  # 첫 1KB만 테스트
                    return encoding
            except UnicodeDecodeError:
                continue
        return 'unknown'
    
    def extract_text(self, file_path: str, max_chars: int = None) -> str:
        """
        검색을 위한 텍스트 추출 (다른 핸들러와 인터페이스 통일)
        
        Args:
            file_path (str): 파일 경로
            max_chars (int): 추출할 최대 문자 수
            
        Returns:
            str: 추출된 텍스트
        """
        content = self.read_file_content(file_path)
        if max_chars and len(content) > max_chars:
            return content[:max_chars] + "..."
        return content