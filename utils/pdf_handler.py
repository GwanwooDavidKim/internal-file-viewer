# -*- coding: utf-8 -*-
"""
PDF 파일 처리 모듈 (PDF File Handler)

PyMuPDF를 사용하여 PDF 파일의 미리보기 이미지 생성 및 텍스트 추출 기능을 제공합니다.
"""
import fitz  # PyMuPDF
from PIL import Image
import io
from typing import List, Optional, Tuple, Dict, Any


class PdfHandler:
    """
    PDF 파일 처리를 위한 클래스입니다.
    
    주요 기능:
    - PDF 페이지를 이미지로 렌더링
    - PDF에서 텍스트 추출
    - PDF 메타데이터 조회
    """
    
    def __init__(self):
        """PdfHandler 인스턴스를 초기화합니다."""
        self.supported_extensions = ['.pdf']
    
    def can_handle(self, file_path: str) -> bool:
        """
        파일이 이 핸들러가 처리할 수 있는 형식인지 확인합니다.
        
        Args:
            file_path (str): 파일 경로
            
        Returns:
            bool: 처리 가능 여부
        """
        return any(file_path.lower().endswith(ext) for ext in self.supported_extensions)
    
    def get_page_count(self, file_path: str) -> int:
        """
        PDF의 총 페이지 수를 반환합니다.
        
        Args:
            file_path (str): PDF 파일 경로
            
        Returns:
            int: 페이지 수 (오류 시 0)
        """
        try:
            with fitz.open(file_path) as doc:
                return len(doc)
        except Exception:
            return 0
    
    def render_page_to_image(self, file_path: str, page_num: int = 0, 
                           zoom: float = 1.0) -> Optional[Image.Image]:
        """
        PDF 페이지를 PIL Image로 렌더링합니다.
        
        Args:
            file_path (str): PDF 파일 경로
            page_num (int): 페이지 번호 (0부터 시작)
            zoom (float): 확대/축소 비율 (1.0 = 100%)
            
        Returns:
            Optional[Image.Image]: 렌더링된 이미지 또는 None
        """
        try:
            with fitz.open(file_path) as doc:
                if page_num >= len(doc) or page_num < 0:
                    return None
                
                page = doc[page_num]
                
                # 매트릭스 설정 (줌 레벨)
                mat = fitz.Matrix(zoom, zoom)
                
                # 페이지를 픽셀맵으로 렌더링
                pix = page.get_pixmap(matrix=mat)
                
                # 픽셀맵을 bytes로 변환
                img_data = pix.tobytes("png")
                
                # PIL Image로 변환
                image = Image.open(io.BytesIO(img_data))
                return image
                
        except Exception as e:
            print(f"PDF 렌더링 오류 ({file_path}, 페이지 {page_num}): {e}")
            return None
    
    def extract_text(self, file_path: str, max_pages: int = None) -> str:
        """
        PDF에서 텍스트를 추출합니다.
        
        Args:
            file_path (str): PDF 파일 경로
            max_pages (int): 추출할 최대 페이지 수 (None이면 전체)
            
        Returns:
            str: 추출된 텍스트
        """
        try:
            text_content = []
            
            with fitz.open(file_path) as doc:
                page_count = len(doc)
                if max_pages:
                    page_count = min(page_count, max_pages)
                
                for page_num in range(page_count):
                    page = doc[page_num]
                    text = page.get_text()
                    if text.strip():
                        text_content.append(f"=== 페이지 {page_num + 1} ===\\n{text}")
            
            return "\\n\\n".join(text_content)
            
        except Exception as e:
            return f"텍스트 추출 오류: {e}"
    
    def get_metadata(self, file_path: str) -> Dict[str, Any]:
        """
        PDF의 메타데이터를 반환합니다.
        
        Args:
            file_path (str): PDF 파일 경로
            
        Returns:
            Dict[str, Any]: 메타데이터 정보
        """
        try:
            with fitz.open(file_path) as doc:
                metadata = doc.metadata
                
                return {
                    'title': metadata.get('title', '제목 없음'),
                    'author': metadata.get('author', '작성자 없음'),
                    'subject': metadata.get('subject', '주제 없음'),
                    'creator': metadata.get('creator', '생성 프로그램 없음'),
                    'producer': metadata.get('producer', '제작 프로그램 없음'),
                    'creation_date': metadata.get('creationDate', '생성일 없음'),
                    'modification_date': metadata.get('modDate', '수정일 없음'),
                    'page_count': len(doc),
                    'encrypted': doc.needs_pass,
                }
                
        except Exception as e:
            return {
                'error': f"메타데이터 조회 오류: {e}",
                'page_count': 0,
                'encrypted': False,
            }
    
    def get_page_preview_info(self, file_path: str, page_num: int = 0) -> Dict[str, Any]:
        """
        특정 페이지의 미리보기 정보를 반환합니다.
        
        Args:
            file_path (str): PDF 파일 경로
            page_num (int): 페이지 번호
            
        Returns:
            Dict[str, Any]: 페이지 미리보기 정보
        """
        try:
            with fitz.open(file_path) as doc:
                if page_num >= len(doc) or page_num < 0:
                    return {'error': '유효하지 않은 페이지 번호'}
                
                page = doc[page_num]
                rect = page.rect
                
                # 첫 100자의 텍스트 추출
                text_preview = page.get_text()[:100]
                if len(page.get_text()) > 100:
                    text_preview += "..."
                
                return {
                    'page_number': page_num + 1,
                    'width': rect.width,
                    'height': rect.height,
                    'text_preview': text_preview.strip() if text_preview.strip() else "[텍스트 없음]",
                    'total_pages': len(doc),
                }
                
        except Exception as e:
            return {'error': f"페이지 정보 오류: {e}"}