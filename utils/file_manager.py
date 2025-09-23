# -*- coding: utf-8 -*-
"""
파일 관리자 모듈 (File Manager)

다양한 파일 형식에 대한 통합 처리 인터페이스를 제공합니다.
각 파일 형식별 핸들러를 통합하여 일관된 방식으로 파일을 처리할 수 있습니다.
"""
import os
from typing import Dict, Any, Optional, List, Union
import config
from utils.pdf_handler import PdfHandler
from utils.image_handler import ImageHandler
from utils.excel_handler import ExcelHandler
from utils.word_handler import WordHandler
from utils.powerpoint_handler import PowerPointHandler


class FileManager:
    """
    파일 처리를 통합 관리하는 클래스입니다.
    
    다양한 파일 형식에 대해 적절한 핸들러를 선택하고,
    일관된 인터페이스를 통해 파일 정보를 제공합니다.
    """
    
    def __init__(self):
        """FileManager 인스턴스를 초기화합니다."""
        self.handlers = {
            'pdf': PdfHandler(),
            'image': ImageHandler(),
            'excel': ExcelHandler(),
            'word': WordHandler(),
            'powerpoint': PowerPointHandler(),
        }
        
        # 핸들러 우선순위 (확인 순서)
        self.handler_priority = ['pdf', 'image', 'excel', 'word', 'powerpoint']
    
    def get_file_type(self, file_path: str) -> Optional[str]:
        """
        파일 경로를 기반으로 파일 타입을 결정합니다.
        각 핸들러의 can_handle 메서드를 사용하여 동적으로 판단합니다.
        
        Args:
            file_path (str): 파일 경로
            
        Returns:
            Optional[str]: 파일 타입 ('pdf', 'image', 'excel', 'word', 'powerpoint') 또는 None
        """
        # 각 핸들러에게 순서대로 파일 처리 가능 여부 확인
        for handler_type in self.handler_priority:
            handler = self.handlers.get(handler_type)
            if handler and handler.can_handle(file_path):
                return handler_type
        return None
    
    def is_supported_file(self, file_path: str) -> bool:
        """
        파일이 지원되는 형식인지 확인합니다.
        
        Args:
            file_path (str): 파일 경로
            
        Returns:
            bool: 지원 여부
        """
        return self.get_file_type(file_path) is not None
    
    def get_file_handler(self, file_path: str):
        """
        파일에 적절한 핸들러를 반환합니다.
        
        Args:
            file_path (str): 파일 경로
            
        Returns:
            Handler 인스턴스 또는 None
        """
        file_type = self.get_file_type(file_path)
        if file_type:
            return self.handlers.get(file_type)
        return None
    
    def get_file_info(self, file_path: str) -> Dict[str, Any]:
        """
        파일의 기본 정보를 반환합니다.
        
        Args:
            file_path (str): 파일 경로
            
        Returns:
            Dict[str, Any]: 파일 정보
        """
        try:
            if not os.path.exists(file_path):
                return {'error': '파일을 찾을 수 없습니다', 'supported': False}
            
            # 기본 파일 정보
            file_size = os.path.getsize(file_path)
            file_name = os.path.basename(file_path)
            file_ext = os.path.splitext(file_path.lower())[1]
            
            basic_info = {
                'filename': file_name,
                'filepath': file_path,
                'extension': file_ext,
                'file_size': file_size,
                'file_size_mb': round(file_size / (1024 * 1024), 2),
                'file_type': self.get_file_type(file_path),
                'supported': self.is_supported_file(file_path),
            }
            
            # 파일 타입별 상세 정보
            if basic_info['supported']:
                handler = self.get_file_handler(file_path)
                if handler:
                    try:
                        if basic_info['file_type'] == 'pdf':
                            detail_info = handler.get_metadata(file_path)
                        elif basic_info['file_type'] == 'image':
                            detail_info = handler.get_image_info(file_path)
                        elif basic_info['file_type'] == 'excel':
                            detail_info = handler.get_workbook_info(file_path)
                        elif basic_info['file_type'] == 'word':
                            detail_info = handler.get_document_info(file_path)
                        elif basic_info['file_type'] == 'powerpoint':
                            detail_info = handler.get_presentation_info(file_path)
                        else:
                            detail_info = {}
                        
                        # 기본 정보와 상세 정보 합치기
                        basic_info.update(detail_info)
                        
                    except Exception as e:
                        basic_info['detail_error'] = f"상세 정보 조회 오류: {e}"
            
            return basic_info
            
        except Exception as e:
            return {'error': f"파일 정보 조회 오류: {e}", 'supported': False}
    
    def extract_text(self, file_path: str, **kwargs) -> str:
        """
        파일에서 텍스트를 추출합니다.
        
        Args:
            file_path (str): 파일 경로
            **kwargs: 핸들러별 추가 옵션
            
        Returns:
            str: 추출된 텍스트
        """
        handler = self.get_file_handler(file_path)
        if not handler:
            return "지원되지 않는 파일 형식입니다."
        
        try:
            file_type = self.get_file_type(file_path)
            
            if file_type == 'pdf':
                return handler.extract_text(file_path, kwargs.get('max_pages'))
            elif file_type == 'word':
                return handler.extract_text(file_path, kwargs.get('include_structure', True))
            elif file_type == 'powerpoint':
                return handler.extract_all_text(file_path)
            elif file_type == 'excel':
                # Excel의 경우 첫 번째 시트의 데이터를 텍스트로 변환
                sheet_data = handler.read_sheet(file_path)
                if 'data' in sheet_data and sheet_data['data']:
                    text_lines = []
                    for row in sheet_data['data'][:10]:  # 처음 10행만
                        values = [str(v) for v in row.values() if v]
                        if values:
                            text_lines.append(" | ".join(values))
                    return "\\n".join(text_lines)
                return "Excel 데이터를 읽을 수 없습니다."
            elif file_type == 'image':
                # 이미지는 텍스트 추출이 불가능
                return "이미지 파일은 텍스트 추출이 불가능합니다."
            else:
                return "알 수 없는 파일 형식입니다."
                
        except Exception as e:
            return f"텍스트 추출 오류: {e}"
    
    def search_in_file(self, file_path: str, search_term: str, 
                      max_results: int = 20) -> List[Dict[str, Any]]:
        """
        파일 내에서 텍스트를 검색합니다.
        
        Args:
            file_path (str): 파일 경로
            search_term (str): 검색할 텍스트
            max_results (int): 최대 결과 수
            
        Returns:
            List[Dict[str, Any]]: 검색 결과 목록
        """
        handler = self.get_file_handler(file_path)
        if not handler:
            return [{'error': '지원되지 않는 파일 형식입니다'}]
        
        try:
            file_type = self.get_file_type(file_path)
            
            if file_type == 'word':
                return handler.search_in_document(file_path, search_term, max_results)
            elif file_type == 'powerpoint':
                return handler.search_in_presentation(file_path, search_term, max_results)
            elif file_type == 'excel':
                # Excel의 경우 모든 시트에서 검색
                results = []
                sheet_names = handler.get_sheet_names(file_path)
                for sheet_name in sheet_names:
                    sheet_results = handler.search_in_sheet(
                        file_path, sheet_name, search_term, max_results - len(results)
                    )
                    for result in sheet_results:
                        result['sheet_name'] = sheet_name
                    results.extend(sheet_results)
                    if len(results) >= max_results:
                        break
                return results
            else:
                # PDF와 이미지는 기본 텍스트 추출 후 검색
                text = self.extract_text(file_path)
                if search_term.lower() in text.lower():
                    # 간단한 검색 결과 반환
                    lines = text.split('\\n')
                    results = []
                    for i, line in enumerate(lines):
                        if search_term.lower() in line.lower():
                            results.append({
                                'location': f'라인 {i + 1}',
                                'context': line[:200] + ('...' if len(line) > 200 else ''),
                                'full_text': line,
                            })
                            if len(results) >= max_results:
                                break
                    return results
                return []
                
        except Exception as e:
            return [{'error': f"파일 검색 오류: {e}"}]
    
    def get_preview_data(self, file_path: str, **kwargs) -> Dict[str, Any]:
        """
        파일의 미리보기 데이터를 반환합니다.
        
        Args:
            file_path (str): 파일 경로
            **kwargs: 미리보기 옵션
            
        Returns:
            Dict[str, Any]: 미리보기 데이터
        """
        handler = self.get_file_handler(file_path)
        if not handler:
            return {'error': '지원되지 않는 파일 형식입니다'}
        
        try:
            file_type = self.get_file_type(file_path)
            
            if file_type == 'pdf':
                page_num = kwargs.get('page', 0)
                return handler.get_page_preview_info(file_path, page_num)
            elif file_type == 'image':
                return handler.get_image_info(file_path)
            elif file_type == 'excel':
                sheet_name = kwargs.get('sheet_name')
                return handler.read_sheet(file_path, sheet_name, max_rows=20, max_cols=10)
            elif file_type == 'word':
                return {'structure': handler.get_document_structure(file_path)}
            elif file_type == 'powerpoint':
                slide_num = kwargs.get('slide', 0)
                return handler.extract_text_from_slide(file_path, slide_num)
            else:
                return {'error': '미리보기를 지원하지 않는 파일 형식입니다'}
                
        except Exception as e:
            return {'error': f"미리보기 생성 오류: {e}"}
    
    def get_supported_extensions(self) -> List[str]:
        """
        지원되는 모든 파일 확장자를 반환합니다.
        각 핸들러의 supported_extensions를 동적으로 수집합니다.
        
        Returns:
            List[str]: 지원되는 확장자 목록
        """
        extensions = []
        for handler in self.handlers.values():
            if hasattr(handler, 'supported_extensions'):
                extensions.extend(handler.supported_extensions)
        return sorted(list(set(extensions)))  # 중복 제거 및 정렬