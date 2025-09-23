# -*- coding: utf-8 -*-
"""
PowerPoint 파일 처리 모듈 (PowerPoint File Handler)

python-pptx를 사용하여 PowerPoint 파일에서 텍스트와 슬라이드 정보를 추출합니다.
"""
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import os
from typing import List, Dict, Any, Optional


class PowerPointHandler:
    """
    PowerPoint 파일 처리를 위한 클래스입니다.
    
    주요 기능:
    - PowerPoint 슬라이드 텍스트 추출
    - 슬라이드 구조 분석
    - 프레젠테이션 메타데이터 조회
    - 슬라이드별 내용 요약
    """
    
    def __init__(self):
        """PowerPointHandler 인스턴스를 초기화합니다."""
        self.supported_extensions = ['.pptx']  # .ppt는 python-pptx에서 지원하지 않음
    
    def can_handle(self, file_path: str) -> bool:
        """
        파일이 이 핸들러가 처리할 수 있는 형식인지 확인합니다.
        
        Args:
            file_path (str): 파일 경로
            
        Returns:
            bool: 처리 가능 여부
        """
        return any(file_path.lower().endswith(ext) for ext in self.supported_extensions)
    
    def get_slide_count(self, file_path: str) -> int:
        """
        PowerPoint의 총 슬라이드 수를 반환합니다.
        
        Args:
            file_path (str): PowerPoint 파일 경로
            
        Returns:
            int: 슬라이드 수 (오류 시 0)
        """
        try:
            prs = Presentation(file_path)
            return len(prs.slides)
        except Exception:
            return 0
    
    def extract_text_from_slide(self, file_path: str, slide_number: int) -> Dict[str, Any]:
        """
        특정 슬라이드에서 텍스트를 추출합니다.
        
        Args:
            file_path (str): PowerPoint 파일 경로
            slide_number (int): 슬라이드 번호 (0부터 시작)
            
        Returns:
            Dict[str, Any]: 슬라이드 텍스트 정보
        """
        try:
            prs = Presentation(file_path)
            
            if slide_number >= len(prs.slides) or slide_number < 0:
                return {'error': '유효하지 않은 슬라이드 번호'}
            
            slide = prs.slides[slide_number]
            
            # 슬라이드 제목 추출
            title = ""
            if slide.shapes.title:
                title = slide.shapes.title.text
            
            # 텍스트 내용 추출
            text_content = []
            bullet_points = []
            
            for shape in slide.shapes:
                # 텍스트가 있는 shape만 처리
                if hasattr(shape, "text") and hasattr(shape, "text_frame") and shape.text.strip():
                    # 제목이 아닌 경우에만 추가
                    if shape != slide.shapes.title:
                        text_content.append(shape.text)
                        
                        # 텍스트 프레임이 있는 경우 단락별로 분석
                        if shape.text_frame:
                            for paragraph in shape.text_frame.paragraphs:
                                if paragraph.text.strip():
                                    bullet_points.append({
                                        'text': paragraph.text,
                                        'level': paragraph.level,
                                    })
            
            # 이미지 및 기타 객체 카운트
            image_count = 0
            chart_count = 0
            table_count = 0
            
            for shape in slide.shapes:
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    image_count += 1
                elif shape.shape_type == MSO_SHAPE_TYPE.CHART:
                    chart_count += 1
                elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                    table_count += 1
            
            return {
                'slide_number': slide_number + 1,
                'title': title or '[제목 없음]',
                'text_content': text_content,
                'bullet_points': bullet_points,
                'full_text': "\\n".join(text_content),
                'image_count': image_count,
                'chart_count': chart_count,
                'table_count': table_count,
                'total_shapes': len(slide.shapes),
            }
            
        except Exception as e:
            return {'error': f"슬라이드 텍스트 추출 오류: {e}"}
    
    def extract_all_text(self, file_path: str) -> str:
        """
        전체 프레젠테이션에서 텍스트를 추출합니다.
        
        Args:
            file_path (str): PowerPoint 파일 경로
            
        Returns:
            str: 추출된 전체 텍스트
        """
        try:
            prs = Presentation(file_path)
            all_text = []
            
            for i, slide in enumerate(prs.slides):
                slide_text = []
                
                # 슬라이드 제목
                if slide.shapes.title:
                    slide_text.append(f"=== 슬라이드 {i + 1}: {slide.shapes.title.text} ===")
                else:
                    slide_text.append(f"=== 슬라이드 {i + 1} ===")
                
                # 슬라이드 내용
                for shape in slide.shapes:
                    if hasattr(shape, "text") and hasattr(shape, "text_frame") and shape.text.strip():
                        if shape != slide.shapes.title:
                            slide_text.append(shape.text)
                
                all_text.append("\\n".join(slide_text))
            
            return "\\n\\n".join(all_text)
            
        except Exception as e:
            return f"PowerPoint 텍스트 추출 오류: {e}"
    
    def get_presentation_info(self, file_path: str) -> Dict[str, Any]:
        """
        프레젠테이션의 상세 정보를 반환합니다.
        
        Args:
            file_path (str): PowerPoint 파일 경로
            
        Returns:
            Dict[str, Any]: 프레젠테이션 정보
        """
        try:
            if not os.path.exists(file_path):
                return {'error': '파일을 찾을 수 없습니다'}
            
            prs = Presentation(file_path)
            
            # 기본 정보
            file_size = os.path.getsize(file_path)
            slide_count = len(prs.slides)
            
            # 프레젠테이션 메타데이터
            core_props = prs.core_properties
            
            # 슬라이드 크기 정보
            slide_width = prs.slide_width or 9144000  # 기본값 (10인치)
            slide_height = prs.slide_height or 6858000  # 기본값 (7.5인치)
            
            # 각 슬라이드의 요약 정보
            slides_summary = []
            total_images = 0
            total_charts = 0
            total_tables = 0
            
            for i, slide in enumerate(prs.slides):
                # 슬라이드 제목
                title = ""
                if slide.shapes.title:
                    title = slide.shapes.title.text
                
                # 객체 카운트
                images = charts = tables = 0
                for shape in slide.shapes:
                    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        images += 1
                    elif shape.shape_type == MSO_SHAPE_TYPE.CHART:
                        charts += 1
                    elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                        tables += 1
                
                total_images += images
                total_charts += charts
                total_tables += tables
                
                # 텍스트 블록 수
                text_shapes = sum(1 for shape in slide.shapes 
                                if hasattr(shape, "text") and hasattr(shape, "text_frame") and shape.text.strip())
                
                slides_summary.append({
                    'slide_number': i + 1,
                    'title': title or f'[슬라이드 {i + 1}]',
                    'text_shapes': text_shapes,
                    'images': images,
                    'charts': charts,
                    'tables': tables,
                    'total_shapes': len(slide.shapes),
                })
            
            info = {
                'filename': os.path.basename(file_path),
                'file_size': file_size,
                'file_size_mb': round(file_size / (1024 * 1024), 2),
                'slide_count': slide_count,
                'slide_width': slide_width,
                'slide_height': slide_height,
                'slide_size_inches': (round(slide_width / 914400, 2), 
                                    round(slide_height / 914400, 2)),
                'total_images': total_images,
                'total_charts': total_charts,
                'total_tables': total_tables,
                'slides_summary': slides_summary,
                'title': core_props.title or '제목 없음',
                'author': core_props.author or '작성자 없음',
                'subject': core_props.subject or '주제 없음',
                'keywords': core_props.keywords or '키워드 없음',
                'created': str(core_props.created) if core_props.created else '생성일 없음',
                'modified': str(core_props.modified) if core_props.modified else '수정일 없음',
                'revision': core_props.revision or '버전 없음',
            }
            
            return info
            
        except Exception as e:
            return {'error': f"프레젠테이션 정보 조회 오류: {e}"}
    
    def search_in_presentation(self, file_path: str, search_term: str, 
                             max_results: int = 20) -> List[Dict[str, Any]]:
        """
        프레젠테이션에서 특정 텍스트를 검색합니다.
        
        Args:
            file_path (str): PowerPoint 파일 경로
            search_term (str): 검색할 텍스트
            max_results (int): 최대 결과 수
            
        Returns:
            List[Dict[str, Any]]: 검색 결과 목록
        """
        try:
            prs = Presentation(file_path)
            results = []
            search_term = search_term.lower()
            
            for i, slide in enumerate(prs.slides):
                # 슬라이드 제목에서 검색
                if slide.shapes.title and search_term in slide.shapes.title.text.lower():
                    results.append({
                        'slide_number': i + 1,
                        'location': '제목',
                        'type': 'title',
                        'text': slide.shapes.title.text,
                        'context': slide.shapes.title.text,
                    })
                    
                    if len(results) >= max_results:
                        return results
                
                # 슬라이드 내용에서 검색
                for shape_idx, shape in enumerate(slide.shapes):
                    if hasattr(shape, "text") and hasattr(shape, "text_frame") and shape.text.strip():
                        if search_term in shape.text.lower():
                            # 검색어 주변 컨텍스트 추출
                            text = shape.text
                            search_pos = text.lower().find(search_term)
                            
                            start = max(0, search_pos - 50)
                            end = min(len(text), search_pos + len(search_term) + 50)
                            context = text[start:end]
                            
                            if start > 0:
                                context = "..." + context
                            if end < len(text):
                                context = context + "..."
                            
                            results.append({
                                'slide_number': i + 1,
                                'location': f'텍스트 블록 {shape_idx + 1}',
                                'type': 'content',
                                'text': text,
                                'context': context,
                            })
                            
                            if len(results) >= max_results:
                                return results
            
            return results
            
        except Exception as e:
            return [{'error': f"프레젠테이션 검색 오류: {e}"}]