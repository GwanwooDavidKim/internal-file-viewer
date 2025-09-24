# -*- coding: utf-8 -*-
"""
PowerPoint 파일 처리 모듈 (PowerPoint File Handler)

python-pptx를 사용하여 PowerPoint 파일에서 텍스트와 슬라이드 정보를 추출합니다.
"""
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import os
from typing import List, Dict, Any, Optional
import io

# PIL을 안전하게 import (Pillow가 없어도 텍스트 기능은 작동)
try:
    from PIL import Image, ImageDraw, ImageFont
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    print("Warning: Pillow not installed. PowerPoint image preview will not be available.")


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
    
    def render_slide_to_image(self, file_path: str, slide_number: int, width: int = 800, height: int = 600) -> Optional['Image.Image']:
        """
        슬라이드를 간단한 이미지로 렌더링합니다.
        
        Args:
            file_path (str): PowerPoint 파일 경로
            slide_number (int): 슬라이드 번호 (0부터 시작)
            width (int): 이미지 너비
            height (int): 이미지 높이
            
        Returns:
            Optional[Image.Image]: 생성된 이미지 (실패 시 None)
        """
        if not PIL_AVAILABLE:
            print(f"PIL not available. Cannot render slide {slide_number} from {file_path}")
            return None
            
        try:
            prs = Presentation(file_path)
            
            if slide_number >= len(prs.slides) or slide_number < 0:
                return None
            
            slide = prs.slides[slide_number]
            
            # 빈 이미지 생성 (흰색 배경)
            img = Image.new('RGB', (width, height), color='white')
            draw = ImageDraw.Draw(img)
            
            # 한글 지원 폰트 설정 (시스템에서 사용 가능한 폰트 사용)
            korean_fonts = [
                # Windows 폰트
                "malgun.ttf",  # 맑은 고딕
                "gulim.ttc",   # 굴림
                "batang.ttc",  # 바탕
                # macOS 폰트
                "/System/Library/Fonts/AppleSDGothicNeo.ttc",
                "/Library/Fonts/Arial Unicode MS.ttf",
                # Linux 폰트
                "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
                "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
                "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
                # 기본 폰트
                "arial.ttf",
                "DejaVuSans.ttf"
            ]
            
            title_font = None
            text_font = None
            small_font = None
            
            # 사용 가능한 폰트 찾기
            for font_path in korean_fonts:
                try:
                    title_font = ImageFont.truetype(font_path, 36)
                    text_font = ImageFont.truetype(font_path, 24)
                    small_font = ImageFont.truetype(font_path, 18)
                    break
                except:
                    continue
            
            # 폰트를 찾지 못한 경우 기본 폰트 사용
            if not title_font:
                title_font = ImageFont.load_default()
                text_font = ImageFont.load_default()
                small_font = ImageFont.load_default()
            
            y_position = 50
            margin = 40
            
            # 슬라이드 제목 렌더링
            if slide.shapes.title and slide.shapes.title.text:
                title_text = slide.shapes.title.text
                # 제목을 여러 줄로 나누기 (너무 길면)
                title_lines = self._wrap_text(title_text, width - 2*margin, title_font, draw)
                
                for line in title_lines:
                    draw.text((margin, y_position), line, font=title_font, fill='black')
                    y_position += 50
                
                # 제목 아래에 구분선
                draw.line([(margin, y_position + 10), (width - margin, y_position + 10)], fill='gray', width=2)
                y_position += 30
            
            # 슬라이드 내용 렌더링 (텍스트, 이미지, 도표 포함)
            content_items = []
            image_count = 0
            chart_count = 0
            table_count = 0
            
            for shape in slide.shapes:
                # 텍스트 내용 추가
                if hasattr(shape, "text") and shape.text and shape != slide.shapes.title:
                    content_items.append(("text", shape.text))
                
                # 이미지 확인
                elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    image_count += 1
                    content_items.append(("image", f"[이미지 {image_count}]"))
                
                # 차트 확인
                elif hasattr(shape, 'chart') and shape.chart is not None:
                    chart_count += 1
                    chart_title = getattr(shape.chart, 'chart_title', None)
                    chart_name = chart_title.text_frame.text if chart_title and hasattr(chart_title, 'text_frame') else f"차트 {chart_count}"
                    content_items.append(("chart", f"[도표: {chart_name}]"))
                
                # 표 확인
                elif hasattr(shape, 'table') and shape.table is not None:
                    table_count += 1
                    table = shape.table
                    table_info = f"표 {table_count} ({len(table.rows)}x{len(table.columns)})"
                    content_items.append(("table", f"[{table_info}]"))
            
            # 내용을 순서대로 렌더링
            for item_type, item_text in content_items[:8]:  # 최대 8개 항목 표시
                if y_position > height - 100:  # 화면 하단 근처면 중단
                    break
                
                # 항목 타입에 따른 스타일링
                if item_type == "text":
                    # 텍스트를 여러 줄로 나누기
                    lines = self._wrap_text(item_text, width - 2*margin, text_font, draw)
                    
                    for line in lines[:3]:  # 각 항목당 최대 3줄
                        if y_position > height - 50:
                            break
                        draw.text((margin, y_position), f"• {line}", font=text_font, fill='black')
                        y_position += 30
                
                elif item_type == "image":
                    # 이미지 플레이스홀더 표시
                    draw.rectangle([(margin, y_position), (margin + 100, y_position + 60)], 
                                 outline='blue', fill='lightblue', width=2)
                    draw.text((margin + 10, y_position + 20), item_text, font=small_font, fill='darkblue')
                    y_position += 70
                
                elif item_type == "chart":
                    # 차트 플레이스홀더 표시
                    draw.rectangle([(margin, y_position), (margin + 120, y_position + 60)], 
                                 outline='green', fill='lightgreen', width=2)
                    draw.text((margin + 10, y_position + 20), item_text, font=small_font, fill='darkgreen')
                    y_position += 70
                
                elif item_type == "table":
                    # 표 플레이스홀더 표시
                    draw.rectangle([(margin, y_position), (margin + 100, y_position + 40)], 
                                 outline='orange', fill='lightyellow', width=2)
                    draw.text((margin + 10, y_position + 12), item_text, font=small_font, fill='darkorange')
                    y_position += 50
                
                y_position += 10  # 항목 간 간격
            
            # 슬라이드 요소 요약 표시
            if image_count > 0 or chart_count > 0 or table_count > 0:
                summary_y = height - 80
                summary_text = f"포함 요소: "
                if image_count > 0:
                    summary_text += f"이미지 {image_count}개 "
                if chart_count > 0:
                    summary_text += f"도표 {chart_count}개 "
                if table_count > 0:
                    summary_text += f"표 {table_count}개"
                
                draw.text((margin, summary_y), summary_text, font=small_font, fill='gray')
            
            # 슬라이드 번호 표시
            slide_info = f"슬라이드 {slide_number + 1} / {len(prs.slides)}"
            draw.text((width - 200, height - 40), slide_info, font=small_font, fill='gray')
            
            # 폰트 경고 메시지 (개발용)
            if title_font == ImageFont.load_default():
                print("Warning: 한글 폰트를 찾을 수 없어 기본 폰트를 사용합니다. 한글이 깨져 보일 수 있습니다.")
            
            return img
            
        except Exception as e:
            print(f"슬라이드 이미지 렌더링 오류 ({file_path}, 슬라이드 {slide_number}): {e}")
            return None
    
    def _wrap_text(self, text: str, max_width: int, font, draw) -> List[str]:
        """
        텍스트를 지정된 너비에 맞게 여러 줄로 나눕니다.
        
        Args:
            text (str): 원본 텍스트
            max_width (int): 최대 너비
            font: 폰트 객체
            draw: ImageDraw 객체
            
        Returns:
            List[str]: 나뉜 텍스트 줄들
        """
        words = text.split()
        lines = []
        current_line = []
        
        for word in words:
            test_line = ' '.join(current_line + [word])
            
            try:
                bbox = draw.textbbox((0, 0), test_line, font=font)
                text_width = bbox[2] - bbox[0]
            except:
                # textbbox가 없는 경우 textsize 사용 (구버전 호환)
                text_width = draw.textsize(test_line, font=font)[0]
            
            if text_width <= max_width:
                current_line.append(word)
            else:
                if current_line:
                    lines.append(' '.join(current_line))
                    current_line = [word]
                else:
                    lines.append(word)  # 단어가 너무 긴 경우
        
        if current_line:
            lines.append(' '.join(current_line))
        
        return lines