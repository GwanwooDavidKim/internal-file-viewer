# -*- coding: utf-8 -*-
"""
PowerPoint 파일 처리 모듈 (PowerPoint File Handler)

python-pptx를 사용하여 PowerPoint 파일에서 텍스트와 슬라이드 정보를 추출합니다.
"""
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import os
import subprocess
import tempfile
from pathlib import Path
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
        self.slide_cache = {}  # 파일별 슬라이드 이미지 캐시
        self.cache_directory = Path(tempfile.gettempdir()) / "ppt_viewer_cache"
        self.cache_directory.mkdir(exist_ok=True)
        
        # 지속 연결을 위한 PowerPoint 인스턴스 관리
        self.current_ppt_app = None
        self.current_presentation = None
        self.current_file_path = None
        
        # 즉시 렌더링용 LRU 캐시 (메모리 효율성)
        self.fast_render_cache = {}
        self.cache_max_size = 20  # 최대 20개 슬라이드 캐시
    
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
        LibreOffice를 사용해서 PowerPoint 슬라이드를 실제 이미지로 렌더링합니다.
        
        Args:
            file_path (str): PowerPoint 파일 경로
            slide_number (int): 슬라이드 번호 (0부터 시작)
            width (int): 이미지 너비 (사용되지 않음 - LibreOffice 기본 해상도 사용)
            height (int): 이미지 높이 (사용되지 않음 - LibreOffice 기본 해상도 사용)
            
        Returns:
            Optional[Image.Image]: 생성된 이미지 (실패 시 None)
        """
        if not PIL_AVAILABLE:
            print(f"PIL not available. Cannot render slide {slide_number} from {file_path}")
            return None
        
        # 캐시에서 우선 확인
        cached_image = self.get_cached_slide(file_path, slide_number)
        if cached_image:
            print(f"💾 캐시된 슬라이드 {slide_number} 이미지 사용 - 즉시 반환!")
            return cached_image
            
        try:
            # Windows COM 자동화 시도 (Windows + PowerPoint가 설치된 경우)
            print(f"🔄 PowerPoint COM 자동화 시도: {file_path}")
            com_image = self._render_slide_with_com(file_path, slide_number)
            if com_image:
                print(f"✅ PowerPoint COM 렌더링 성공! 원본 이미지 반환")
                return com_image
            else:
                print("PowerPoint COM 렌더링 실패 또는 지원되지 않는 환경")
            
            # COM 실패 시 LibreOffice 시도
            print(f"🔄 LibreOffice 렌더링 시도: {file_path}")
            native_image = self._render_slide_with_libreoffice(file_path, slide_number)
            if native_image:
                print(f"✅ LibreOffice 렌더링 성공! 원본 이미지 반환")
                return native_image
            
            # 모든 원본 렌더링 실패 시 텍스트 기반으로 폴백
            print(f"⚠️ 모든 원본 렌더링 실패, 텍스트 기반 렌더링으로 폴백 (슬라이드 {slide_number})")
            
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
    
    def _render_slide_with_libreoffice(self, file_path: str, slide_number: int) -> Optional['Image.Image']:
        """
        LibreOffice를 사용해서 PowerPoint 슬라이드를 실제 이미지로 렌더링합니다.
        
        Args:
            file_path (str): PowerPoint 파일 경로
            slide_number (int): 슬라이드 번호 (0부터 시작)
            
        Returns:
            Optional[Image.Image]: 렌더링된 이미지 또는 None
        """
        if not PIL_AVAILABLE:
            return None
            
        try:
            # 임시 디렉토리 생성
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_path = Path(temp_dir)
                
                # 1단계: PowerPoint를 PDF로 변환
                
                # LibreOffice 바이너리 찾기 (libreoffice 또는 soffice)
                libreoffice_cmd = None
                for cmd_name in ["libreoffice", "soffice"]:
                    try:
                        result = subprocess.run([cmd_name, "--version"], capture_output=True, timeout=5)
                        if result.returncode == 0:
                            libreoffice_cmd = cmd_name
                            break
                    except (subprocess.TimeoutExpired, FileNotFoundError):
                        continue
                
                if not libreoffice_cmd:
                    print("LibreOffice not found (tried 'libreoffice' and 'soffice')")
                    return None
                
                cmd = [
                    libreoffice_cmd,
                    "--headless",
                    "--convert-to", "pdf",
                    "--outdir", str(temp_path),
                    file_path
                ]
                
                print(f"🔄 LibreOffice 명령 실행: {' '.join(cmd)}")
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
                if result.returncode != 0:
                    print(f"❌ LibreOffice PDF conversion failed (cmd: {libreoffice_cmd})")
                    print(f"Return code: {result.returncode}")
                    print(f"stderr: {result.stderr}")
                    print(f"stdout: {result.stdout}")
                    return None
                else:
                    print(f"✅ LibreOffice PDF 변환 성공")
                
                # 생성된 PDF 파일 찾기
                pdf_files = list(temp_path.glob("*.pdf"))
                if not pdf_files:
                    print("No PDF file generated by LibreOffice")
                    return None
                
                actual_pdf_path = pdf_files[0]
                
                # 2단계: PDF에서 특정 페이지를 이미지로 변환 (PyMuPDF 사용)
                try:
                    import fitz  # PyMuPDF
                    
                    doc = fitz.open(str(actual_pdf_path))
                    
                    if slide_number >= len(doc) or slide_number < 0:
                        doc.close()
                        return None
                    
                    # 해당 슬라이드(페이지) 렌더링
                    page = doc.load_page(slide_number)
                    
                    # 고해상도 매트릭스 (기본 150 DPI로 렌더링)
                    # PDF 기본 72 DPI의 약 2배로 고품질 렌더링
                    scale_factor = 2.0
                    mat = fitz.Matrix(scale_factor, scale_factor)
                    pix = page.get_pixmap(matrix=mat)
                    
                    # PIL Image로 변환
                    img_data = pix.tobytes("png")
                    image = Image.open(io.BytesIO(img_data))
                    
                    doc.close()
                    
                    print(f"Successfully rendered slide {slide_number} using LibreOffice + PyMuPDF")
                    return image
                    
                except ImportError:
                    print("PyMuPDF not available for PDF to image conversion")
                    return None
                except Exception as e:
                    print(f"PDF to image conversion failed: {e}")
                    return None
                    
        except subprocess.TimeoutExpired:
            print("LibreOffice conversion timed out")
            return None
        except Exception as e:
            print(f"LibreOffice rendering error: {e}")
            return None
    
    def _render_slide_with_com(self, file_path: str, slide_number: int) -> Optional['Image.Image']:
        """
        Windows PowerPoint COM 자동화를 사용해서 슬라이드를 이미지로 렌더링합니다.
        
        Args:
            file_path (str): PowerPoint 파일 경로
            slide_number (int): 슬라이드 번호 (0부터 시작)
            
        Returns:
            Optional[Image.Image]: 렌더링된 이미지 또는 None
        """
        if not PIL_AVAILABLE:
            return None
        
        # Windows 플랫폼 체크
        import sys
        if sys.platform != 'win32':
            print("PowerPoint COM은 Windows에서만 사용 가능합니다")
            return None
            
        try:
            # Windows COM 라이브러리 import
            try:
                import win32com.client
                import pythoncom
                import os
                import tempfile
                from pathlib import Path
            except ImportError:
                print("Windows COM 라이브러리를 찾을 수 없습니다 (pywin32 설치 필요: pip install pywin32)")
                return None
            
            # COM 초기화 (멀티스레드 환경에서 필요)
            pythoncom.CoInitialize()
            
            ppt_app = None
            presentation = None
            
            try:
                # PowerPoint 애플리케이션 시작
                print("PowerPoint 애플리케이션 시작...")
                ppt_app = win32com.client.Dispatch("PowerPoint.Application")
                
                # 일부 PowerPoint 버전에서는 Visible=False가 차단되므로 True로 설정
                try:
                    ppt_app.Visible = False  # 우선 숨김 모드 시도
                    print("PowerPoint 숨김 모드로 실행")
                except Exception as e:
                    print(f"숨김 모드 실패, 보이는 모드로 실행: {e}")
                    ppt_app.Visible = True  # 보이는 모드로 폴백
                    
                    # 사용자 방해 최소화
                    try:
                        ppt_app.DisplayAlerts = 0  # 알림 비활성화
                        ppt_app.WindowState = 2   # 최소화 (ppWindowMinimized)
                        print("PowerPoint 창 최소화 및 알림 비활성화")
                    except:
                        pass  # 일부 버전에서 지원되지 않을 수 있음
                
                # PowerPoint 파일 열기
                print(f"PowerPoint 파일 열기: {file_path}")
                presentation = ppt_app.Presentations.Open(os.path.abspath(file_path), ReadOnly=True)
                
                # 슬라이드 수 확인
                slide_count = presentation.Slides.Count
                if slide_number >= slide_count or slide_number < 0:
                    print(f"잘못된 슬라이드 번호: {slide_number} (총 {slide_count}개 슬라이드)")
                    return None
                
                # 해당 슬라이드 가져오기 (PowerPoint는 1부터 시작)
                slide = presentation.Slides(slide_number + 1)
                
                # 슬라이드 크기 확인 및 고해상도 계산
                slide_width = presentation.PageSetup.SlideWidth  # 포인트 단위
                slide_height = presentation.PageSetup.SlideHeight  # 포인트 단위
                
                # 고해상도 설정 (200 DPI 기준)
                dpi = 200
                width_px = int(slide_width * dpi / 72)  # 72 포인트 = 1인치
                height_px = int(slide_height * dpi / 72)
                
                # 안전한 임시 파일 경로 생성 (8.3 형식 문제 해결)
                import time
                safe_filename = f"slide_{slide_number}_{int(time.time() * 1000)}.png"
                image_path = self.cache_directory / safe_filename
                
                try:
                    # 슬라이드를 고해상도 PNG로 내보내기
                    print(f"슬라이드를 이미지로 내보내기: {image_path} ({width_px}x{height_px})")
                    slide.Export(str(image_path), "PNG", width_px, height_px)
                    
                    # 내보낸 이미지 로딩
                    if image_path.exists():
                        image = Image.open(str(image_path))
                        print(f"PowerPoint COM 렌더링 성공: {image.size}")
                        
                        # 임시 파일 정리
                        try:
                            image_path.unlink()
                        except:
                            pass
                            
                        return image
                    else:
                        print("이미지 파일이 생성되지 않았습니다")
                        return None
                except Exception as export_error:
                    print(f"슬라이드 Export 오류: {export_error}")
                    return None
                        
            except Exception as e:
                print(f"PowerPoint COM 처리 오류: {e}")
                return None
                
            finally:
                # 리소스 정리 (순서 중요)
                try:
                    if presentation is not None:
                        presentation.Close()
                        print("프레젠테이션 닫기 완료")
                except:
                    pass
                    
                try:
                    if ppt_app is not None:
                        ppt_app.Quit()
                        print("PowerPoint 애플리케이션 종료 완료")
                except:
                    pass
                    
                # COM 정리
                pythoncom.CoUninitialize()
                    
        except Exception as e:
            print(f"PowerPoint COM 자동화 전체 오류: {e}")
            return None
    
    def open_persistent_connection(self, file_path: str) -> bool:
        """
        PowerPoint 파일에 대한 지속적인 연결을 엽니다.
        (사용자 제안: 파일 선택 시 PowerPoint 열어두고 슬라이드별 즉시 렌더링)
        
        Args:
            file_path (str): PowerPoint 파일 경로
            
        Returns:
            bool: 연결 성공 여부
        """
        try:
            # 기존 연결이 있다면 정리
            self.close_persistent_connection()
            
            # Windows 플랫폼 체크
            import sys
            if sys.platform != 'win32':
                print("⚠️ PowerPoint COM은 Windows에서만 사용 가능합니다")
                return False
                
            try:
                import win32com.client
                import pythoncom
            except ImportError:
                print("⚠️ Windows COM 라이브러리를 찾을 수 없습니다 (pywin32 설치 필요)")
                return False
            
            # COM 초기화
            pythoncom.CoInitialize()
            
            print(f"🚀 PowerPoint 지속 연결 시작: {file_path}")
            
            # PowerPoint 애플리케이션 시작
            self.current_ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            
            # 🔥 강력한 백그라운드 실행 (PowerPoint 2016 호환)
            try:
                # PowerPoint 2016은 Visible=False 허용 안함! 바로 강제 숨김으로!
                self.current_ppt_app.DisplayAlerts = 0  # 알림 완전 비활성화
                
                # 즉시 강력한 창 숨김 (사용자 요청: 넉넉한 딜레이 적용)
                import win32gui
                import win32con
                import time
                
                # 넉넉한 딜레이로 PowerPoint 완전 로드까지 대기
                time.sleep(0.3)
                
                def hide_all_powerpoint_windows(hwnd, lparam):
                    window_text = win32gui.GetWindowText(hwnd)
                    class_name = win32gui.GetClassName(hwnd)
                    if ("PowerPoint" in window_text or 
                        "Microsoft PowerPoint" in window_text or
                        "PPTFrameClass" in class_name):
                        win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
                        print(f"🔥 PowerPoint 창 즉시 숨김: {window_text}")
                    return True
                
                win32gui.EnumWindows(hide_all_powerpoint_windows, None)
                print("✅ PowerPoint 애플리케이션 백그라운드 실행 완료!")
                
            except Exception as e:
                print(f"⚠️ 백그라운드 실행 설정 실패: {e}")
                # 실패해도 계속 진행
            
            # PowerPoint 파일 열기
            self.current_presentation = self.current_ppt_app.Presentations.Open(
                os.path.abspath(file_path), ReadOnly=True
            )
            self.current_file_path = file_path
            
            # 🔥 파일 열기 후 강화된 숨김 처리! (사용자 요청: 넉넉한 딜레이로 완전 해결)
            try:
                # 추가로 프레젠테이션 창도 숨김 설정
                if hasattr(self.current_presentation, 'SlideShowSettings'):
                    self.current_presentation.SlideShowSettings.ShowType = 1  # 발표자 모드
                
                # 파일 열기 후 강화된 창 숨김 (넉넉한 딜레이 적용!)
                import win32gui
                import win32con
                import time
                
                # 사용자 요청: 넉넉한 딜레이! 파일이 완전히 열릴 때까지 충분히 대기
                time.sleep(1.0)  # 1초 대기로 증가!
                
                # 여러 번 시도해서 완전히 숨김
                for attempt in range(3):  # 3번 시도
                    try:
                        def hide_all_powerpoint_windows(hwnd, lparam):
                            window_text = win32gui.GetWindowText(hwnd)
                            class_name = win32gui.GetClassName(hwnd)
                            if ("PowerPoint" in window_text or 
                                "Microsoft PowerPoint" in window_text or
                                "PPTFrameClass" in class_name or
                                ".ppt" in window_text.lower() or
                                ".pptx" in window_text.lower() or
                                "Large Area Display" in window_text):  # 파일명도 감지
                                win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
                                print(f"🔥 시도 {attempt+1}: PowerPoint 창 강제 숨김: {window_text}")
                            return True
                        
                        win32gui.EnumWindows(hide_all_powerpoint_windows, None)
                        time.sleep(0.2)  # 각 시도 사이에 잠시 대기
                        
                    except Exception as hide_error:
                        print(f"⚠️ 시도 {attempt+1} 실패: {hide_error}")
                
                print("✅ 파일 열기 후 강화된 PowerPoint 창 숨김 완료!")
                    
            except Exception as post_open_error:
                print(f"⚠️ 파일 열기 후 숨김 처리 실패: {post_open_error}")
            
            slide_count = self.current_presentation.Slides.Count
            print(f"✅ PowerPoint 지속 연결 완료! 슬라이드 수: {slide_count}")
            
            return True
            
        except Exception as e:
            print(f"❌ PowerPoint 지속 연결 실패: {e}")
            self.close_persistent_connection()
            return False
    
    def close_persistent_connection(self):
        """PowerPoint 지속 연결을 종료합니다."""
        try:
            if self.current_presentation:
                self.current_presentation.Close()
                print("프레젠테이션 닫기 완료")
                self.current_presentation = None
                
            if self.current_ppt_app:
                self.current_ppt_app.Quit()
                print("PowerPoint 애플리케이션 종료 완료")
                self.current_ppt_app = None
                
            self.current_file_path = None
            
            # COM 정리
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except:
                pass
                
        except Exception as e:
            print(f"PowerPoint 연결 정리 오류: {e}")
    
    def render_slide_fast(self, slide_number: int, width: int = 800, height: int = 600) -> Optional['Image.Image']:
        """
        이미 열린 PowerPoint에서 특정 슬라이드를 즉시 렌더링합니다.
        (사용자 제안: 슬라이드 변경 시 즉시 렌더링)
        
        Args:
            slide_number (int): 슬라이드 번호 (0부터 시작)
            width (int): 렌더링 너비
            height (int): 렌더링 높이
            
        Returns:
            Optional[Image.Image]: 렌더링된 이미지 또는 None
        """
        if not PIL_AVAILABLE:
            return None
            
        if not self.current_presentation:
            print("❌ PowerPoint 연결이 없습니다")
            return None
            
        try:
            # 캐시 확인 (빠른 응답)
            cache_key = f"{self.current_file_path}_{slide_number}"
            if cache_key in self.fast_render_cache:
                print(f"💾 캐시 히트: 슬라이드 {slide_number}")
                return self.fast_render_cache[cache_key]
            
            # 슬라이드 번호 유효성 검사
            slide_count = self.current_presentation.Slides.Count
            if slide_number < 0 or slide_number >= slide_count:
                print(f"❌ 유효하지 않은 슬라이드 번호: {slide_number} (총 {slide_count}개)")
                return None
                
            slide = self.current_presentation.Slides(slide_number + 1)  # 1부터 시작
            
            # 임시 파일 경로
            import time
            safe_filename = f"fast_slide_{slide_number}_{int(time.time() * 1000)}.png"
            image_path = self.cache_directory / safe_filename
            
            # 슬라이드 내보내기 (즉시!)
            print(f"⚡ 슬라이드 {slide_number + 1} 즉시 렌더링...")
            
            # 고해상도 설정
            dpi = 150
            actual_width = max(width, 800)
            actual_height = max(height, 600)
            
            slide.Export(str(image_path), "PNG", actual_width, actual_height)
            
            # 이미지 로딩 (파일 핸들 누수 방지)
            if image_path.exists():
                # 이미지를 완전히 메모리로 로딩 후 파일 핸들 해제
                with Image.open(str(image_path)) as img:
                    img.load()  # 픽셀 데이터를 메모리로 로딩
                    image = img.copy()  # 독립적인 복사본 생성
                
                print(f"✅ 즉시 렌더링 성공! 크기: {image.size}")
                
                # 이제 안전하게 임시 파일 정리
                try:
                    image_path.unlink()
                except:
                    pass
                
                # LRU 캐시에 저장 (메모리 효율성)
                self._add_to_cache(cache_key, image)
                    
                return image
            else:
                print("❌ 이미지 파일 생성 실패")
                return None
                
        except Exception as e:
            print(f"❌ 즉시 렌더링 오류: {e}")
            return None
    
    def _add_to_cache(self, cache_key: str, image: 'Image.Image'):
        """LRU 캐시에 이미지를 추가합니다."""
        # 캐시 크기 제한
        if len(self.fast_render_cache) >= self.cache_max_size:
            # 가장 오래된 항목 제거 (간단한 FIFO)
            oldest_key = next(iter(self.fast_render_cache))
            del self.fast_render_cache[oldest_key]
        
        self.fast_render_cache[cache_key] = image
    
    def is_connected(self) -> bool:
        """PowerPoint 연결 상태를 확인합니다."""
        return self.current_presentation is not None
    
    def render_all_slides_batch(self, file_path: str) -> Dict[int, 'Image.Image']:
        """
        PowerPoint 파일의 모든 슬라이드를 한 번에 이미지로 렌더링합니다.
        (깜빡임 없이 빠른 슬라이드 전환을 위한 배치 처리)
        
        Args:
            file_path (str): PowerPoint 파일 경로
            
        Returns:
            Dict[int, Image.Image]: {슬라이드번호: 이미지} 딕셔너리
        """
        if not PIL_AVAILABLE:
            return {}
        
        # 캐시 키 생성
        import hashlib
        file_stat = os.stat(file_path)
        cache_key = f"{file_path}_{file_stat.st_mtime}_{file_stat.st_size}"
        
        # 이미 캐시된 경우 반환
        if cache_key in self.slide_cache:
            print("💾 캐시된 슬라이드 이미지 사용")
            return self.slide_cache[cache_key]
        
        print(f"🚀 PowerPoint 배치 렌더링 시작: {file_path}")
        rendered_slides = {}
        
        # Windows 플랫폼 체크
        import sys
        if sys.platform != 'win32':
            print("PowerPoint COM은 Windows에서만 사용 가능합니다")
            return {}
            
        try:
            # Windows COM 라이브러리 import
            try:
                import win32com.client
                import pythoncom
                import os
                import tempfile
                from pathlib import Path
            except ImportError:
                print("Windows COM 라이브러리를 찾을 수 없습니다 (pywin32 설치 필요: pip install pywin32)")
                return {}
            
            # COM 초기화
            pythoncom.CoInitialize()
            
            ppt_app = None
            presentation = None
            
            try:
                # PowerPoint 애플리케이션 시작
                print("PowerPoint 애플리케이션 시작 (배치 모드)...")
                ppt_app = win32com.client.Dispatch("PowerPoint.Application")
                
                # Visible 설정
                try:
                    ppt_app.Visible = False
                    print("PowerPoint 숨김 모드로 실행")
                except Exception as e:
                    print(f"숨김 모드 실패, 보이는 모드로 실행: {e}")
                    ppt_app.Visible = True
                    
                    # 사용자 방해 최소화
                    try:
                        ppt_app.DisplayAlerts = 0
                        ppt_app.WindowState = 2
                        print("PowerPoint 창 최소화 및 알림 비활성화")
                    except:
                        pass
                
                # PowerPoint 파일 열기
                print(f"PowerPoint 파일 열기: {file_path}")
                presentation = ppt_app.Presentations.Open(os.path.abspath(file_path), ReadOnly=True)
                
                # 슬라이드 수 확인
                slide_count = presentation.Slides.Count
                print(f"📄 총 {slide_count}개 슬라이드 배치 렌더링 시작")
                
                # 슬라이드 크기 확인
                slide_width = presentation.PageSetup.SlideWidth
                slide_height = presentation.PageSetup.SlideHeight
                
                # 고해상도 설정
                dpi = 200
                width_px = int(slide_width * dpi / 72)
                height_px = int(slide_height * dpi / 72)
                
                # 모든 슬라이드 배치 렌더링
                for slide_idx in range(slide_count):
                    try:
                        slide = presentation.Slides(slide_idx + 1)
                        
                        # 안전한 임시 파일 경로
                        import time
                        safe_filename = f"batch_slide_{slide_idx}_{int(time.time() * 1000)}.png"
                        image_path = self.cache_directory / safe_filename
                        
                        # 슬라이드 내보내기
                        print(f"📸 슬라이드 {slide_idx + 1}/{slide_count} 렌더링...")
                        slide.Export(str(image_path), "PNG", width_px, height_px)
                        
                        # 이미지 로딩 및 저장
                        if image_path.exists():
                            image = Image.open(str(image_path))
                            rendered_slides[slide_idx] = image
                            
                            # 임시 파일 정리
                            try:
                                image_path.unlink()
                            except:
                                pass
                        else:
                            print(f"⚠️ 슬라이드 {slide_idx} 이미지 생성 실패")
                            
                    except Exception as slide_error:
                        print(f"❌ 슬라이드 {slide_idx} 렌더링 오류: {slide_error}")
                        continue
                
                # 캐시에 저장
                self.slide_cache[cache_key] = rendered_slides
                print(f"✅ 배치 렌더링 완료! {len(rendered_slides)}개 슬라이드 캐시됨")
                
                return rendered_slides
                        
            except Exception as e:
                print(f"PowerPoint 배치 렌더링 오류: {e}")
                return {}
                
            finally:
                # 리소스 정리
                try:
                    if presentation is not None:
                        presentation.Close()
                        print("프레젠테이션 닫기 완료")
                except:
                    pass
                    
                try:
                    if ppt_app is not None:
                        ppt_app.Quit()
                        print("PowerPoint 애플리케이션 종료 완료")
                except:
                    pass
                    
                # COM 정리
                pythoncom.CoUninitialize()
                    
        except Exception as e:
            print(f"PowerPoint 배치 렌더링 전체 오류: {e}")
            return {}
    
    def get_cached_slide(self, file_path: str, slide_number: int) -> Optional['Image.Image']:
        """
        캐시에서 슬라이드 이미지를 가져옵니다.
        
        Args:
            file_path (str): PowerPoint 파일 경로
            slide_number (int): 슬라이드 번호 (0부터 시작)
            
        Returns:
            Optional[Image.Image]: 캐시된 이미지 또는 None
        """
        try:
            file_stat = os.stat(file_path)
            cache_key = f"{file_path}_{file_stat.st_mtime}_{file_stat.st_size}"
            
            if cache_key in self.slide_cache and slide_number in self.slide_cache[cache_key]:
                return self.slide_cache[cache_key][slide_number]
                
        except Exception as e:
            print(f"캐시 확인 오류: {e}")
            
        return None