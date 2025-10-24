# -*- coding: utf-8 -*-
"""
Aspose.Slides 기반 PowerPoint 변환기 (평가판 사용)

Aspose.Slides for Python을 사용하여 PowerPoint 파일을 PDF로 변환합니다.
사용자 간섭 없이 안전하게 변환이 가능합니다.

주요 특징:
- 사용자가 PowerPoint 파일을 편집 중이어도 간섭 없음
- Microsoft Office 설치 불필요
- LibreOffice보다 빠른 성능
- 평가판 사용 (워터마크 포함)
"""

import os
import time
import logging
from pathlib import Path
import threading
from typing import Optional, Dict, Any

logger = logging.getLogger(__name__)

# Aspose.Slides를 안전하게 import
try:
    import aspose.slides as slides
    ASPOSE_AVAILABLE = True
    logger.info("[완료] Aspose.Slides 라이브러리 로드 완료 (평가판) - 워터마크 포함")
except ImportError as e:
    ASPOSE_AVAILABLE = False
    slides = None
    logger.warning(f"[경고] Aspose.Slides 라이브러리 없음: {e} - Aspose 방식 사용 불가")


class AsposePowerPointConverter:
    """
    Aspose.Slides 기반 PowerPoint → PDF 변환기 (평가판)
    
    특징:
    - 사용자 간섭 없는 안전한 변환
    - Microsoft Office 설치 불필요
    - 고성능 변환 (LibreOffice보다 빠름)
    - 평가판 사용 (워터마크 포함)
    """
    
    def __init__(self, cache_dir: str = "/tmp/aspose_ppt_pdf_cache"):
        """
        Aspose PowerPoint 변환기를 초기화합니다.
        
        Args:
            cache_dir (str): PDF 캐시 디렉토리 경로
        """
        self.cache_dir = Path(cache_dir)
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        
        # 스레드 안전성을 위한 락
        self._lock = threading.Lock()
        
        # 캐시 설정
        self.max_cache_size_mb = 1024  # 1GB
        self.max_cache_age_days = 7
        
        print("[시작] AsposePowerPointConverter 초기화 (평가판)")
        print(f"   [폴더] 캐시 폴더: {self.cache_dir}")
        
        if ASPOSE_AVAILABLE:
            print("   [완료] Aspose.Slides 방식 사용 가능! (평가판 - 워터마크 포함)")
            print("   [변환기] 사용자 간섭 없는 고성능 변환 준비 완료")
            print("   [안전] Microsoft Office 설치 불필요")
            print("   [평가판] 평가판 - PDF에 워터마크가 포함될 수 있음")
        else:
            print("   [오류] Aspose 방식 사용 불가 (라이브러리 없음)")
    
    def is_available(self) -> bool:
        """Aspose 변환기 사용 가능 여부 확인"""
        return ASPOSE_AVAILABLE
    
    def _get_cache_key(self, ppt_file_path: str) -> str:
        """파일의 캐시 키 생성 (경로 + 수정시간)"""
        abs_path = os.path.abspath(ppt_file_path)
        if os.path.exists(abs_path):
            mtime = os.path.getmtime(abs_path)
            return f"{abs_path}_{mtime}".replace("/", "_").replace("\\", "_").replace(":", "_")
        return abs_path.replace("/", "_").replace("\\", "_").replace(":", "_")
    
    def _cleanup_cache(self):
        """오래된 캐시 파일 정리"""
        try:
            current_time = time.time()
            max_age_seconds = self.max_cache_age_days * 24 * 3600
            
            for cache_file in self.cache_dir.glob("*.pdf"):
                if current_time - cache_file.stat().st_mtime > max_age_seconds:
                    cache_file.unlink()
                    logger.info(f"[삭제] 오래된 캐시 파일 삭제: {cache_file.name}")
                    
        except Exception as e:
            logger.warning(f"캐시 정리 중 오류: {e}")
    
    def convert_to_pdf(self, ppt_file_path: str) -> Optional[str]:
        """
        PowerPoint 파일을 PDF로 변환합니다 (캐시 지원).
        
        Args:
            ppt_file_path (str): PowerPoint 파일 경로
            
        Returns:
            Optional[str]: 변환된 PDF 파일 경로 (실패 시 None)
        """
        if not self.is_available():
            logger.error("[오류] Aspose 변환기를 사용할 수 없습니다")
            return None
        
        if not os.path.exists(ppt_file_path):
            logger.error(f"[오류] PowerPoint 파일을 찾을 수 없습니다: {ppt_file_path}")
            return None
        
        try:
            # 캐시 키 생성
            cache_key = self._get_cache_key(ppt_file_path)
            cached_pdf = self.cache_dir / f"{cache_key}.pdf"
            
            # 캐시된 파일이 있으면 반환
            if cached_pdf.exists() and cached_pdf.stat().st_size > 0:
                logger.info(f"[캐시] 캐시된 PDF 사용: {os.path.basename(ppt_file_path)}")
                return str(cached_pdf)
            
            # 변환 시작
            logger.info(f"[처리중] Aspose.Slides로 PowerPoint → PDF 변환 시작: {os.path.basename(ppt_file_path)}")
            start_time = time.time()
            
            with self._lock:
                # 프레젠테이션 로드 (사용자 파일에 간섭 없음)
                logger.info("   [폴더] 프레젠테이션 로드 중...")
                abs_ppt_path = os.path.abspath(ppt_file_path)
                
                # slides 모듈이 None이 아님을 확인 (타입 체킹용)
                if slides is None:
                    logger.error("[오류] slides 모듈이 None입니다")
                    return None
                
                with slides.Presentation(abs_ppt_path) as presentation:
                    # PDF로 저장
                    logger.info("   [저장] PDF로 변환 중...")
                    abs_pdf_path = os.path.abspath(str(cached_pdf))
                    
                    # PDF 옵션 설정 (평가판용 - 기본 설정)
                    pdf_options = slides.export.PdfOptions()
                    # 평가판용 적절한 품질 설정
                    pdf_options.jpeg_quality = 85  # 적당한 JPEG 품질
                    pdf_options.sufficient_resolution = 200  # 적당한 해상도
                    pdf_options.text_compression = slides.export.PdfTextCompression.FLATE  # 텍스트 압축
                    pdf_options.compliance = slides.export.PdfCompliance.PDF15  # PDF 버전
                    pdf_options.save_metafiles_as_png = True  # 메타파일을 PNG로
                    
                    # PDF로 저장 (평가판 - 워터마크 포함될 수 있음)
                    presentation.save(abs_pdf_path, slides.export.SaveFormat.PDF, pdf_options)
                
                # 변환 완료 확인
                if cached_pdf.exists() and cached_pdf.stat().st_size > 0:
                    elapsed = time.time() - start_time
                    logger.info(f"[완료] Aspose.Slides 변환 완료! ({elapsed:.1f}초)")
                    logger.info(f"   [파일] PDF 생성: {os.path.basename(cached_pdf)}")
                    logger.info("   [평가판] 평가판 - 워터마크가 포함될 수 있습니다")
                    
                    # 오래된 캐시 정리
                    self._cleanup_cache()
                    
                    return str(cached_pdf)
                else:
                    logger.error("[오류] PDF 파일이 생성되지 않았습니다")
                    return None
                    
        except Exception as e:
            logger.error(f"[오류] Aspose 변환 오류: {e}")
            # 실패한 캐시 파일 정리
            if 'cached_pdf' in locals() and cached_pdf.exists():
                try:
                    cached_pdf.unlink()
                except:
                    pass
            return None
    
    def convert_to_images(self, ppt_file_path: str, slide_number: int = None) -> Optional[list]:
        """
        PowerPoint 슬라이드를 이미지로 변환합니다.
        
        Args:
            ppt_file_path (str): PowerPoint 파일 경로
            slide_number (int, optional): 특정 슬라이드 번호 (None이면 모든 슬라이드)
            
        Returns:
            Optional[list]: 생성된 이미지 파일 경로들 (실패 시 None)
        """
        if not self.is_available():
            logger.error("[오류] Aspose 변환기를 사용할 수 없습니다")
            return None
        
        if slides is None:
            logger.error("[오류] slides 모듈이 None입니다")
            return None
        
        try:
            with self._lock:
                with slides.Presentation(ppt_file_path) as presentation:
                    image_paths = []
                    
                    # 캐시 키를 사용한 고유 폴더 생성
                    cache_key = self._get_cache_key(ppt_file_path)
                    images_dir = self.cache_dir / f"images_{cache_key}"
                    images_dir.mkdir(exist_ok=True)
                    
                    if slide_number is not None:
                        # 특정 슬라이드만 변환
                        if 0 <= slide_number < len(presentation.slides):
                            slide = presentation.slides[slide_number]
                            image_path = images_dir / f"slide_{slide_number}.png"
                            slide.get_thumbnail(1.0, 1.0).save(str(image_path), slides.ImageFormat.PNG)
                            image_paths.append(str(image_path))
                    else:
                        # 모든 슬라이드 변환
                        for i, slide in enumerate(presentation.slides):
                            image_path = images_dir / f"slide_{i}.png"
                            slide.get_thumbnail(1.0, 1.0).save(str(image_path), slides.ImageFormat.PNG)
                            image_paths.append(str(image_path))
                    
                    return image_paths
                    
        except Exception as e:
            logger.error(f"[오류] 이미지 변환 오류: {e}")
            return None
    
    def get_slide_count(self, ppt_file_path: str) -> int:
        """
        PowerPoint 파일의 슬라이드 수를 반환합니다.
        
        Args:
            ppt_file_path (str): PowerPoint 파일 경로
            
        Returns:
            int: 슬라이드 수 (오류 시 0)
        """
        if not self.is_available():
            return 0
        
        if slides is None:
            logger.error("[오류] slides 모듈이 None입니다")
            return 0
        
        try:
            with self._lock:
                with slides.Presentation(ppt_file_path) as presentation:
                    return len(presentation.slides)
        except Exception as e:
            logger.error(f"슬라이드 수 확인 오류: {e}")
            return 0
    
    def extract_text(self, ppt_file_path: str) -> str:
        """
        PowerPoint 파일에서 텍스트를 추출합니다.
        
        Args:
            ppt_file_path (str): PowerPoint 파일 경로
            
        Returns:
            str: 추출된 텍스트
        """
        if not self.is_available():
            return ""
        
        if slides is None:
            logger.error("[오류] slides 모듈이 None입니다")
            return ""
        
        try:
            with self._lock:
                with slides.Presentation(ppt_file_path) as presentation:
                    all_text = []
                    
                    for i, slide in enumerate(presentation.slides):
                        slide_text = [f"=== 슬라이드 {i + 1} ==="]
                        
                        for shape in slide.shapes:
                            if hasattr(shape, 'text_frame') and shape.text_frame:
                                for paragraph in shape.text_frame.paragraphs:
                                    for portion in paragraph.portions:
                                        if portion.text.strip():
                                            slide_text.append(portion.text)
                        
                        all_text.append("\n".join(slide_text))
                    
                    return "\n\n".join(all_text)
                    
        except Exception as e:
            logger.error(f"텍스트 추출 오류: {e}")
            return f"텍스트 추출 오류: {e}"
    
    def get_cache_info(self) -> Dict[str, Any]:
        """캐시 정보 반환"""
        cache_files = list(self.cache_dir.glob("*.pdf"))
        cache_size = sum(f.stat().st_size for f in cache_files) / (1024 * 1024)  # MB
        
        return {
            'aspose_available': ASPOSE_AVAILABLE,
            'converter_available': self.is_available(),
            'cache_dir': str(self.cache_dir),
            'cache_files': len(cache_files),
            'cache_size_mb': round(cache_size, 2),
            'cache_max_size_mb': self.max_cache_size_mb,
            'cache_max_age_days': self.max_cache_age_days,
            'converter_type': 'Aspose.Slides (평가판)',
            'advantages': [
                '사용자 간섭 없음',
                'Office 설치 불필요', 
                'LibreOffice보다 빠름',
                '완벽한 변환 품질',
                '평가판 (워터마크 포함)'
            ],
            'note': '평가판 사용 - PDF에 워터마크가 포함될 수 있습니다'
        }


# 싱글톤 인스턴스
_aspose_converter_instance = None

def get_aspose_converter() -> AsposePowerPointConverter:
    """
    Aspose PowerPoint 변환기의 싱글톤 인스턴스를 반환합니다.
    
    Returns:
        AsposePowerPointConverter: 변환기 인스턴스
    """
    global _aspose_converter_instance
    if _aspose_converter_instance is None:
        _aspose_converter_instance = AsposePowerPointConverter()
    return _aspose_converter_instance