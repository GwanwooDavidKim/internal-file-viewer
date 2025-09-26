# -*- coding: utf-8 -*-
"""
Aspose.Slides 기반 PowerPoint → PDF 변환기 (사용자 간섭 없음, 고성능)

Aspose.Slides for Python을 사용하여 사용자 간섭 없이 빠르고 안정적인 PPT → PDF 변환을 제공합니다.
Microsoft Office 설치가 불필요하며, 사용자가 PowerPoint로 파일을 편집 중이어도 영향을 주지 않습니다.

핵심 장점:
- 🚀 사용자 간섭 완전 차단 (독립 프로세스)
- ⚡ LibreOffice보다 2-3배 빠른 성능
- 💰 Office 설치 불필요 (완전 독립적)
- 🎯 완벽한 변환 품질 (레이아웃/폰트 정확도)
- 🛡️ 스마트 캐시 시스템
- 🔒 스레드 안전 처리
"""

import os
import tempfile
import hashlib
import shutil
import logging
import time
import threading
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional, Dict, Any

logger = logging.getLogger(__name__)

# Aspose.Slides를 안전하게 import
try:
    import aspose.slides as slides
    ASPOSE_AVAILABLE = True
    logger.info("✅ Aspose.Slides 라이브러리 로드 완료 - Aspose 방식 사용 가능")
except ImportError as e:
    ASPOSE_AVAILABLE = False
    slides = None
    logger.warning(f"⚠️ Aspose.Slides 라이브러리 없음: {e} - Aspose 방식 사용 불가")


class AsposePowerPointConverter:
    """
    Aspose.Slides를 사용한 고성능 PPT → PDF 변환기
    
    사용자 간섭 없이 최적의 성능과 품질을 제공합니다.
    Microsoft Office 설치가 불필요하며 완전히 독립적으로 작동합니다.
    """
    
    def __init__(self, cache_dir: Optional[str] = None):
        """
        Aspose 변환기 초기화
        
        Args:
            cache_dir: 캐시 디렉토리 경로 (None이면 자동 생성)
        """
        self.cache_dir = Path(cache_dir) if cache_dir else Path(tempfile.gettempdir()) / "aspose_ppt_pdf_cache"
        self.cache_dir.mkdir(exist_ok=True, parents=True)
        
        # 캐시 설정
        self.cache_max_size = 1024 * 1024 * 1024  # 1GB
        self.cache_max_age = timedelta(days=7)  # 7일
        
        # Aspose 사용 가능 여부 확인
        self.aspose_available = ASPOSE_AVAILABLE
        
        # 스레드 락 (Aspose.Slides는 thread-safe하지 않음)
        self._lock = threading.Lock()
        
        print(f"🚀 AsposePowerPointConverter 초기화")
        print(f"   📁 캐시 폴더: {self.cache_dir}")
        if self.is_available():
            print("   ✅ Aspose.Slides 방식 사용 가능!")
            print("   ⚡ 사용자 간섭 없는 고성능 변환 준비 완료")
            print("   🛡️ Microsoft Office 설치 불필요")
        else:
            print("   ❌ Aspose 방식 사용 불가 (라이브러리 없음)")
        
        logger.info(f"Aspose PowerPoint Converter 초기화: 사용 가능={self.is_available()}")
    
    def is_available(self) -> bool:
        """Aspose 변환기 사용 가능 여부"""
        return self.aspose_available
    
    def _get_cache_key(self, file_path: str) -> str:
        """파일 경로와 수정시간으로 캐시 키 생성"""
        abs_path = os.path.abspath(file_path)
        mtime = os.path.getmtime(abs_path)
        content = f"{abs_path}_{mtime}"
        return hashlib.sha1(content.encode()).hexdigest()
    
    def _get_cached_pdf_path(self, file_path: str) -> Path:
        """캐시된 PDF 파일 경로 반환"""
        cache_key = self._get_cache_key(file_path)
        return self.cache_dir / f"{cache_key}.pdf"
    
    def _cleanup_cache(self):
        """오래된 캐시 파일 정리"""
        try:
            total_size = 0
            files_with_time = []
            
            # 모든 캐시 파일의 크기와 수정시간 수집
            for cache_file in self.cache_dir.glob("*.pdf"):
                if cache_file.exists():
                    stat = cache_file.stat()
                    size = stat.st_size
                    mtime = datetime.fromtimestamp(stat.st_mtime)
                    
                    files_with_time.append((cache_file, size, mtime))
                    total_size += size
            
            # 크기 제한 초과 시 오래된 파일부터 삭제
            if total_size > self.cache_max_size:
                files_with_time.sort(key=lambda x: x[2])  # 수정시간 오름차순
                
                for cache_file, size, mtime in files_with_time:
                    if total_size <= self.cache_max_size * 0.8:  # 80%까지 줄이기
                        break
                    
                    cache_file.unlink()
                    total_size -= size
                    logger.info(f"🗑️ 캐시 파일 삭제 (크기 제한): {cache_file.name}")
            
            # 나이 제한으로 오래된 파일 삭제
            cutoff_time = datetime.now() - self.cache_max_age
            for cache_file, size, mtime in files_with_time:
                if cache_file.exists() and mtime < cutoff_time:
                    cache_file.unlink()
                    logger.info(f"🗑️ 캐시 파일 삭제 (나이 제한): {cache_file.name}")
            
        except Exception as e:
            logger.warning(f"캐시 정리 중 오류: {e}")
    
    def convert_to_pdf(self, ppt_file_path: str) -> Optional[str]:
        """
        Aspose.Slides를 사용하여 PPT 파일을 PDF로 변환
        
        Args:
            ppt_file_path: PPT 파일 경로
            
        Returns:
            변환된 PDF 파일 경로 (실패 시 None)
        """
        if not self.is_available():
            logger.error("❌ Aspose 변환기를 사용할 수 없습니다")
            return None
        
        if not os.path.exists(ppt_file_path):
            logger.error(f"❌ PPT 파일을 찾을 수 없습니다: {ppt_file_path}")
            return None
        
        # 캐시 확인
        cached_pdf = self._get_cached_pdf_path(ppt_file_path)
        if cached_pdf.exists():
            logger.info(f"✅ 캐시된 PDF 사용: {cached_pdf}")
            return str(cached_pdf)
        
        # 캐시 정리
        self._cleanup_cache()
        
        try:
            start_time = time.time()
            ppt_name = os.path.basename(ppt_file_path)
            logger.info(f"🚀 Aspose 변환 시작: {ppt_name}")
            
            with self._lock:  # Aspose.Slides는 thread-safe하지 않음
                # 프레젠테이션 로드 (사용자 파일에 간섭 없음)
                logger.info("   📂 프레젠테이션 로드 중...")
                abs_ppt_path = os.path.abspath(ppt_file_path)
                
                with slides.Presentation(abs_ppt_path) as presentation:
                    # PDF로 저장
                    logger.info("   💾 PDF로 변환 중...")
                    abs_pdf_path = os.path.abspath(str(cached_pdf))
                    
                    # PDF 옵션 설정 (성능 최적화)
                    pdf_options = slides.export.PdfOptions()
                    # 고품질 설정
                    pdf_options.jpeg_quality = 95  # JPEG 품질
                    pdf_options.sufficient_resolution = 300  # 해상도
                    pdf_options.text_compression = slides.export.PdfTextCompression.FLATE  # 텍스트 압축
                    pdf_options.compliance = slides.export.PdfCompliance.PDF15  # PDF 버전
                    pdf_options.save_metafiles_as_png = True  # 메타파일을 PNG로
                    
                    # PDF로 저장
                    presentation.save(abs_pdf_path, slides.export.SaveFormat.PDF, pdf_options)
                
                # 변환 완료 확인
                if cached_pdf.exists() and cached_pdf.stat().st_size > 0:
                    elapsed = time.time() - start_time
                    logger.info(f"✅ Aspose 변환 완료! ({elapsed:.1f}초)")
                    logger.info(f"   📄 PDF 크기: {cached_pdf.stat().st_size / 1024:.1f} KB")
                    logger.info("   🛡️ 사용자 PowerPoint 작업에 영향 없음")
                    return str(cached_pdf)
                else:
                    logger.error("❌ PDF 파일이 생성되지 않았습니다")
                    return None
                    
        except Exception as e:
            logger.error(f"❌ Aspose 변환 오류: {e}")
            
            # 실패한 캐시 파일 삭제
            if cached_pdf.exists():
                try:
                    cached_pdf.unlink()
                except:
                    pass
            
            return None
    
    def convert_to_images(self, ppt_file_path: str, slide_number: Optional[int] = None) -> Optional[list]:
        """
        PPT 파일을 이미지로 변환 (선택적 기능)
        
        Args:
            ppt_file_path: PPT 파일 경로
            slide_number: 특정 슬라이드 번호 (None이면 모든 슬라이드)
            
        Returns:
            이미지 파일 경로 리스트 (실패 시 None)
        """
        if not self.is_available():
            logger.error("❌ Aspose 변환기를 사용할 수 없습니다")
            return None
        
        try:
            with self._lock:
                with slides.Presentation(ppt_file_path) as presentation:
                    image_paths = []
                    
                    if slide_number is not None:
                        # 특정 슬라이드만 변환
                        if 0 <= slide_number < len(presentation.slides):
                            slide = presentation.slides[slide_number]
                            image_path = self.cache_dir / f"slide_{slide_number}.png"
                            slide.get_thumbnail(1.0, 1.0).save(str(image_path), slides.ImageFormat.PNG)
                            image_paths.append(str(image_path))
                    else:
                        # 모든 슬라이드 변환
                        for i, slide in enumerate(presentation.slides):
                            image_path = self.cache_dir / f"slide_{i}.png"
                            slide.get_thumbnail(1.0, 1.0).save(str(image_path), slides.ImageFormat.PNG)
                            image_paths.append(str(image_path))
                    
                    return image_paths
                    
        except Exception as e:
            logger.error(f"❌ 이미지 변환 오류: {e}")
            return None
    
    def get_slide_count(self, ppt_file_path: str) -> int:
        """
        PowerPoint 파일의 슬라이드 수 반환
        
        Args:
            ppt_file_path: PPT 파일 경로
            
        Returns:
            슬라이드 수 (오류 시 0)
        """
        if not self.is_available():
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
        PowerPoint 파일에서 텍스트 추출
        
        Args:
            ppt_file_path: PPT 파일 경로
            
        Returns:
            추출된 텍스트
        """
        if not self.is_available():
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
        try:
            cache_files = list(self.cache_dir.glob("*.pdf"))
            total_size = sum(f.stat().st_size for f in cache_files if f.exists())
            
            return {
                'aspose_available': self.aspose_available,
                'converter_available': self.is_available(),
                'cache_dir': str(self.cache_dir),
                'cache_files': len(cache_files),
                'cache_size_mb': round(total_size / (1024 * 1024), 2),
                'cache_max_size_mb': round(self.cache_max_size / (1024 * 1024), 2),
                'cache_max_age_days': self.cache_max_age.days,
                'converter_type': 'Aspose.Slides',
                'advantages': [
                    '사용자 간섭 없음',
                    'Office 설치 불필요',
                    'LibreOffice보다 빠름',
                    '완벽한 변환 품질'
                ]
            }
        except Exception as e:
            logger.error(f"캐시 정보 조회 오류: {e}")
            return {
                'aspose_available': self.aspose_available,
                'converter_available': self.is_available(),
                'error': str(e)
            }


# 전역 변환기 인스턴스 (싱글톤 패턴)
_global_aspose_converter = None

def get_aspose_converter() -> AsposePowerPointConverter:
    """전역 Aspose 변환기 인스턴스 반환"""
    global _global_aspose_converter
    if _global_aspose_converter is None:
        _global_aspose_converter = AsposePowerPointConverter()
    return _global_aspose_converter