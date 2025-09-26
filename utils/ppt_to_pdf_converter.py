"""
PPT to PDF Converter
PPT 파일을 PDF로 변환하는 안전한 변환기

사용자의 PowerPoint 작업에 영향을 주지 않기 위해:
1. LibreOffice 헤드리스 모드 사용 (완전 격리)
2. 임시 폴더에 PDF 저장
3. 캐시 시스템으로 성능 최적화
4. 자동 정리 기능
"""

import os
import tempfile
import hashlib
import shutil
import subprocess
from pathlib import Path
from datetime import datetime, timedelta
from typing import Optional, Dict, Any
import logging

logger = logging.getLogger(__name__)

class PptToPdfConverter:
    """PPT 파일을 PDF로 변환하는 클래스"""
    
    def __init__(self, cache_dir: Optional[str] = None):
        """
        초기화
        
        Args:
            cache_dir: 캐시 디렉토리 경로 (None이면 자동 생성)
        """
        self.cache_dir = Path(cache_dir) if cache_dir else Path(tempfile.gettempdir()) / "ppt_pdf_cache"
        self.cache_dir.mkdir(exist_ok=True, parents=True)
        
        # LibreOffice 실행 파일 찾기
        self.libreoffice_path = self._find_libreoffice()
        
        # 캐시 정보
        self.cache_max_size = 1024 * 1024 * 1024  # 1GB
        self.cache_max_age = timedelta(days=7)  # 7일
        
        print(f"🔄 PptToPdfConverter 초기화 - 캐시 폴더: {self.cache_dir}")
        if self.libreoffice_path:
            print(f"✅ LibreOffice 발견: {self.libreoffice_path}")
            logger.info(f"✅ LibreOffice 발견: {self.libreoffice_path}")
        else:
            print("❌ LibreOffice를 찾을 수 없습니다. PPT 미리보기가 제한됩니다.")
            logger.error("❌ LibreOffice를 찾을 수 없습니다. PPT 미리보기가 제한됩니다.")
    
    def _find_libreoffice(self) -> Optional[str]:
        """LibreOffice 실행 파일을 찾습니다"""
        possible_paths = [
            # Windows
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            # Linux
            "/usr/bin/libreoffice",
            "/usr/bin/soffice",
            # macOS
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        ]
        
        # PATH에서 찾기 (타임아웃 늘리고 디버깅 강화)
        logger.info("🔍 PATH에서 LibreOffice 검색 중...")
        try:
            result = subprocess.run(["soffice", "--version"], 
                                  capture_output=True, text=True, timeout=15)
            if result.returncode == 0:
                logger.info(f"✅ PATH에서 soffice 발견: {result.stdout.strip()}")
                return "soffice"
            else:
                logger.warning(f"⚠️ soffice 실행 실패: returncode={result.returncode}")
        except subprocess.TimeoutExpired:
            logger.warning("⚠️ soffice --version 타임아웃 (15초)")
        except FileNotFoundError:
            logger.warning("⚠️ soffice 명령어를 찾을 수 없음")
        except Exception as e:
            logger.warning(f"⚠️ soffice 실행 중 예외: {e}")
        
        # libreoffice 명령도 시도
        logger.info("🔍 libreoffice 명령으로 재시도...")
        try:
            result = subprocess.run(["libreoffice", "--version"], 
                                  capture_output=True, text=True, timeout=15)
            if result.returncode == 0:
                logger.info(f"✅ PATH에서 libreoffice 발견: {result.stdout.strip()}")
                return "libreoffice"
            else:
                logger.warning(f"⚠️ libreoffice 실행 실패: returncode={result.returncode}")
        except subprocess.TimeoutExpired:
            logger.warning("⚠️ libreoffice --version 타임아웃 (15초)")
        except FileNotFoundError:
            logger.warning("⚠️ libreoffice 명령어를 찾을 수 없음")
        except Exception as e:
            logger.warning(f"⚠️ libreoffice 실행 중 예외: {e}")
        
        # 직접 경로에서 찾기
        logger.info("🔍 하드코딩된 경로에서 LibreOffice 검색 중...")
        for path in possible_paths:
            if os.path.exists(path):
                logger.info(f"✅ 경로에서 발견: {path}")
                return path
            else:
                logger.debug(f"❌ 경로 없음: {path}")
        
        logger.error("❌ LibreOffice를 찾을 수 없습니다")
        return None
    
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
    
    def convert_to_pdf(self, ppt_file_path: str) -> Optional[str]:
        """
        PPT 파일을 PDF로 변환합니다
        
        Args:
            ppt_file_path: PPT 파일 경로
            
        Returns:
            변환된 PDF 파일 경로 (실패 시 None)
        """
        if not os.path.exists(ppt_file_path):
            logger.error(f"❌ PPT 파일을 찾을 수 없습니다: {ppt_file_path}")
            return None
        
        # 캐시 확인
        cached_pdf = self._get_cached_pdf_path(ppt_file_path)
        if cached_pdf.exists():
            logger.info(f"✅ 캐시된 PDF 사용: {cached_pdf}")
            return str(cached_pdf)
        
        # LibreOffice가 없으면 변환 불가
        if not self.libreoffice_path:
            logger.error("❌ LibreOffice를 찾을 수 없어 PDF 변환 불가")
            return None
        
        try:
            ppt_name = os.path.basename(ppt_file_path)
            print(f"\\n📋 LibreOffice 변환 시작: {ppt_name}")
            print(f"   🔧 헤드리스 모드 (사용자 작업 방해 없음)")
            logger.info(f"🔄 PPT → PDF 변환 시작: {ppt_file_path}")
            
            # LibreOffice 헤드리스 모드로 변환
            cmd = [
                self.libreoffice_path,
                "--headless",
                "--convert-to", "pdf",
                "--outdir", str(self.cache_dir),
                ppt_file_path
            ]
            
            result = subprocess.run(
                cmd, 
                capture_output=True, 
                text=True, 
                timeout=120,  # 2분 타임아웃
                cwd=str(self.cache_dir)
            )
            
            if result.returncode == 0:
                # 변환된 파일명 찾기
                original_name = Path(ppt_file_path).stem
                temp_pdf = self.cache_dir / f"{original_name}.pdf"
                
                if temp_pdf.exists():
                    # 캐시 키로 파일명 변경
                    shutil.move(str(temp_pdf), str(cached_pdf))
                    file_size_kb = cached_pdf.stat().st_size / 1024
                    print(f"✅ LibreOffice 변환 완료! {ppt_name} → PDF")
                    print(f"   📄 PDF 크기: {file_size_kb:.1f} KB")
                    print(f"   🔧 LibreOffice 헤드리스 엔진 사용 성공!")
                    logger.info(f"✅ PDF 변환 완료: {cached_pdf}")
                    
                    # 캐시 정리
                    self._cleanup_old_cache()
                    
                    return str(cached_pdf)
                else:
                    logger.error(f"❌ 변환된 PDF 파일을 찾을 수 없음: {temp_pdf}")
                    return None
            else:
                logger.error(f"❌ PDF 변환 실패: {result.stderr}")
                return None
                
        except subprocess.TimeoutExpired:
            logger.error("❌ PDF 변환 타임아웃 (2분 초과)")
            return None
        except Exception as e:
            logger.error(f"❌ PDF 변환 오류: {e}")
            return None
    
    def _cleanup_old_cache(self):
        """오래된 캐시 파일 정리"""
        try:
            current_time = datetime.now()
            total_size = 0
            files_info = []
            
            # 모든 캐시 파일 정보 수집
            for file_path in self.cache_dir.glob("*.pdf"):
                if file_path.is_file():
                    stat = file_path.stat()
                    size = stat.st_size
                    mtime = datetime.fromtimestamp(stat.st_mtime)
                    
                    total_size += size
                    files_info.append({
                        'path': file_path,
                        'size': size,
                        'mtime': mtime,
                        'age': current_time - mtime
                    })
            
            # 나이 기준 정리 (7일 이상)
            for file_info in files_info[:]:
                if file_info['age'] > self.cache_max_age:
                    file_info['path'].unlink()
                    files_info.remove(file_info)
                    total_size -= file_info['size']
                    logger.debug(f"🗑️ 오래된 캐시 삭제: {file_info['path']}")
            
            # 크기 기준 정리 (1GB 초과)
            if total_size > self.cache_max_size:
                # 오래된 것부터 정렬
                files_info.sort(key=lambda x: x['mtime'])
                
                for file_info in files_info:
                    if total_size <= self.cache_max_size * 0.8:  # 80%까지 줄이기
                        break
                    
                    file_info['path'].unlink()
                    total_size -= file_info['size']
                    logger.debug(f"🗑️ 크기 제한으로 캐시 삭제: {file_info['path']}")
            
            logger.debug(f"📊 캐시 정리 완료 - 파일: {len(files_info)}개, 크기: {total_size/1024/1024:.1f}MB")
            
        except Exception as e:
            logger.error(f"❌ 캐시 정리 오류: {e}")
    
    def clear_cache(self):
        """모든 캐시 파일 삭제"""
        try:
            for file_path in self.cache_dir.glob("*.pdf"):
                file_path.unlink()
            logger.info("🗑️ 모든 캐시 파일 삭제 완료")
        except Exception as e:
            logger.error(f"❌ 캐시 삭제 오류: {e}")
    
    def get_cache_info(self) -> Dict[str, Any]:
        """캐시 정보 반환"""
        try:
            files = list(self.cache_dir.glob("*.pdf"))
            total_size = sum(f.stat().st_size for f in files)
            
            return {
                'cache_dir': str(self.cache_dir),
                'file_count': len(files),
                'total_size_mb': total_size / 1024 / 1024,
                'libreoffice_available': self.libreoffice_path is not None
            }
        except Exception as e:
            logger.error(f"❌ 캐시 정보 조회 오류: {e}")
            return {'error': str(e)}


# 전역 변환기 인스턴스
_converter = None

def get_converter() -> PptToPdfConverter:
    """전역 변환기 인스턴스 반환"""
    global _converter
    if _converter is None:
        _converter = PptToPdfConverter()
    return _converter