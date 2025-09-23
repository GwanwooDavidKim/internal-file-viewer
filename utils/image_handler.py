# -*- coding: utf-8 -*-
"""
이미지 파일 처리 모듈 (Image File Handler)

Pillow를 사용하여 다양한 형식의 이미지 파일을 처리합니다.
"""
from PIL import Image, ImageOps
import os
from typing import Optional, Tuple, Dict, Any


class ImageHandler:
    """
    이미지 파일 처리를 위한 클래스입니다.
    
    주요 기능:
    - 이미지 로드 및 리사이즈
    - 이미지 회전 및 변환
    - 이미지 메타데이터 조회
    - 썸네일 생성
    """
    
    def __init__(self):
        """ImageHandler 인스턴스를 초기화합니다."""
        self.supported_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', 
                                   '.tiff', '.tif', '.webp']  # .svg는 Pillow 기본 지원 안됨
        self.max_size = (1920, 1080)  # 최대 표시 크기
        self.thumbnail_size = (300, 300)  # 썸네일 크기
    
    def can_handle(self, file_path: str) -> bool:
        """
        파일이 이 핸들러가 처리할 수 있는 형식인지 확인합니다.
        
        Args:
            file_path (str): 파일 경로
            
        Returns:
            bool: 처리 가능 여부
        """
        return any(file_path.lower().endswith(ext) for ext in self.supported_extensions)
    
    def load_image(self, file_path: str, max_size: Optional[Tuple[int, int]] = None) -> Optional[Image.Image]:
        """
        이미지 파일을 로드합니다.
        
        Args:
            file_path (str): 이미지 파일 경로
            max_size (Optional[Tuple[int, int]]): 최대 크기 (width, height)
            
        Returns:
            Optional[Image.Image]: 로드된 이미지 또는 None
        """
        try:
            if not os.path.exists(file_path):
                return None
            
            # 이미지 로드
            image = Image.open(file_path)
            
            # RGBA 모드로 변환 (투명도 지원)
            if image.mode != 'RGBA':
                image = image.convert('RGBA')
            
            # 크기 조정
            if max_size is None:
                max_size = self.max_size
            
            # 비율을 유지하면서 리사이즈
            image.thumbnail(max_size, Image.Resampling.LANCZOS)
            
            return image
            
        except Exception as e:
            print(f"이미지 로드 오류 ({file_path}): {e}")
            return None
    
    def create_thumbnail(self, file_path: str, size: Optional[Tuple[int, int]] = None) -> Optional[Image.Image]:
        """
        이미지의 썸네일을 생성합니다.
        
        Args:
            file_path (str): 이미지 파일 경로
            size (Optional[Tuple[int, int]]): 썸네일 크기
            
        Returns:
            Optional[Image.Image]: 썸네일 이미지 또는 None
        """
        try:
            if size is None:
                size = self.thumbnail_size
            
            image = Image.open(file_path)
            
            # 썸네일 생성 (비율 유지, 크롭)
            thumbnail = ImageOps.fit(image, size, Image.Resampling.LANCZOS)
            
            # RGB 모드로 변환 (호환성)
            if thumbnail.mode != 'RGB':
                thumbnail = thumbnail.convert('RGB')
            
            return thumbnail
            
        except Exception as e:
            print(f"썸네일 생성 오류 ({file_path}): {e}")
            return None
    
    def get_image_info(self, file_path: str) -> Dict[str, Any]:
        """
        이미지의 상세 정보를 반환합니다.
        
        Args:
            file_path (str): 이미지 파일 경로
            
        Returns:
            Dict[str, Any]: 이미지 정보
        """
        try:
            if not os.path.exists(file_path):
                return {'error': '파일을 찾을 수 없습니다'}
            
            # 파일 크기
            file_size = os.path.getsize(file_path)
            
            with Image.open(file_path) as image:
                # 기본 정보
                info = {
                    'filename': os.path.basename(file_path),
                    'format': image.format or '알 수 없음',
                    'mode': image.mode,
                    'size': image.size,
                    'width': image.width,
                    'height': image.height,
                    'file_size': file_size,
                    'file_size_mb': round(file_size / (1024 * 1024), 2),
                }
                
                # EXIF 데이터 (있는 경우)
                if hasattr(image, '_getexif') and image._getexif() is not None:
                    exif_data = image._getexif()
                    info['has_exif'] = True
                    
                    # 주요 EXIF 정보 추출
                    if 306 in exif_data:  # DateTime
                        info['datetime'] = exif_data[306]
                    if 271 in exif_data:  # Make (카메라 제조사)
                        info['camera_make'] = exif_data[271]
                    if 272 in exif_data:  # Model (카메라 모델)
                        info['camera_model'] = exif_data[272]
                else:
                    info['has_exif'] = False
                
                # 애니메이션 여부 (GIF 등)
                info['is_animated'] = getattr(image, 'is_animated', False)
                if info['is_animated']:
                    info['frame_count'] = getattr(image, 'n_frames', 1)
                
                return info
                
        except Exception as e:
            return {'error': f"이미지 정보 조회 오류: {e}"}
    
    def rotate_image(self, image: Image.Image, degrees: int) -> Image.Image:
        """
        이미지를 회전시킵니다.
        
        Args:
            image (Image.Image): 원본 이미지
            degrees (int): 회전 각도 (시계 반대 방향)
            
        Returns:
            Image.Image: 회전된 이미지
        """
        try:
            return image.rotate(degrees, expand=True)
        except Exception as e:
            print(f"이미지 회전 오류: {e}")
            return image
    
    def get_dominant_colors(self, file_path: str, num_colors: int = 5) -> Optional[list]:
        """
        이미지의 주요 색상을 추출합니다.
        
        Args:
            file_path (str): 이미지 파일 경로
            num_colors (int): 추출할 색상 수
            
        Returns:
            Optional[list]: RGB 색상 리스트 또는 None
        """
        try:
            with Image.open(file_path) as image:
                # 이미지 크기 축소 (성능 향상)
                image = image.resize((150, 150))
                
                # RGB 모드로 변환
                if image.mode != 'RGB':
                    image = image.convert('RGB')
                
                # 색상 양자화
                quantized = image.quantize(colors=num_colors)
                palette = quantized.getpalette()
                
                # RGB 색상 추출
                colors = []
                for i in range(num_colors):
                    r = palette[i * 3]
                    g = palette[i * 3 + 1]  
                    b = palette[i * 3 + 2]
                    colors.append((r, g, b))
                
                return colors
                
        except Exception as e:
            print(f"주요 색상 추출 오류 ({file_path}): {e}")
            return None