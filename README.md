# 사내 파일 뷰어 (Internal File Viewer)

## 프로젝트 개요

사내 파일 뷰어는 팀 내에 흩어져 있는 다양한 포맷의 업무 자료(PDF, PPT, Excel, Word, 이미지)를 하나의 애플리케이션에서 신속하게 탐색하고 내용을 확인할 수 있도록 도와주는 데스크톱 애플리케이션입니다.

## 주요 기능

### 🔐 보안 인증 시스템
- bcrypt 해싱을 사용한 안전한 비밀번호 관리
- 관리자 및 일반 사용자 역할 구분
- 계정 만료 관리 및 세션 유효성 검사

### 📄 다양한 파일 형식 지원
- **PDF**: 페이지별 미리보기 및 텍스트 추출
- **이미지**: JPG, PNG, GIF, BMP, TIFF, WebP 지원
- **Microsoft Office**: Excel (.xlsx, .xlsm), Word (.docx), PowerPoint (.pptx)
- **파일 정보**: 메타데이터, 크기, 생성/수정일 표시

### 🖥️ 직관적인 GUI
- PyQt6 기반의 모던한 사용자 인터페이스
- 2-Pane 레이아웃 (파일 탐색기 + 콘텐츠 뷰어)
- 실시간 파일 미리보기
- 탭 기반 인터페이스 (파일 뷰어 + 검색)

### 🔍 강력한 검색 기능
- 전문 검색 (Full-text Search) 지원
- 역 인덱스 기반 빠른 검색
- 한국어/영어 토큰화 및 불용어 처리
- 검색 결과 하이라이팅

### 📊 로깅 및 모니터링
- 포괄적인 로깅 시스템
- 파일 작업, 인증, 성능 측정 로그
- 로그 파일 자동 로테이션

## 기술 스택

- **언어**: Python 3.11+
- **GUI 프레임워크**: PyQt6
- **파일 처리**: 
  - PyMuPDF (PDF)
  - python-pptx (PowerPoint)
  - python-docx (Word)
  - openpyxl (Excel)
  - Pillow (이미지)
- **보안**: bcrypt
- **데이터 처리**: pandas, numpy

## 설치 및 실행

### 1. 요구사항 설치
```bash
pip install -r requirements.txt
```

### 2. 애플리케이션 실행

#### GUI 모드 (권장)
```bash
python main.py --gui
```

#### 콘솔 모드
```bash
python main.py
```

### 3. 데모 계정
- **관리자**: admin / admin123
- **팀원**: team1 / team123

## 프로젝트 구조

```
사내파일뷰어/
├── main.py                 # 메인 진입점
├── config.py              # 설정 파일
├── requirements.txt       # 의존성 목록
├── core/                  # 핵심 모듈
│   ├── __init__.py
│   └── auth.py           # 인증 시스템
├── ui/                   # GUI 컴포넌트
│   ├── __init__.py
│   ├── main_window.py    # 메인 윈도우
│   ├── login_dialog.py   # 로그인 다이얼로그
│   ├── file_browser.py   # 파일 브라우저
│   ├── content_viewer.py # 콘텐츠 뷰어
│   └── search_widget.py  # 검색 위젯
└── utils/                # 유틸리티 모듈
    ├── __init__.py
    ├── file_manager.py   # 파일 관리자
    ├── pdf_handler.py    # PDF 처리
    ├── image_handler.py  # 이미지 처리
    ├── excel_handler.py  # Excel 처리
    ├── word_handler.py   # Word 처리
    ├── powerpoint_handler.py # PowerPoint 처리
    ├── search_indexer.py # 검색 인덱서
    └── logger.py         # 로깅 시스템
```

## 주요 특징

### 보안성
- 하드코딩된 인증 정보 제거
- bcrypt를 통한 안전한 비밀번호 해싱
- 세션 관리 및 권한 제어

### 확장성
- 모듈화된 파일 핸들러 구조
- 동적 파일 형식 지원 추가 가능
- 플러그인 방식의 아키텍처

### 성능
- 백그라운드 파일 로딩
- 인덱스 기반 빠른 검색
- 메모리 효율적인 미리보기

### 사용성
- 직관적인 2-Pane 레이아웃
- 실시간 파일 필터링
- 컨텍스트 기반 검색 결과

## 개발자 정보

이 프로젝트는 팀 내 문서 관리 효율성 향상을 위해 개발되었습니다.

## 라이선스

이 프로젝트는 사내 사용을 위한 것입니다.