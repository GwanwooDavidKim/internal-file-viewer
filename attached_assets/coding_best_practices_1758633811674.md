# 코딩 모범 사례 (Coding Best Practices)

## 1. 코드 스타일 (Code Style)
- 모든 Python 코드는 **PEP 8** 스타일 가이드를 준수합니다.
- `black`, `flake8` 같은 코드 포매터 및 린터(linter) 사용을 강력히 권장하여 일관성을 유지합니다.

## 2. 네이밍 컨벤션 (Naming Conventions)
- **변수, 함수:** `snake_case` (예: `file_path`, `load_user_data`)
- **상수:** `UPPER_SNAKE_CASE` (예: `MAX_FILE_SIZE`, `DEFAULT_PATH`)
- **클래스:** `PascalCase` (예: `MainApplication`, `FileViewerWidget`)
- **모듈:** 짧고 의미 있는 `snake_case` (예: `utils.py`, `file_handlers.py`)

## 3. 주석 및 문서화 (Comments & Documentation)
- 복잡한 로직이나 비즈니스 규칙이 담긴 코드에는 '왜' 이렇게 작성했는지 설명하는 주석을 추가합니다.
- 모든 함수와 클래스에는 기능, 인자(Arguments), 반환 값(Returns)을 설명하는 **Docstring**을 작성합니다.

## 4. 모듈화 및 구조 (Modularity & Structure)
- 코드를 기능 단위로 분리하여 모듈화합니다.
  - `main.py`: 애플리케이션 진입점
  - `ui/`: UI와 관련된 모든 클래스 및 로직 (예: `main_window.py`, `viewer_widget.py`)
  - `core/`: 핵심 비즈니스 로직 (예: `file_search.py`, `auth.py`, `indexing.py`)
  - `utils/`: 파일 처리 등 보조 기능 (예: `pdf_handler.py`, `excel_handler.py`)
  - `config.py`: 설정 값 (계정 정보, 기본 경로 등)

## 5. 에러 처리 (Error Handling)
- 파일을 읽거나 쓰는 등 실패할 가능성이 있는 모든 작업에는 `try...except` 블록을 사용하여 예외를 처리합니다.
- 사용자에게는 명확하고 이해하기 쉬운 에러 메시지를 UI를 통해 보여줍니다. (예: "파일을 열 수 없습니다. 손상되었거나 지원하지 않는 형식입니다.")

## 6. 의존성 관리 (Dependency Management)
- 프로젝트에 필요한 모든 라이브러리는 `requirements.txt` 파일에 명시하고 관리합니다.
- `pip freeze > requirements.txt` 명령어를 사용하여 버전을 고정합니다.

## 7. 버전 관리 (Version Control)
- **Git**을 사용하여 모든 코드 변경 사항을 관리합니다.
- 커밋 메시지는 "[기능] 로그인 기능 추가", "[수정] 파일 로딩 속도 개선"과 같이 명확하고 구체적으로 작성합니다.

## 8. 플랫폼 종속성 관리 (Managing Platform Dependencies)
- **목적:** `pywin32`를 사용한 Windows COM 자동화와 같이 특정 운영체제(OS)에서만 동작하는 코드를 안전하게 관리하기 위함입니다.
- **지침 8.1 (모듈 분리):**
  - 플랫폼 종속적인 코드는 범용 코드와 분리하여 별도의 모듈(예: `windows_ppt_handler.py`)로 작성합니다.
- **지침 8.2 (조건부 임포트):**
  - 해당 모듈을 불러올(import) 때는, 시스템 플랫폼을 확인하는 조건문을 사용하여 다른 OS 환경에서 오류가 발생하지 않도록 방지합니다.
  ```python
  import sys
  
  # main.py 또는 관련 모듈에서
  if sys.platform == 'win32':
      from . import windows_ppt_handler as ppt_handler
  else:
      # Windows가 아닐 경우, 대체 핸들러를 사용하거나 기능을 비활성화
      from . import aspose_ppt_handler as ppt_handler
  ```
