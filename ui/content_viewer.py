# -*- coding: utf-8 -*-
"""
콘텐츠 뷰어 위젯 (Content Viewer Widget)

다양한 파일 형식의 내용을 미리보기하는 위젯입니다.
"""
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTextEdit, 
                            QScrollArea, QPushButton, QStackedWidget, QTableWidget,
                            QTableWidgetItem, QTabWidget, QSpinBox, QFrame, QComboBox)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt6.QtGui import QFont, QPixmap, QTextCursor
import os
from typing import Optional, Dict, Any
import config
from utils.file_manager import FileManager


class FileLoadWorker(QThread):
    """
    파일 로딩을 백그라운드에서 처리하는 워커 스레드입니다.
    """
    
    # 신호 정의
    load_completed = pyqtSignal(dict)  # 로딩 완료 시 파일 정보 전달
    load_error = pyqtSignal(str)       # 오류 발생 시 메시지 전달
    
    def __init__(self, file_path: str, file_manager: FileManager):
        super().__init__()
        self.file_path = file_path
        self.file_manager = file_manager
    
    def run(self):
        """파일 로딩을 실행합니다."""
        try:
            # 파일 정보 조회
            file_info = self.file_manager.get_file_info(self.file_path)
            
            if not file_info.get('supported', False):
                self.load_error.emit("지원되지 않는 파일 형식입니다.")
                return
            
            # FileManager의 get_file_type() 결과를 사용 (text, pdf, word 등)
            file_type = self.file_manager.get_file_type(self.file_path)
            if file_type:  # None이 아닌 경우에만 덮어쓰기
                file_info['file_type'] = file_type
            
            # 파일 타입별 추가 데이터 로딩
            if file_type == 'pdf':
                file_info['preview'] = self.file_manager.get_preview_data(self.file_path, page=0)
                file_info['text_sample'] = self.file_manager.extract_text(self.file_path, max_pages=1)
            
            elif file_type == 'image':
                # 이미지는 파일 정보에 이미 포함됨
                pass
            
            elif file_type == 'excel':
                file_info['preview'] = self.file_manager.get_preview_data(self.file_path)
                
            elif file_type == 'word':
                file_info['preview'] = self.file_manager.get_preview_data(self.file_path)
                file_info['text_sample'] = self.file_manager.extract_text(self.file_path)[:1000]
            
            elif file_type == 'powerpoint':
                file_info['preview'] = self.file_manager.get_preview_data(self.file_path, slide=0)
                file_info['text_sample'] = self.file_manager.extract_text(self.file_path)[:1000]
            
            elif file_type in ['text', 'Plain Text', 'Markdown', 'Log File', 'Text File']:
                # 텍스트 파일의 경우 미리보기 준비
                text_handler = self.file_manager.handlers['text']
                file_info['text_sample'] = text_handler.get_preview(self.file_path, max_lines=10)
                file_info.update(text_handler.get_metadata(self.file_path))
            
            self.load_completed.emit(file_info)
            
        except Exception as e:
            self.load_error.emit(f"파일 로딩 오류: {str(e)}")


class ContentViewer(QWidget):
    """
    콘텐츠 뷰어 위젯 클래스입니다.
    
    파일 형식에 따라 적절한 미리보기를 제공합니다.
    """
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_manager = FileManager()
        self.current_file_path = ""
        self.current_file_info = {}
        self.load_worker = None
        self.setup_ui()
    
    def setup_ui(self):
        """UI 구성 요소를 설정합니다."""
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        # 상단 정보 패널
        self.info_frame = QFrame()
        info_layout = QVBoxLayout()
        self.info_frame.setLayout(info_layout)
        
        # 상단 헤더 (파일명 + 원본 열기 버튼)
        header_layout = QHBoxLayout()
        
        # 파일명과 기본 정보 (왼쪽)
        title_info_layout = QVBoxLayout()
        self.title_label = QLabel("파일을 선택하세요")
        self.title_label.setFont(QFont(config.UI_FONTS["font_family"], 
                                     config.UI_FONTS["subtitle_size"], 
                                     QFont.Weight.Bold))
        self.title_label.setStyleSheet(f"color: {config.UI_COLORS['primary']};")
        
        self.details_label = QLabel("")
        self.details_label.setStyleSheet(f"color: {config.UI_COLORS['text']};")
        
        title_info_layout.addWidget(self.title_label)
        title_info_layout.addWidget(self.details_label)
        
        # 파일 작업 버튼들 (오른쪽 상단)
        buttons_layout = QHBoxLayout()
        
        # 폴더 열기 버튼
        self.open_folder_button = QPushButton("📁 폴더 열기")
        self.open_folder_button.setFont(QFont(config.UI_FONTS["font_family"], 10))
        self.open_folder_button.setFixedSize(100, 35)
        self.open_folder_button.setStyleSheet("""
            QPushButton {
                background-color: #FF9800;
                color: white;
                border: none;
                border-radius: 5px;
                font-weight: bold;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #F57C00;
            }
            QPushButton:pressed {
                background-color: #E65100;
            }
            QPushButton:disabled {
                background-color: #CCCCCC;
                color: #666666;
            }
        """)
        self.open_folder_button.clicked.connect(self.open_folder_location)
        self.open_folder_button.hide()  # 기본적으로 숨김 (파일 선택 시 표시)
        buttons_layout.addWidget(self.open_folder_button)
        
        # 원본 열기 버튼
        self.open_file_button = QPushButton("📂 원본 열기")
        self.open_file_button.setFont(QFont(config.UI_FONTS["font_family"], 10))
        self.open_file_button.setFixedSize(100, 35)
        self.open_file_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                border-radius: 5px;
                font-weight: bold;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
            QPushButton:disabled {
                background-color: #CCCCCC;
                color: #666666;
            }
        """)
        self.open_file_button.clicked.connect(self.open_original_file)
        self.open_file_button.hide()  # 기본적으로 숨김 (파일 선택 시 표시)
        buttons_layout.addWidget(self.open_file_button)
        
        header_layout.addLayout(title_info_layout)
        header_layout.addStretch()  # 공간 확보
        header_layout.addLayout(buttons_layout)
        
        info_layout.addLayout(header_layout)
        
        layout.addWidget(self.info_frame)
        
        # 메인 콘텐츠 영역 (스택 위젯)
        self.content_stack = QStackedWidget()
        
        # 1. 빈 상태 페이지
        self.empty_page = QLabel("📄\\n\\n파일을 선택하면 여기에 미리보기가 표시됩니다.")
        self.empty_page.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.empty_page.setStyleSheet(f"""
            QLabel {{
                color: {config.UI_COLORS['secondary']};
                font-size: {config.UI_FONTS['title_size']}px;
            }}
        """)
        self.content_stack.addWidget(self.empty_page)
        
        # 2. 개선된 로딩 페이지 (사용자 요청: 명확한 로딩 상태 표시)
        self.loading_page = QWidget()
        loading_layout = QVBoxLayout()
        self.loading_page.setLayout(loading_layout)
        
        # 로딩 스피너 및 메시지
        loading_container = QVBoxLayout()
        loading_container.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # 로딩 아이콘과 메시지
        self.loading_icon = QLabel("🔄")
        self.loading_icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.loading_icon.setStyleSheet(f"""
            QLabel {{
                color: {config.UI_COLORS['accent']};
                font-size: 48px;
                margin: 20px;
            }}
        """)
        
        self.loading_text = QLabel("파일을 로딩 중입니다...")
        self.loading_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.loading_text.setStyleSheet(f"""
            QLabel {{
                color: {config.UI_COLORS['accent']};
                font-size: {config.UI_FONTS['title_size']}px;
                font-weight: bold;
                margin: 10px;
            }}
        """)
        
        self.loading_file_name = QLabel("")
        self.loading_file_name.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.loading_file_name.setStyleSheet(f"""
            QLabel {{
                color: {config.UI_COLORS['text']};
                font-size: {config.UI_FONTS['body_size']}px;
                margin: 5px;
            }}
        """)
        
        loading_container.addWidget(self.loading_icon)
        loading_container.addWidget(self.loading_text)
        loading_container.addWidget(self.loading_file_name)
        
        loading_layout.addStretch()
        loading_layout.addLayout(loading_container)
        loading_layout.addStretch()
        
        self.content_stack.addWidget(self.loading_page)
        
        # 3. 텍스트 뷰어 페이지
        self.text_viewer = QTextEdit()
        self.text_viewer.setReadOnly(True)
        self.text_viewer.setStyleSheet(f"""
            QTextEdit {{
                background-color: white;
                border: 1px solid {config.UI_COLORS['secondary']};
                font-family: 'Consolas', 'Monaco', monospace;
                font-size: {config.UI_FONTS['body_size']}px;
                line-height: 1.4;
            }}
        """)
        self.content_stack.addWidget(self.text_viewer)
        
        # 4. 이미지 뷰어 페이지
        self.image_viewer = QScrollArea()
        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.image_label.setStyleSheet("background-color: white;")
        self.image_viewer.setWidget(self.image_label)
        self.image_viewer.setWidgetResizable(True)
        self.content_stack.addWidget(self.image_viewer)
        
        # 5. 테이블 뷰어 페이지 (Excel)
        self.table_viewer = QTableWidget()
        self.table_viewer.setAlternatingRowColors(True)
        self.table_viewer.setStyleSheet(f"""
            QTableWidget {{
                background-color: white;
                alternate-background-color: #F8F9FA;
                border: 1px solid {config.UI_COLORS['secondary']};
                gridline-color: {config.UI_COLORS['secondary']};
            }}
            QHeaderView::section {{
                background-color: {config.UI_COLORS['secondary']};
                color: {config.UI_COLORS['text']};
                padding: 6px;
                border: 1px solid {config.UI_COLORS['primary']};
                font-weight: bold;
            }}
        """)
        self.content_stack.addWidget(self.table_viewer)
        
        # 6. 문서 뷰어 페이지 (원본 + 텍스트 탭)
        self.document_viewer = QTabWidget()
        
        # 원본 탭 (PDF 렌더링, Word/PPT 이미지)
        self.original_tab = QScrollArea()
        self.original_label = QLabel()
        self.original_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.original_label.setStyleSheet("background-color: white;")
        self.original_tab.setWidget(self.original_label)
        self.original_tab.setWidgetResizable(True)
        self.document_viewer.addTab(self.original_tab, "📄 원본")
        
        # 텍스트 탭
        self.doc_text_viewer = QTextEdit()
        self.doc_text_viewer.setReadOnly(True)
        self.doc_text_viewer.setStyleSheet(f"""
            QTextEdit {{
                background-color: white;
                border: 1px solid {config.UI_COLORS['secondary']};
                font-family: 'Consolas', 'Monaco', monospace;
                font-size: {config.UI_FONTS['body_size']}px;
                line-height: 1.4;
            }}
        """)
        self.document_viewer.addTab(self.doc_text_viewer, "📝 텍스트")
        self.content_stack.addWidget(self.document_viewer)
        
        # 7. 오류 페이지
        self.error_page = QLabel("❌\\n\\n파일을 로딩할 수 없습니다.")
        self.error_page.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.error_page.setStyleSheet(f"""
            QLabel {{
                color: #E74C3C;
                font-size: {config.UI_FONTS['title_size']}px;
            }}
        """)
        self.content_stack.addWidget(self.error_page)
        
        layout.addWidget(self.content_stack)
        
        # 하단 컨트롤 패널
        self.control_frame = QFrame()
        control_layout = QHBoxLayout()
        self.control_frame.setLayout(control_layout)
        
        # 페이지 네비게이션 (PDF, PowerPoint용)
        self.page_label = QLabel("페이지:")
        control_layout.addWidget(self.page_label)
        
        self.page_spin = QSpinBox()
        self.page_spin.setMinimum(1)
        self.page_spin.valueChanged.connect(self.on_page_changed)
        control_layout.addWidget(self.page_spin)
        
        self.page_total_label = QLabel("/ 1")
        control_layout.addWidget(self.page_total_label)
        
        control_layout.addStretch()
        
        # 시트 선택 (Excel용)
        self.sheet_label = QLabel("시트:")
        control_layout.addWidget(self.sheet_label)
        
        self.sheet_combo = QComboBox()
        self.sheet_combo.currentTextChanged.connect(self.on_sheet_changed)
        control_layout.addWidget(self.sheet_combo)
        
        layout.addWidget(self.control_frame)
        
        # 초기에는 컨트롤 패널 숨김
        self.control_frame.hide()
        
        # 기본 페이지 표시
        self.content_stack.setCurrentWidget(self.empty_page)
    
    def load_file(self, file_path: str):
        """
        파일을 로딩합니다.
        
        Args:
            file_path (str): 로딩할 파일 경로
        """
        if not os.path.exists(file_path):
            self.show_error("파일을 찾을 수 없습니다.")
            return
        
        # 기존 PowerPoint 연결이 있다면 정리 (다른 파일 선택 시)
        if hasattr(self, 'current_file_path') and self.current_file_path and self.current_file_path != file_path:
            self.cleanup_powerpoint_connection()
        
        self.current_file_path = file_path
        
        # 로딩 페이지 표시 (개선된 로딩 메시지)
        filename = os.path.basename(file_path)
        self.loading_text.setText("파일을 로딩 중입니다...")
        self.loading_file_name.setText(f"📄 {filename}")
        self.content_stack.setCurrentWidget(self.loading_page)
        self.control_frame.hide()
        # 로딩 시작 시 버튼들 숨김
        self.open_file_button.hide()
        self.open_folder_button.hide()
        
        # 기존 워커가 있으면 정리
        if self.load_worker:
            self.load_worker.quit()
            self.load_worker.wait()
        
        # 새 워커 시작
        self.load_worker = FileLoadWorker(file_path, self.file_manager)
        self.load_worker.load_completed.connect(self.on_file_loaded)
        self.load_worker.load_error.connect(self.show_error)
        self.load_worker.start()
    
    def on_file_loaded(self, file_info: Dict[str, Any]):
        """파일 로딩 완료 시 호출됩니다."""
        self.current_file_info = file_info
        
        # 파일 정보 표시
        self.title_label.setText(f"📄 {file_info['filename']}")
        
        details = f"크기: {file_info['file_size_mb']} MB | 형식: {file_info['file_type'].upper()}"
        if 'page_count' in file_info:
            details += f" | 페이지: {file_info['page_count']}"
        elif 'sheet_count' in file_info:
            details += f" | 시트: {file_info['sheet_count']}"
        
        self.details_label.setText(details)
        
        # 파일 로딩 완료 시 버튼들 표시
        self.open_file_button.show()
        self.open_folder_button.show()
        
        # 파일 타입별 뷰어 설정
        file_type = file_info['file_type']
        
        if file_type == 'pdf':
            self.setup_pdf_viewer(file_info)
        elif file_type == 'image':
            self.setup_image_viewer(file_info)
        elif file_type == 'excel':
            self.setup_excel_viewer(file_info)
        elif file_type in ['word', 'powerpoint']:
            self.setup_document_viewer(file_info)
        elif file_type in ['text', 'Plain Text', 'Markdown', 'Log File', 'Text File']:
            self.setup_text_file_viewer(file_info)
        else:
            self.show_error(f"지원되지 않는 파일 형식입니다. (파일 타입: {file_type})")
    
    def setup_pdf_viewer(self, file_info: Dict[str, Any]):
        """PDF 뷰어를 설정합니다."""
        # 원본 PDF 렌더링
        self.render_pdf_page(self.current_file_path, 0)
        
        # 텍스트 탭 설정
        text_content = file_info.get('text_sample', '')
        if text_content and not text_content.startswith('텍스트 추출 오류'):
            self.doc_text_viewer.setPlainText(text_content)
        else:
            # 전체 텍스트 추출 시도
            full_text = self.file_manager.extract_text(self.current_file_path)
            self.doc_text_viewer.setPlainText(full_text)
        
        # 페이지 네비게이션 설정
        page_count = file_info.get('page_count', 1)
        if page_count > 1:
            self.page_spin.setMaximum(page_count)
            self.page_total_label.setText(f"/ {page_count}")
            self.page_label.show()
            self.page_spin.show()
            self.page_total_label.show()
            self.control_frame.show()
        
        # 시트 컨트롤 숨김
        self.sheet_label.hide()
        self.sheet_combo.hide()
        
        self.content_stack.setCurrentWidget(self.document_viewer)
    
    def render_pdf_page(self, file_path: str, page_num: int = 0):
        """PDF 페이지를 이미지로 렌더링합니다."""
        try:
            pdf_handler = self.file_manager.handlers['pdf']
            image = pdf_handler.render_page_to_image(file_path, page_num, zoom=1.5)
            
            if image:
                # PIL Image를 QPixmap으로 변환
                import io
                buffer = io.BytesIO()
                image.save(buffer, format='PNG')
                buffer.seek(0)
                
                pixmap = QPixmap()
                pixmap.loadFromData(buffer.getvalue())
                
                # 화면에 맞게 크기 조정
                max_width = 800
                if pixmap.width() > max_width:
                    pixmap = pixmap.scaledToWidth(max_width, Qt.TransformationMode.SmoothTransformation)
                
                self.original_label.setPixmap(pixmap)
            else:
                self.original_label.setText("PDF 렌더링 실패")
                
        except Exception as e:
            self.original_label.setText(f"PDF 렌더링 오류: {str(e)}")
    
    def setup_document_viewer(self, file_info: Dict[str, Any]):
        """Word/PowerPoint 문서 뷰어를 설정합니다."""
        file_type = file_info['file_type']
        
        # PowerPoint와 Word 문서 공통 처리
        if file_type == 'powerpoint':
            # PowerPoint의 경우 즉시 첫 번째 슬라이드 렌더링 시작
            slide_count = file_info.get('slide_count', 1)
            
            # 로딩 메시지 표시
            self.original_label.setText(f"""
🎯 PowerPoint 슬라이드 렌더링 중...

📄 파일명: {file_info['filename']}
📊 슬라이드 수: {slide_count}개
💾 크기: {file_info['file_size_mb']} MB

⚡ win32com을 사용한 고속 렌더링으로 곧 표시됩니다!
            """)
            
            # PowerPoint 지속 연결 시작 (사용자 제안 방식!)
            print(f"🚀 PowerPoint 파일 감지! 지속 연결 시작: {self.current_file_path}")
            ppt_handler = self.file_manager.handlers['powerpoint']
            
            if ppt_handler.open_persistent_connection(self.current_file_path):
                # 첫 번째 슬라이드 즉시 렌더링
                self.render_slide_instantly(0)
                # 성공 메시지 업데이트
                current_text = self.original_label.text()
                updated_text = current_text.replace(
                    "⚡ win32com을 사용한 고속 렌더링으로 곧 표시됩니다!",
                    "✅ PowerPoint 연결 완료! 슬라이드 즉시 렌더링 준비됨"
                )
                self.original_label.setText(updated_text)
            else:
                # 연결 실패 시 기존 방식으로 폴백
                print("⚠️ PowerPoint 지속 연결 실패 - 기존 방식으로 폴백")
                self.original_label.setText("PowerPoint 연결 실패 - 개별 렌더링으로 전환됩니다")
                # 첫 번째 슬라이드 폴백 렌더링
                self.render_individual_slide_fallback(0)
            
            # 슬라이드가 여러 개인 경우 네비게이션 컨트롤 표시
            if slide_count > 1:
                self.page_spin.setRange(1, slide_count)
                self.page_spin.setValue(1)  # 첫 번째 슬라이드로 초기화
                self.page_total_label.setText(f"/ {slide_count}")
                self.page_label.setText("슬라이드:")
                self.page_label.show()
                self.page_spin.show()
                self.page_total_label.show()
                self.control_frame.show()
            else:
                # 슬라이드가 1개인 경우 컨트롤 숨김
                self.control_frame.hide()
            
            # 첫 번째 슬라이드 텍스트 로드
            self.load_powerpoint_slide_text(1)
            
        else:
            # Word 문서의 경우
            self.original_label.setText(f"""
📄 {file_type.upper()} 문서

파일명: {file_info['filename']}
크기: {file_info['file_size_mb']} MB

텍스트 내용은 "텍스트" 탭에서 확인하실 수 있습니다.
원본 파일을 열려면 상단의 "원본 열기" 버튼을 클릭하세요.
            """)
            
            # 컨트롤 숨김
            self.control_frame.hide()
        
        # 텍스트 탭 설정 (Word/PowerPoint 공통)
        text_content = file_info.get('text_sample', '')
        if not text_content:
            text_content = self.file_manager.extract_text(self.current_file_path)
        
        self.doc_text_viewer.setPlainText(text_content)
        
        # 시트 컨트롤 숨김
        self.sheet_label.hide()
        self.sheet_combo.hide()
        
        self.content_stack.setCurrentWidget(self.document_viewer)
    
        
    def render_slide_instantly(self, slide_num: int):
        """지속 연결된 PowerPoint에서 슬라이드를 즉시 렌더링합니다. (사용자 제안 방식!)"""
        try:
            print(f"⚡ PowerPoint 즉시 렌더링: 슬라이드 {slide_num}")
            ppt_handler = self.file_manager.handlers['powerpoint']
            
            # 지속 연결된 PowerPoint에서 즉시 렌더링
            image = ppt_handler.render_slide_fast(slide_num, width=800, height=600)
            
            if image:
                print(f"✅ 즉시 렌더링 성공! 이미지 크기: {image.size}")
                # PIL Image를 QPixmap으로 변환
                import io
                buffer = io.BytesIO()
                image.save(buffer, format='PNG')
                buffer.seek(0)
                
                pixmap = QPixmap()
                success = pixmap.loadFromData(buffer.getvalue())
                
                if success and not pixmap.isNull():
                    # 화면에 맞게 크기 조정
                    max_width = 800
                    if pixmap.width() > max_width:
                        pixmap = pixmap.scaledToWidth(max_width, Qt.TransformationMode.SmoothTransformation)
                    
                    self.original_label.setPixmap(pixmap)
                    print("🖼️ 즉시 렌더링 이미지 표시 완료!")
                else:
                    print("❌ QPixmap 변환 실패")
                    self.original_label.setText("이미지 변환 실패")
            else:
                print("❌ 즉시 렌더링 실패 - 개별 렌더링으로 폴백")
                # 기존 방식으로 폴백
                self.render_individual_slide_fallback(slide_num)
                
        except Exception as e:
            print(f"❌ 즉시 렌더링 예외: {e}")
            self.render_individual_slide_fallback(slide_num)
    
    def render_individual_slide_fallback(self, slide_num: int):
        """지속 연결 실패 시 기존 방식으로 폴백 렌더링"""
        try:
            print(f"🔄 폴백 렌더링: 슬라이드 {slide_num}")
            ppt_handler = self.file_manager.handlers['powerpoint']
            image = ppt_handler.render_slide_to_image(self.current_file_path, slide_num, width=800, height=600)
            
            if image:
                print(f"✅ 폴백 렌더링 성공! 이미지 크기: {image.size}")
                # PIL Image를 QPixmap으로 변환
                import io
                buffer = io.BytesIO()
                image.save(buffer, format='PNG')
                buffer.seek(0)
                
                pixmap = QPixmap()
                success = pixmap.loadFromData(buffer.getvalue())
                
                if success and not pixmap.isNull():
                    # 화면에 맞게 크기 조정
                    max_width = 800
                    if pixmap.width() > max_width:
                        pixmap = pixmap.scaledToWidth(max_width, Qt.TransformationMode.SmoothTransformation)
                    
                    self.original_label.setPixmap(pixmap)
                    print("🖼️ 폴백 이미지 표시 완료!")
                else:
                    print("❌ QPixmap 변환 실패")
                    self.original_label.setText("이미지 변환 실패")
            else:
                print("❌ 폴백 렌더링도 실패")
                self.original_label.setText("슬라이드 렌더링 실패")
                
        except Exception as e:
            print(f"❌ 폴백 렌더링 예외: {e}")
            self.original_label.setText(f"슬라이드 렌더링 오류: {str(e)}")
    
    def cleanup_powerpoint_connection(self):
        """다른 파일 선택 시 PowerPoint 연결을 정리합니다."""
        try:
            ppt_handler = self.file_manager.handlers['powerpoint']
            ppt_handler.close_persistent_connection()
            print("PowerPoint 연결 정리 완료")
        except Exception as e:
            print(f"PowerPoint 연결 정리 오류: {e}")
    
    def closeEvent(self, event):
        """위젯 종료 시 PowerPoint 연결을 정리합니다."""
        try:
            self.cleanup_powerpoint_connection()
        except:
            pass
        super().closeEvent(event)
    
    def setup_text_file_viewer(self, file_info: Dict[str, Any]):
        """텍스트 파일 뷰어를 설정합니다."""
        text_handler = self.file_manager.handlers['text']
        content = text_handler.read_file_content(self.current_file_path)
        
        # 마크다운 파일의 경우 간단한 형식 표시
        if self.current_file_path.lower().endswith('.md'):
            self.text_viewer.setMarkdown(content)
        else:
            self.text_viewer.setPlainText(content)
        
        self.control_frame.hide()
        self.content_stack.setCurrentWidget(self.text_viewer)
    
    def setup_image_viewer(self, file_info: Dict[str, Any]):
        """이미지 뷰어를 설정합니다."""
        try:
            # 이미지 로딩 및 표시
            pixmap = QPixmap(self.current_file_path)
            
            if not pixmap.isNull():
                # 이미지 크기 조정 (최대 800x600)
                max_size = 800
                if pixmap.width() > max_size or pixmap.height() > max_size:
                    pixmap = pixmap.scaled(max_size, max_size, 
                                         Qt.AspectRatioMode.KeepAspectRatio, 
                                         Qt.TransformationMode.SmoothTransformation)
                
                self.image_label.setPixmap(pixmap)
                self.content_stack.setCurrentWidget(self.image_viewer)
            else:
                self.show_error("이미지를 로딩할 수 없습니다.")
        
        except Exception as e:
            self.show_error(f"이미지 로딩 오류: {str(e)}")
        
        self.control_frame.hide()
    
    def setup_excel_viewer(self, file_info: Dict[str, Any]):
        """Excel 뷰어를 설정합니다."""
        preview_data = file_info.get('preview', {})
        
        if 'data' in preview_data and preview_data['data']:
            # 테이블 설정
            data = preview_data['data']
            columns = preview_data['columns']
            
            self.table_viewer.setRowCount(len(data))
            self.table_viewer.setColumnCount(len(columns))
            self.table_viewer.setHorizontalHeaderLabels(columns)
            
            # 데이터 채우기
            for row_idx, row_data in enumerate(data):
                for col_idx, col_name in enumerate(columns):
                    value = str(row_data.get(col_name, ''))
                    item = QTableWidgetItem(value)
                    self.table_viewer.setItem(row_idx, col_idx, item)
            
            # 열 크기 자동 조정
            self.table_viewer.resizeColumnsToContents()
            
            # 시트 선택 설정
            sheet_names = file_info.get('sheet_names', [])
            if len(sheet_names) > 1:
                # 시그널 연결 해제 후 설정
                self.sheet_combo.currentTextChanged.disconnect()
                self.sheet_combo.clear()
                self.sheet_combo.addItems(sheet_names)
                
                # 현재 시트 선택
                current_sheet = file_info.get('current_sheet')
                if current_sheet and current_sheet in sheet_names:
                    self.sheet_combo.setCurrentText(current_sheet)
                
                # 시그널 다시 연결
                self.sheet_combo.currentTextChanged.connect(self.on_sheet_changed)
                
                self.sheet_label.show()
                self.sheet_combo.show()
                self.control_frame.show()
            
            # 페이지 컨트롤 숨김
            self.page_label.hide()
            self.page_spin.hide()
            self.page_total_label.hide()
            
            self.content_stack.setCurrentWidget(self.table_viewer)
        else:
            self.show_error("Excel 데이터를 읽을 수 없습니다.")
    
    def setup_text_viewer(self, file_info: Dict[str, Any]):
        """텍스트 뷰어를 설정합니다."""
        text_content = file_info.get('text_sample', '')
        
        if text_content:
            self.text_viewer.setPlainText(text_content)
        else:
            self.text_viewer.setPlainText(f"{file_info['file_type'].upper()} 문서\\n\\n파일명: {file_info['filename']}\\n\\n텍스트를 추출할 수 없습니다.")
        
        # PowerPoint의 경우 슬라이드 네비게이션
        if file_info['file_type'] == 'powerpoint':
            slide_count = file_info.get('slide_count', 1)
            if slide_count > 1:
                self.page_spin.setMaximum(slide_count)
                self.page_total_label.setText(f"/ {slide_count}")
                self.page_label.setText("슬라이드:")
                self.page_label.show()
                self.page_spin.show()
                self.page_total_label.show()
                self.control_frame.show()
        
        # 시트 컨트롤 숨김
        self.sheet_label.hide()
        self.sheet_combo.hide()
        
        self.content_stack.setCurrentWidget(self.text_viewer)
    
    def show_error(self, message: str):
        """오류 메시지를 표시합니다."""
        self.error_page.setText(f"❌\\n\\n{message}")
        self.content_stack.setCurrentWidget(self.error_page)
        self.control_frame.hide()
        
        self.title_label.setText("오류")
        self.details_label.setText(message)
    
    def on_page_changed(self, page_num: int):
        """페이지 변경 시 호출됩니다."""
        if not self.current_file_path or not self.current_file_info:
            return
        
        file_type = self.current_file_info.get('file_type')
        
        if file_type == 'pdf':
            # PDF 페이지 변경 - 원본 이미지 렌더링
            self.render_pdf_page(self.current_file_path, page_num - 1)
            
            # 해당 페이지의 텍스트도 업데이트
            pdf_handler = self.file_manager.handlers['pdf']
            try:
                import fitz
                with fitz.open(self.current_file_path) as doc:
                    if page_num - 1 < len(doc):
                        page = doc[page_num - 1]
                        page_text = page.get_text()
                        self.doc_text_viewer.setPlainText(f"=== 페이지 {page_num} ===\n\n{page_text}")
            except Exception as e:
                self.doc_text_viewer.setPlainText(f"페이지 {page_num} 텍스트 로딩 오류: {str(e)}")
        
        elif file_type == 'powerpoint':
            # PowerPoint 슬라이드 변경 시 즉시 렌더링 (지속 연결 방식)
            print(f"🔄 PowerPoint 슬라이드 변경: {page_num} (즉시 렌더링)")
            
            # 연결 상태 확인 후 적절한 렌더링 방식 선택
            ppt_handler = self.file_manager.handlers['powerpoint']
            if ppt_handler.is_connected():
                self.render_slide_instantly(page_num - 1)  # 0부터 시작
            else:
                print("⚠️ PowerPoint 연결 끊어짐 - 폴백 렌더링")
                self.render_individual_slide_fallback(page_num - 1)
                
            self.load_powerpoint_slide_text(page_num)
    
    def open_original_file(self):
        """원본 파일을 기본 프로그램으로 엽니다."""
        if not self.current_file_path:
            return
        
        try:
            import subprocess
            import sys
            import os
            
            if sys.platform == "win32":
                # Windows에서는 os.startfile 사용
                os.startfile(self.current_file_path)
            elif sys.platform == "darwin":
                # macOS에서는 open 명령 사용
                subprocess.call(["open", self.current_file_path])
            else:
                # Linux에서는 xdg-open 사용
                subprocess.call(["xdg-open", self.current_file_path])
                
        except Exception as e:
            print(f"❌ 파일 열기 실패: {e}")
    
    def open_folder_location(self):
        """선택된 파일이 있는 폴더를 엽니다."""
        if not self.current_file_path or not os.path.exists(self.current_file_path):
            print(f"❌ 폴더 열기 실패: 파일 경로가 없거나 존재하지 않습니다. {self.current_file_path}")
            return
        
        try:
            import subprocess
            import sys
            
            # 절대 경로로 변환
            file_path = os.path.abspath(self.current_file_path)
            folder_path = os.path.dirname(file_path)
            
            print(f"📁 파일 경로: {file_path}")
            print(f"📂 폴더 경로: {folder_path}")
            
            if sys.platform == "win32":
                # Windows에서는 explorer의 /select 옵션을 사용하여 파일을 선택한 상태로 폴더 열기
                file_path_normalized = os.path.normpath(file_path)
                subprocess.run(['explorer', '/select,', file_path_normalized])
                print(f"✅ Windows 폴더 열기 성공: {folder_path}")
            elif sys.platform == "darwin":
                # macOS에서는 open 명령 사용
                subprocess.call(["open", folder_path])
                print(f"✅ macOS 폴더 열기 성공: {folder_path}")
            else:
                # Linux에서는 xdg-open 사용
                subprocess.call(["xdg-open", folder_path])
                print(f"✅ Linux 폴더 열기 성공: {folder_path}")
            
        except Exception as e:
            print(f"❌ 폴더 열기 실패: {e}")
            print(f"❌ 파일 경로: {self.current_file_path}")
            print(f"❌ 폴더 경로: {os.path.dirname(self.current_file_path)}")
    
    def load_powerpoint_slide_text(self, slide_num: int):
        """PowerPoint 슬라이드의 텍스트를 로드합니다."""
        try:
            ppt_handler = self.file_manager.handlers['powerpoint']
            slide_data = ppt_handler.extract_text_from_slide(self.current_file_path, slide_num - 1)
            if slide_data and 'full_text' in slide_data:
                self.doc_text_viewer.setPlainText(f"=== 슬라이드 {slide_num} ===\n\n{slide_data['full_text']}")
            else:
                self.doc_text_viewer.setPlainText(f"슬라이드 {slide_num} 텍스트 로딩 오류")
        except Exception as e:
            self.doc_text_viewer.setPlainText(f"슬라이드 {slide_num} 텍스트 로딩 오류: {str(e)}")
    
    def on_sheet_changed(self, sheet_name: str):
        """시트 변경 시 호출됩니다."""
        if not self.current_file_path or not sheet_name:
            return
        
        # 현재 시트와 같으면 무시 (무한 루프 방지)
        if self.current_file_info.get('current_sheet') == sheet_name:
            return
        
        try:
            # Excel 시트 변경 - 직접 엑셀 핸들러 사용
            excel_handler = self.file_manager.handlers['excel']
            preview_data = excel_handler.get_preview_data(self.current_file_path, sheet_name=sheet_name)
            
            if preview_data and 'data' in preview_data:
                self.current_file_info['preview'] = preview_data
                self.current_file_info['current_sheet'] = sheet_name
                
                # 테이블만 업데이트 (시트 콤보박스는 건드리지 않음)
                self.update_excel_table(preview_data)
            else:
                self.show_error(f"시트 '{sheet_name}' 로딩 실패")
                
        except Exception as e:
            self.show_error(f"시트 변경 오류: {str(e)}")
    
    def update_excel_table(self, preview_data: Dict[str, Any]):
        """Excel 테이블만 업데이트합니다."""
        try:
            if 'data' in preview_data and preview_data['data']:
                # 테이블 설정
                data = preview_data['data']
                columns = preview_data['columns']
                
                self.table_viewer.setRowCount(len(data))
                self.table_viewer.setColumnCount(len(columns))
                self.table_viewer.setHorizontalHeaderLabels(columns)
                
                # 데이터 채우기
                for row_idx, row_data in enumerate(data):
                    for col_idx, col_name in enumerate(columns):
                        value = str(row_data.get(col_name, ''))
                        item = QTableWidgetItem(value)
                        self.table_viewer.setItem(row_idx, col_idx, item)
                
                # 열 크기 자동 조정
                self.table_viewer.resizeColumnsToContents()
            else:
                self.table_viewer.setRowCount(0)
                self.table_viewer.setColumnCount(0)
        except Exception as e:
            print(f"테이블 업데이트 오류: {e}")
    
    def clear(self):
        """뷰어를 초기화합니다."""
        self.current_file_path = ""
        self.current_file_info = {}
        self.content_stack.setCurrentWidget(self.empty_page)
        self.control_frame.hide()
        self.title_label.setText("파일을 선택하세요")
        self.details_label.setText("")