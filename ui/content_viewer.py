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
            
            file_type = file_info.get('file_type')
            
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
        
        # 파일명과 기본 정보
        self.title_label = QLabel("파일을 선택하세요")
        self.title_label.setFont(QFont(config.UI_FONTS["font_family"], 
                                     config.UI_FONTS["subtitle_size"], 
                                     QFont.Weight.Bold))
        self.title_label.setStyleSheet(f"color: {config.UI_COLORS['primary']};")
        info_layout.addWidget(self.title_label)
        
        self.details_label = QLabel("")
        self.details_label.setStyleSheet(f"color: {config.UI_COLORS['text']};")
        info_layout.addWidget(self.details_label)
        
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
        
        # 2. 로딩 페이지
        self.loading_page = QLabel("⏳\\n\\n파일을 로딩 중입니다...")
        self.loading_page.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.loading_page.setStyleSheet(f"""
            QLabel {{
                color: {config.UI_COLORS['accent']};
                font-size: {config.UI_FONTS['title_size']}px;
            }}
        """)
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
        
        # 6. 오류 페이지
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
        
        self.current_file_path = file_path
        
        # 로딩 페이지 표시
        self.content_stack.setCurrentWidget(self.loading_page)
        self.control_frame.hide()
        
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
        
        # 파일 타입별 뷰어 설정
        file_type = file_info['file_type']
        
        if file_type == 'pdf':
            self.setup_pdf_viewer(file_info)
        elif file_type == 'image':
            self.setup_image_viewer(file_info)
        elif file_type == 'excel':
            self.setup_excel_viewer(file_info)
        elif file_type in ['word', 'powerpoint']:
            self.setup_text_viewer(file_info)
        else:
            self.show_error("지원되지 않는 파일 형식입니다.")
    
    def setup_pdf_viewer(self, file_info: Dict[str, Any]):
        """PDF 뷰어를 설정합니다."""
        text_content = file_info.get('text_sample', '')
        
        if text_content and not text_content.startswith('PDF'):
            self.text_viewer.setPlainText(text_content)
        else:
            self.text_viewer.setPlainText(f"PDF 문서\\n\\n파일명: {file_info['filename']}\\n페이지 수: {file_info.get('page_count', 'N/A')}\\n\\n텍스트 추출이 제한적일 수 있습니다.")
        
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
                self.sheet_combo.clear()
                self.sheet_combo.addItems(sheet_names)
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
            # PDF 페이지 변경 (실제 구현 시 PDF 핸들러 사용)
            text_content = self.file_manager.extract_text(self.current_file_path, max_pages=1)
            self.text_viewer.setPlainText(f"PDF 페이지 {page_num}\\n\\n{text_content}")
        
        elif file_type == 'powerpoint':
            # PowerPoint 슬라이드 변경
            preview_data = self.file_manager.get_preview_data(self.current_file_path, slide=page_num-1)
            if 'full_text' in preview_data:
                self.text_viewer.setPlainText(preview_data['full_text'])
    
    def on_sheet_changed(self, sheet_name: str):
        """시트 변경 시 호출됩니다."""
        if not self.current_file_path or not sheet_name:
            return
        
        # Excel 시트 변경
        preview_data = self.file_manager.get_preview_data(self.current_file_path, sheet_name=sheet_name)
        self.current_file_info['preview'] = preview_data
        self.setup_excel_viewer(self.current_file_info)
    
    def clear(self):
        """뷰어를 초기화합니다."""
        self.current_file_path = ""
        self.current_file_info = {}
        self.content_stack.setCurrentWidget(self.empty_page)
        self.control_frame.hide()
        self.title_label.setText("파일을 선택하세요")
        self.details_label.setText("")