# -*- coding: utf-8 -*-
"""
검색 위젯 (Search Widget)

파일 내용 검색을 위한 UI 위젯입니다.
"""
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLineEdit, 
                            QPushButton, QListWidget, QListWidgetItem, QLabel,
                            QProgressBar, QFrame, QSplitter, QTextEdit, QComboBox, QMessageBox, QApplication)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt6.QtGui import QFont
import os
from typing import List, Dict, Any
import config
from utils.search_indexer import SearchIndexer


class IndexingWorker(QThread):
    """
    백그라운드에서 인덱싱을 수행하는 워커 스레드입니다.
    """
    
    # 신호 정의
    progress_updated = pyqtSignal(str, float)  # 파일 경로, 진행률
    indexing_finished = pyqtSignal(int)        # 인덱싱된 파일 수
    
    def __init__(self, indexer: SearchIndexer, directory_path: str):
        super().__init__()
        self.indexer = indexer
        self.directory_path = directory_path
    
    def run(self):
        """인덱싱을 실행합니다."""
        def progress_callback(file_path: str, progress: float):
            self.progress_updated.emit(file_path, progress)
        
        initial_count = len(self.indexer.indexed_paths)
        self.indexer.index_directory(self.directory_path, recursive=True, 
                                   progress_callback=progress_callback)
        final_count = len(self.indexer.indexed_paths)
        
        self.indexing_finished.emit(final_count - initial_count)


class SearchWidget(QWidget):
    """
    검색 위젯 클래스입니다.
    
    파일 내용 검색 및 결과 표시 기능을 제공합니다.
    """
    
    # 파일 선택 시 발생하는 신호
    file_selected = pyqtSignal(str)  # 파일 경로
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.indexer = SearchIndexer()
        self.indexing_worker = None
        self.current_directory = ""
        self.current_selected_file = None  # 현재 선택된 파일 경로
        self.search_mode = "content"  # "content" 또는 "filename"
        self.setup_ui()
        
        # 자동 검색 제거 (사용자 요청: 검색 버튼과 엔터키만 사용)
    
    def setup_ui(self):
        """UI 구성 요소를 설정합니다."""
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        # 상단 검색 영역
        search_frame = QFrame()
        search_layout = QVBoxLayout()
        search_frame.setLayout(search_layout)
        
        # 검색 입력
        search_input_layout = QHBoxLayout()
        
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("검색어 입력... (2글자 이상)")
        self.search_input.textChanged.connect(self.on_search_text_changed)
        self.search_input.returnPressed.connect(self.perform_search)
        search_input_layout.addWidget(self.search_input)
        
        self.search_button = QPushButton("🔍 검색")
        self.search_button.clicked.connect(self.perform_search)
        search_input_layout.addWidget(self.search_button)
        
        search_layout.addLayout(search_input_layout)
        
        # 검색 모드 선택
        search_mode_layout = QHBoxLayout()
        
        mode_label = QLabel("검색 모드:")
        search_mode_layout.addWidget(mode_label)
        
        self.search_mode_combo = QComboBox()
        self.search_mode_combo.addItems(["📄 파일 내용 검색", "📝 파일명 검색"])
        self.search_mode_combo.setCurrentIndex(0)  # 기본은 파일 내용 검색
        self.search_mode_combo.currentTextChanged.connect(self.on_search_mode_changed)
        search_mode_layout.addWidget(self.search_mode_combo)
        
        search_mode_layout.addStretch()
        
        search_layout.addLayout(search_mode_layout)
        
        # 인덱싱 컨트롤
        indexing_layout = QHBoxLayout()
        
        self.index_button = QPushButton("📂 폴더 인덱싱")
        self.index_button.clicked.connect(self.start_indexing)
        indexing_layout.addWidget(self.index_button)
        
        self.clear_index_button = QPushButton("🧹 인덱스 초기화")
        self.clear_index_button.clicked.connect(self.clear_index)
        indexing_layout.addWidget(self.clear_index_button)
        
        indexing_layout.addStretch()
        
        self.index_stats_label = QLabel("인덱스: 0개 파일")
        indexing_layout.addWidget(self.index_stats_label)
        
        search_layout.addLayout(indexing_layout)
        
        # 인덱싱 대상 파일 확장자 표시
        self.indexed_extensions_label = QLabel("인덱싱 대상: .pdf .ppt .pptx .doc .docx .txt (※ Excel 제외)")
        self.indexed_extensions_label.setStyleSheet(f"""
            QLabel {{
                color: {config.UI_COLORS['text']};
                font-size: {config.UI_FONTS['small_size']}px;
                font-style: italic;
                padding: 2px;
                background-color: {config.UI_COLORS['background']};
            }}
        """)
        search_layout.addWidget(self.indexed_extensions_label)
        
        # 진행률 표시
        self.progress_bar = QProgressBar()
        self.progress_bar.hide()
        search_layout.addWidget(self.progress_bar)
        
        self.progress_label = QLabel("")
        self.progress_label.hide()
        search_layout.addWidget(self.progress_label)
        
        layout.addWidget(search_frame)
        
        # 검색 결과 영역
        results_splitter = QSplitter(Qt.Orientation.Vertical)
        
        # 결과 목록
        results_frame = QFrame()
        results_layout = QVBoxLayout()
        results_frame.setLayout(results_layout)
        
        self.results_label = QLabel("검색 결과")
        self.results_label.setFont(QFont(config.UI_FONTS["font_family"], 
                                       config.UI_FONTS["subtitle_size"], 
                                       QFont.Weight.Bold))
        results_layout.addWidget(self.results_label)
        
        self.results_list = QListWidget()
        self.results_list.itemClicked.connect(self.on_result_selected)
        self.results_list.setMinimumHeight(200)
        results_layout.addWidget(self.results_list)
        
        results_splitter.addWidget(results_frame)
        
        # 파일 작업 영역
        actions_frame = QFrame()
        actions_layout = QHBoxLayout()
        actions_frame.setLayout(actions_layout)
        
        actions_layout.addStretch()
        
        # 파일 뷰어에서 열기 버튼
        self.open_viewer_button = QPushButton("파일 뷰어에서 열기")
        self.open_viewer_button.setFixedSize(140, 35)
        self.open_viewer_button.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                border-radius: 5px;
                font-weight: bold;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:pressed {
                background-color: #0D47A1;
            }
            QPushButton:disabled {
                background-color: #CCCCCC;
                color: #666666;
            }
        """)
        self.open_viewer_button.clicked.connect(self.open_in_viewer)
        self.open_viewer_button.setEnabled(False)
        actions_layout.addWidget(self.open_viewer_button)
        
        # 폴더 열기 버튼
        self.open_folder_button = QPushButton("📁 폴더 열기")
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
        self.open_folder_button.setEnabled(False)
        actions_layout.addWidget(self.open_folder_button)
        
        # 원본 열기 버튼
        self.open_original_button = QPushButton("📂 원본 열기")
        self.open_original_button.setFixedSize(100, 35)
        self.open_original_button.setStyleSheet("""
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
        self.open_original_button.clicked.connect(self.open_original_file)
        self.open_original_button.setEnabled(False)  # 기본적으로 비활성화
        actions_layout.addWidget(self.open_original_button)
        
        results_splitter.addWidget(actions_frame)
        
        # 스플리터 비율 설정
        results_splitter.setSizes([400, 50])
        
        layout.addWidget(results_splitter)
        
        self.apply_styles()
        self.update_index_stats()
    
    def apply_styles(self):
        """스타일을 적용합니다."""
        search_style = f"""
            QLineEdit {{
                padding: 8px;
                font-size: {config.UI_FONTS['body_size']}px;
                border: 2px solid {config.UI_COLORS['secondary']};
                border-radius: 4px;
            }}
            QLineEdit:focus {{
                border-color: {config.UI_COLORS['accent']};
            }}
        """
        self.search_input.setStyleSheet(search_style)
        
        button_style = f"""
            QPushButton {{
                background-color: {config.UI_COLORS['accent']};
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
                font-size: {config.UI_FONTS['body_size']}px;
            }}
            QPushButton:hover {{
                background-color: {config.UI_COLORS['hover']};
            }}
            QPushButton:pressed {{
                background-color: {config.UI_COLORS['primary']};
            }}
        """
        self.search_button.setStyleSheet(button_style)
        self.index_button.setStyleSheet(button_style)
        self.clear_index_button.setStyleSheet(button_style)
        
        # 검색 모드 드롭다운 스타일
        combo_style = f"""
            QComboBox {{
                background-color: white;
                border: 2px solid {config.UI_COLORS['secondary']};
                padding: 6px 12px;
                border-radius: 4px;
                font-weight: bold;
                font-size: {config.UI_FONTS['body_size']}px;
                color: {config.UI_COLORS['text']};
            }}
            QComboBox:hover {{
                border-color: {config.UI_COLORS['accent']};
            }}
            QComboBox::drop-down {{
                border: none;
            }}
            QComboBox::down-arrow {{
                width: 12px;
                height: 12px;
            }}
        """
        self.search_mode_combo.setStyleSheet(combo_style)
        
        list_style = f"""
            QListWidget {{
                background-color: white;
                border: 1px solid {config.UI_COLORS['secondary']};
                font-size: {config.UI_FONTS['body_size']}px;
            }}
            QListWidget::item {{
                padding: 8px;
                border-bottom: 1px solid #EEEEEE;
            }}
            QListWidget::item:hover {{
                background-color: {config.UI_COLORS['hover']};
            }}
            QListWidget::item:selected {{
                background-color: {config.UI_COLORS['accent']};
                color: white;
            }}
        """
        self.results_list.setStyleSheet(list_style)
        
        # 텍스트 스타일 (미리보기가 제거되어 더 이상 사용하지 않음)
    
    def set_directory(self, directory_path: str):
        """
        검색 대상 디렉토리를 설정합니다.
        
        Args:
            directory_path (str): 디렉토리 경로
        """
        self.current_directory = directory_path
        self.index_button.setText(f"📂 '{os.path.basename(directory_path)}' 인덱싱")
        self.index_button.setEnabled(True)
    
    def start_indexing(self):
        """인덱싱을 시작합니다."""
        if not self.current_directory or not os.path.exists(self.current_directory):
            self.results_label.setText("검색 결과 - 디렉토리를 먼저 선택해주세요")
            return
        
        if self.indexing_worker and self.indexing_worker.isRunning():
            return
        
        # UI 업데이트
        self.index_button.setEnabled(False)
        self.progress_bar.show()
        self.progress_bar.setValue(0)
        self.progress_label.show()
        self.progress_label.setText("인덱싱 준비 중...")
        
        # 워커 시작
        self.indexing_worker = IndexingWorker(self.indexer, self.current_directory)
        self.indexing_worker.progress_updated.connect(self.on_indexing_progress)
        self.indexing_worker.indexing_finished.connect(self.on_indexing_finished)
        self.indexing_worker.start()
    
    def on_indexing_progress(self, file_path: str, progress: float):
        """인덱싱 진행 상태를 업데이트합니다."""
        self.progress_bar.setValue(int(progress))
        self.progress_label.setText(f"인덱싱 중: {os.path.basename(file_path)}")
    
    def on_indexing_finished(self, indexed_count: int):
        """인덱싱 완료 시 호출됩니다."""
        self.progress_bar.hide()
        self.progress_label.hide()
        self.index_button.setEnabled(True)
        
        self.update_index_stats()
        self.results_label.setText(f"검색 결과 - {indexed_count}개 파일이 새로 인덱싱됨")
    
    def clear_index(self):
        """인덱스를 초기화합니다."""
        self.indexer.clear_index()
        self.results_list.clear()
        self.update_index_stats()
        self.results_label.setText("검색 결과 - 인덱스 초기화됨")
        
        # 버튼들 비활성화
        self.open_viewer_button.setEnabled(False)
        self.open_original_button.setEnabled(False)
        self.open_folder_button.setEnabled(False)
        self.current_selected_file = None
    
    def update_index_stats(self):
        """인덱스 통계를 업데이트합니다."""
        stats = self.indexer.get_index_statistics()
        self.index_stats_label.setText(f"인덱스: {stats['total_files']}개 파일, {stats['total_tokens']}개 토큰")
    
    def on_search_text_changed(self, text: str):
        """검색 텍스트 변경 시 호출됩니다."""
        # 자동 검색 제거 - 결과 초기화만 수행
        if len(text.strip()) < 2:
            self.results_list.clear()
            self.results_label.setText("검색 결과")
    
    def perform_search(self):
        """검색을 수행합니다."""
        query = self.search_input.text().strip()
        
        if len(query) < 2:
            self.results_label.setText("검색 결과 - 2글자 이상 입력해주세요")
            return
        
        # 🔍 조회중 상태 표시
        self.results_label.setText(f"🔍 '{query}' 조회 중...")
        self.results_list.clear()
        
        # 조회중 표시 아이템 추가
        loading_item = QListWidgetItem("⏳ 검색 중입니다...")
        loading_item.setData(Qt.ItemDataRole.UserRole, None)
        self.results_list.addItem(loading_item)
        
        # UI 업데이트 강제 실행
        QApplication.processEvents()
        
        # 검색 모드에 따라 다른 검색 수행
        if self.search_mode == "content":
            # 파일 내용 검색 - 인덱싱 완료 체크
            if not self.indexer or len(self.indexer.indexed_paths) == 0:
                QMessageBox.warning(self, "인덱싱 필요", 
                                   "파일 내용 검색을 위해서는 먼저 인덱싱을 완료해야 합니다.\n\n'📂 폴더 인덱싱' 버튼을 클릭하여 인덱싱을 시작하세요.")
                # 조회중 상태 제거
                self.results_list.clear()
                self.results_label.setText("검색 결과")
                return
            search_results = self.indexer.search_files(query, max_results=100)
        else:
            # 파일명 검색 - JSON 캐시 활용 (사용자 요청: 고속 검색)
            if hasattr(self.indexer, 'search_files_by_filename_from_json'):
                search_results = self.indexer.search_files_by_filename_from_json(query, max_results=100)
            else:
                # 폴백: 기존 방식
                search_results = self.search_by_filename(query, max_results=100)
        
        # 결과 표시 - 조회중 상태 제거
        self.results_list.clear()
        
        if not search_results:
            self.results_label.setText(f"검색 결과 - '{query}'에 대한 결과 없음")
            return
        
        self.results_label.setText(f"검색 결과 - '{query}' ({len(search_results)}개)")
        
        for result in search_results:
            item = QListWidgetItem()
            
            # 결과 항목 텍스트 구성
            filename = result['filename']
            file_type = result['file_type'].upper()
            file_size = result['file_size_mb']
            
            item_text = f"📄 {filename} ({file_type}, {file_size}MB)"
            item.setText(item_text)
            
            # 결과 데이터 저장
            item.setData(Qt.ItemDataRole.UserRole, result)
            
            self.results_list.addItem(item)
    
    def on_result_selected(self, item: QListWidgetItem):
        """검색 결과 선택 시 호출됩니다."""
        result = item.data(Qt.ItemDataRole.UserRole)
        
        if result:
            self.current_selected_file = result['file_path']
            
            # 버튼들 활성화
            self.open_viewer_button.setEnabled(True)
            self.open_original_button.setEnabled(True)
            self.open_folder_button.setEnabled(True)
    
    def add_file_to_index(self, file_path: str):
        """
        파일을 인덱스에 추가합니다.
        
        Args:
            file_path (str): 추가할 파일 경로
        """
        self.indexer.add_file_to_index(file_path)
        self.update_index_stats()
    
    def remove_file_from_index(self, file_path: str):
        """
        파일을 인덱스에서 제거합니다.
        
        Args:
            file_path (str): 제거할 파일 경로
        """
        self.indexer.remove_file_from_index(file_path)
        self.update_index_stats()
    
    def get_search_statistics(self) -> Dict[str, Any]:
        """
        검색 통계를 반환합니다.
        
        Returns:
            Dict[str, Any]: 통계 정보
        """
        return self.indexer.get_index_statistics()
    
    def open_original_file(self):
        """선택된 파일을 기본 프로그램으로 엽니다."""
        if not self.current_selected_file or not os.path.exists(self.current_selected_file):
            return
        
        try:
            import subprocess
            import sys
            
            if sys.platform == "win32":
                # Windows에서는 os.startfile 사용
                os.startfile(self.current_selected_file)
            elif sys.platform == "darwin":
                # macOS에서는 open 명령 사용
                subprocess.call(["open", self.current_selected_file])
            else:
                # Linux에서는 xdg-open 사용
                subprocess.call(["xdg-open", self.current_selected_file])
                
            print(f"✅ 원본 파일 열기: {self.current_selected_file}")
            
        except Exception as e:
            print(f"❌ 원본 파일 열기 실패: {e}")
    
    def open_folder_location(self):
        """선택된 파일이 있는 폴더를 엽니다."""
        if not self.current_selected_file or not os.path.exists(self.current_selected_file):
            print(f"❌ 폴더 열기 실패: 파일 경로가 없거나 존재하지 않습니다. {self.current_selected_file}")
            return
        
        try:
            import subprocess
            import sys
            
            # 절대 경로로 변환
            file_path = os.path.abspath(self.current_selected_file)
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
            print(f"❌ 파일 경로: {self.current_selected_file}")
            print(f"❌ 폴더 경로: {os.path.dirname(self.current_selected_file)}")
    
    def open_in_viewer(self):
        """선택된 파일을 파일 뷰어에서 엽니다."""
        if not self.current_selected_file or not os.path.exists(self.current_selected_file):
            return
        
        # 로딩 중 버튼 비활성화 (UX 개선: 중복 클릭 방지)
        self.open_viewer_button.setEnabled(False)
        
        # 로딩 알림창 표시 (제대로 된 modal dialog)
        from PyQt6.QtWidgets import QProgressDialog
        from PyQt6.QtCore import Qt
        
        self.loading_dialog = QProgressDialog("파일 로딩중입니다...", None, 0, 0, self)
        self.loading_dialog.setWindowTitle("파일 로딩 중")
        self.loading_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.loading_dialog.setAutoClose(False)
        self.loading_dialog.setAutoReset(False)
        self.loading_dialog.show()
        
        print(f"🔄 파일 뷰어에서 열기: {self.current_selected_file}")
        
        # 파일 선택 신호 발생
        self.file_selected.emit(self.current_selected_file)
    
    def close_loading_dialog(self):
        """로딩 알림창을 닫습니다."""
        if hasattr(self, 'loading_dialog') and self.loading_dialog:
            self.loading_dialog.close()
            self.loading_dialog = None
            print("✅ 파일 로딩 완료 - 알림창 닫음")
        
        # 버튼 다시 활성화 (로딩 완료 후)
        if self.current_selected_file:
            self.open_viewer_button.setEnabled(True)
    
    def on_search_mode_changed(self):
        """검색 모드 변경 시 호출됩니다."""
        current_text = self.search_mode_combo.currentText()
        
        if "파일 내용" in current_text:
            self.search_mode = "content"
            self.search_input.setPlaceholderText("파일 내용 검색... (2글자 이상)")
        elif "파일명" in current_text:
            self.search_mode = "filename"
            self.search_input.setPlaceholderText("파일명 검색... (확장자 제외, 2글자 이상)")
    
    def search_by_filename(self, query: str, max_results: int = 100):
        """
        파일명으로 검색을 수행합니다 (확장자 제외).
        
        Args:
            query (str): 검색 쿼리
            max_results (int): 최대 결과 수
            
        Returns:
            List[Dict]: 검색 결과
        """
        if not self.current_directory or not os.path.exists(self.current_directory):
            return []
        
        results = []
        query_lower = query.lower()
        
        try:
            # 현재 디렉토리에서 파일 검색
            for root, dirs, files in os.walk(self.current_directory):
                for file in files:
                    file_path = os.path.join(root, file)
                    
                    # 확장자를 제외한 파일명 추출
                    filename_without_ext = os.path.splitext(file)[0]
                    
                    # 확장자를 제외한 파일명에 검색어가 포함되어 있는지 확인
                    if query_lower in filename_without_ext.lower():
                        # 지원되는 파일만 결과에 포함
                        if self.indexer.file_manager.is_supported_file(file_path):
                            file_info = self.indexer.file_manager.get_file_info(file_path)
                            
                            if file_info.get('supported', False):
                                result = {
                                    'filename': file_info['filename'],
                                    'file_path': file_path,
                                    'file_type': file_info['file_type'],
                                    'file_size_mb': file_info['file_size_mb'],
                                    'relevance_score': 1.0,  # 파일명 매칭이므로 높은 점수
                                    'preview': f"파일명 매칭: {filename_without_ext}"
                                }
                                results.append(result)
                                
                                if len(results) >= max_results:
                                    break
                
                if len(results) >= max_results:
                    break
                    
        except Exception as e:
            print(f"❌ 파일명 검색 중 오류: {e}")
        
        # 관련성 점수로 정렬 (파일명 일치도)
        results.sort(key=lambda x: x['relevance_score'], reverse=True)
        
        return results