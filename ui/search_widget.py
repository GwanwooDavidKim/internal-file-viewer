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
        self.current_selected_result = None  # 현재 선택된 검색 결과 (matching_pages 포함)
        self.search_mode = "content"  # "content" 또는 "filename"
        
        # 🆕 검색 결과 및 정렬 상태 
        self.current_search_results = []
        self.current_sort_mode = "[정렬] 관련성 순 (기본)"
        
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
        
        # 🆕 파일명 검색 입력
        filename_search_layout = QHBoxLayout()
        
        filename_label = QLabel("[텍스트] 파일명:")
        filename_label.setMinimumWidth(60)
        filename_search_layout.addWidget(filename_label)
        
        self.filename_search_input = QLineEdit()
        self.filename_search_input.setPlaceholderText("파일명 필터 (쉼표로 구분, 예: TFT,BOE)")
        self.filename_search_input.returnPressed.connect(self.perform_search)
        filename_search_layout.addWidget(self.filename_search_input)
        
        search_layout.addLayout(filename_search_layout)
        
        # 🆕 내용 검색 입력
        content_search_layout = QHBoxLayout()
        
        content_label = QLabel("[파일] 내용:")
        content_label.setMinimumWidth(60)
        content_search_layout.addWidget(content_label)
        
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("내용 검색 (쉼표로 구분, 띄어쓰기 무시, 예: 자사,Fab,별,Capa)")
        self.search_input.textChanged.connect(self.on_search_text_changed)
        self.search_input.returnPressed.connect(self.perform_search)
        content_search_layout.addWidget(self.search_input)
        
        self.search_button = QPushButton("🔍 검색")
        self.search_button.clicked.connect(self.perform_search)
        content_search_layout.addWidget(self.search_button)
        
        search_layout.addLayout(content_search_layout)
        
        # 🆕 검색 도움말
        help_label = QLabel("💡 팁: 파일명과 내용을 동시에 입력하면 두 조건을 모두 만족하는 파일만 검색됩니다")
        help_label.setStyleSheet(f"""
            QLabel {{
                color: {config.UI_COLORS['text']};
                font-size: {config.UI_FONTS['small_size']}px;
                font-style: italic;
                padding: 5px;
                background-color: {config.UI_COLORS['background']};
            }}
        """)
        search_layout.addWidget(help_label)
        
        # 인덱싱 컨트롤
        indexing_layout = QHBoxLayout()
        
        self.index_button = QPushButton("[경로] 폴더 인덱싱")
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
        
        # 🆕 검색 결과 정렬 옵션
        sort_layout = QHBoxLayout()
        
        sort_label = QLabel("정렬 순서:")
        sort_layout.addWidget(sort_label)
        
        self.sort_combo = QComboBox()
        self.sort_combo.addItems([
            "[정렬] 관련성 순 (기본)",
            "[폴더] 파일명 (오름차순)", 
            "[폴더] 파일명 (내림차순)",
            "[날짜] 최신 변경일 순",
            "[날짜] 오래된 변경일 순",
            "📏 파일크기 (큰순)",
            "📏 파일크기 (작은순)"
        ])
        self.sort_combo.setCurrentIndex(0)
        self.sort_combo.currentTextChanged.connect(self.on_sort_changed)
        sort_layout.addWidget(self.sort_combo)
        
        sort_layout.addStretch()
        
        search_layout.addLayout(sort_layout)
        
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
        self.open_folder_button = QPushButton("[폴더] 폴더 열기")
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
        self.open_original_button = QPushButton("[경로] 원본 열기")
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
        self.index_button.setText(f"[경로] '{os.path.basename(directory_path)}' 인덱싱")
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
        """검색을 수행합니다 (2단계 필터링 지원)."""
        # 🆕 파일명 및 내용 검색어 가져오기
        filename_query = self.filename_search_input.text().strip()
        content_query = self.search_input.text().strip()
        
        # 최소 하나의 검색어는 있어야 함
        if not filename_query and not content_query:
            self.results_label.setText("검색 결과 - 파일명 또는 내용 중 하나는 입력해주세요")
            return
        
        # 검색어 표시용 텍스트 생성
        search_display = []
        if filename_query:
            search_display.append(f"파일명:{filename_query}")
        if content_query:
            search_display.append(f"내용:{content_query}")
        display_text = ", ".join(search_display)
        
        # 🔍 조회중 상태 표시
        self.results_label.setText(f"🔍 '{display_text}' 조회 중...")
        self.results_list.clear()
        
        # 조회중 표시 아이템 추가
        loading_item = QListWidgetItem("⏳ 검색 중입니다...")
        loading_item.setData(Qt.ItemDataRole.UserRole, None)
        self.results_list.addItem(loading_item)
        
        # UI 업데이트 강제 실행
        QApplication.processEvents()
        
        # 🆕 2단계 필터링 검색
        if content_query:
            # 파일 내용 검색 - 인덱싱 완료 체크
            if not self.indexer or len(self.indexer.indexed_paths) == 0:
                QMessageBox.warning(self, "인덱싱 필요", 
                                   "파일 내용 검색을 위해서는 먼저 인덱싱을 완료해야 합니다.\n\n'[경로] 폴더 인덱싱' 버튼을 클릭하여 인덱싱을 시작하세요.")
                # 조회중 상태 제거
                self.results_list.clear()
                self.results_label.setText("검색 결과")
                return
            
            # 내용으로 검색
            search_results = self.indexer.search_files(content_query, max_results=200)
            
            # 🆕 파일명 필터링 (2단계)
            if filename_query and search_results:
                search_results = self._filter_by_filename(search_results, filename_query)
        else:
            # 파일명만 검색
            if hasattr(self.indexer, 'search_files_by_filename_from_json'):
                search_results = self.indexer.search_files_by_filename_from_json(filename_query, max_results=100)
            else:
                search_results = self.search_by_filename(filename_query, max_results=100)
        
        # 🆕 검색 결과 저장 및 정렬하여 표시
        self.current_search_results = search_results
        self._display_sorted_results(display_text)
    
    def _filter_by_filename(self, results: List[Dict[str, Any]], filename_query: str) -> List[Dict[str, Any]]:
        """파일명으로 검색 결과를 필터링합니다."""
        # 다중 키워드 지원
        if ',' in filename_query:
            keywords = [kw.strip().lower() for kw in filename_query.split(',') if kw.strip()]
        else:
            keywords = [filename_query.lower()]
        
        # 공백 제거 버전 키워드
        keywords_no_space = [kw.replace(' ', '').replace('\n', '').replace('\t', '') for kw in keywords]
        
        filtered_results = []
        for result in results:
            filename = result.get('filename', '').lower()
            filename_no_space = filename.replace(' ', '').replace('\n', '').replace('\t', '')
            
            # 모든 키워드가 파일명에 포함되어야 함
            all_found = True
            for i, keyword in enumerate(keywords):
                keyword_no_space = keywords_no_space[i]
                if keyword not in filename and keyword_no_space not in filename_no_space:
                    all_found = False
                    break
            
            if all_found:
                filtered_results.append(result)
        
        return filtered_results
    
    def on_result_selected(self, item: QListWidgetItem):
        """검색 결과 선택 시 호출됩니다."""
        result = item.data(Qt.ItemDataRole.UserRole)
        
        # [항목] 헤더 항목이나 선택 불가 항목은 무시
        if result is None:
            # 버튼들 비활성화 (헤더 선택 시)
            self.open_viewer_button.setEnabled(False)
            self.open_original_button.setEnabled(False)
            self.open_folder_button.setEnabled(False)
            self.current_selected_file = None
            self.current_selected_result = None
            return
        
        if result:
            self.current_selected_file = result['file_path']
            self.current_selected_result = result  # 전체 결과 저장 (matching_pages 포함)
            
            # 버튼들 활성화
            self.open_viewer_button.setEnabled(True)
            self.open_original_button.setEnabled(True)
            self.open_folder_button.setEnabled(True)
    
    def on_sort_changed(self, sort_text: str):
        """정렬 방식 변경 시 호출됩니다."""
        self.current_sort_mode = sort_text
        if self.current_search_results:
            # 현재 검색 결과를 새로운 정렬 방식으로 다시 표시
            self._display_sorted_results(self.search_input.text().strip())
    
    def _sort_results(self, results: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """검색 결과를 현재 정렬 모드에 따라 정렬합니다."""
        if not results:
            return results
        
        sort_mode = self.current_sort_mode
        
        if "관련성" in sort_mode:
            # 기본 관련성 순 (이미 정렬되어 있음)
            return results
        elif "파일명 (오름차순)" in sort_mode:
            return sorted(results, key=lambda x: x['filename'].lower())
        elif "파일명 (내림차순)" in sort_mode:
            return sorted(results, key=lambda x: x['filename'].lower(), reverse=True)
        elif "최신 변경일" in sort_mode:
            # 파일 변경일 기준 정렬 (최신순)
            return sorted(results, key=lambda x: self._get_file_mtime(x['file_path']), reverse=True)
        elif "오래된 변경일" in sort_mode:
            # 파일 변경일 기준 정렬 (오래된순)
            return sorted(results, key=lambda x: self._get_file_mtime(x['file_path']))
        elif "파일크기 (큰순)" in sort_mode:
            return sorted(results, key=lambda x: x.get('file_size_mb', 0), reverse=True)
        elif "파일크기 (작은순)" in sort_mode:
            return sorted(results, key=lambda x: x.get('file_size_mb', 0))
        else:
            return results
    
    def _get_file_mtime(self, file_path: str) -> float:
        """파일의 수정 시간을 반환합니다."""
        try:
            import os
            return os.path.getmtime(file_path)
        except:
            return 0.0
    
    def _group_by_extension(self, results: List[Dict[str, Any]]) -> Dict[str, List[Dict[str, Any]]]:
        """검색 결과를 확장자별로 그룹핑합니다."""
        groups = {}
        
        # 확장자 우선순위 정의 (사용자 요청: ppt → pdf → txt 등 순서)
        extension_priority = {
            'ppt': 1, 'pptx': 1,
            'pdf': 2,
            'doc': 3, 'docx': 3,
            'txt': 4,
            'xls': 5, 'xlsx': 5,
            'jpg': 6, 'jpeg': 6, 'png': 6, 'gif': 6, 'bmp': 6
        }
        
        for result in results:
            ext = result.get('file_type', 'unknown').lower()
            if ext not in groups:
                groups[ext] = []
            groups[ext].append(result)
        
        # 확장자별로 정렬된 딕셔너리 반환 (우선순위 순서)
        sorted_groups = {}
        for ext in sorted(groups.keys(), key=lambda x: extension_priority.get(x, 99)):
            sorted_groups[ext] = groups[ext]
        
        return sorted_groups
    
    def _group_by_directory(self, results: List[Dict[str, Any]]) -> Dict[str, List[Dict[str, Any]]]:
        """결과를 디렉토리별로 그룹화합니다."""
        import os
        groups = {}
        
        for result in results:
            file_path = result.get('file_path', '')
            directory = os.path.dirname(file_path)
            
            # 디렉토리 이름 추출 (전체 경로 대신 마지막 폴더명만)
            if not directory:
                directory = "(루트)"
            
            if directory not in groups:
                groups[directory] = []
            groups[directory].append(result)
        
        # 디렉토리명으로 정렬
        sorted_groups = dict(sorted(groups.items()))
        return sorted_groups
    
    def _display_sorted_results(self, query: str):
        """정렬된 검색 결과를 표시합니다."""
        self.results_list.clear()
        
        if not self.current_search_results:
            self.results_label.setText(f"검색 결과 - '{query}'에 대한 결과 없음")
            return
        
        # [로딩] 정렬 수행
        sorted_results = self._sort_results(self.current_search_results)
        
        # [항목] 확장자별 그룹핑
        grouped_results = self._group_by_extension(sorted_results)
        
        total_count = len(sorted_results)
        self.results_label.setText(f"검색 결과 - '{query}' ({total_count}개) | {self.current_sort_mode}")
        
        # 그룹별로 결과 표시
        for ext, ext_results in grouped_results.items():
            # 확장자 헤더 추가
            if len(grouped_results) > 1:  # 여러 확장자가 있을 때만 헤더 표시
                header_item = QListWidgetItem()
                header_text = f"[폴더] {ext.upper()} 파일 ({len(ext_results)}개)"
                header_item.setText(header_text)
                header_item.setData(Qt.ItemDataRole.UserRole, None)  # 헤더는 선택 불가
                
                # 헤더 스타일 설정
                header_item.setBackground(QApplication.palette().alternateBase())
                self.results_list.addItem(header_item)
            
            # 🆕 디렉토리별로 다시 그룹화
            dir_groups = self._group_by_directory(ext_results)
            
            # 디렉토리별로 결과 표시
            for directory, dir_results in dir_groups.items():
                # 🆕 상대 경로 계산 (검색 루트 기준)
                if directory == "(루트)":
                    display_path = "(루트)"
                elif self.current_directory:
                    try:
                        # 검색 디렉토리 기준 상대 경로
                        rel_path = os.path.relpath(directory, self.current_directory)
                        display_path = rel_path if rel_path != "." else "(루트)"
                    except ValueError:
                        # 다른 드라이브인 경우 전체 경로 표시
                        display_path = directory
                else:
                    display_path = directory
                
                # 🆕 디렉토리 헤더 항상 표시 (경로 정보 제공)
                dir_header = QListWidgetItem()
                dir_header_text = f"  [경로] {display_path} ({len(dir_results)}개)"
                dir_header.setText(dir_header_text)
                dir_header.setData(Qt.ItemDataRole.UserRole, None)
                
                # 디렉토리 헤더 스타일
                font = dir_header.font()
                font.setBold(True)
                dir_header.setFont(font)
                # 툴팁에 전체 경로 표시
                dir_header.setToolTip(f"전체 경로: {directory}")
                self.results_list.addItem(dir_header)
                
                # 해당 디렉토리의 파일들 표시 (항상 들여쓰기)
                for result in dir_results:
                    item = QListWidgetItem()
                    
                    # 결과 항목 텍스트 구성
                    filename = result['filename']
                    file_type = result['file_type'].upper()
                    file_size = result['file_size_mb']
                    matching_pages = result.get('matching_pages', [])
                    
                    # 페이지 정보 추가
                    page_info = ""
                    if matching_pages:
                        if len(matching_pages) <= 5:
                            page_info = f" | 페이지: {', '.join(map(str, matching_pages))}"
                        else:
                            page_info = f" | 페이지: {', '.join(map(str, matching_pages[:5]))}... ({len(matching_pages)}개)"
                    
                    # 파일 아이콘과 정보 표시 (들여쓰기)
                    item_text = f"    [파일] {filename} ({file_type}, {file_size}MB){page_info}"
                    item.setText(item_text)
                    
                    # 결과 데이터 저장
                    item.setData(Qt.ItemDataRole.UserRole, result)
                    
                    # 툴팁에 전체 경로 + 전체 페이지 번호 표시
                    tooltip = f"전체 경로: {result.get('file_path', '')}"
                    if matching_pages:
                        tooltip += f"\n검색어 포함 페이지: {', '.join(map(str, matching_pages))}"
                    item.setToolTip(tooltip)
                    
                    self.results_list.addItem(item)
    
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
                
            print(f"[성공] 원본 파일 열기: {self.current_selected_file}")
            
        except Exception as e:
            print(f"[오류] 원본 파일 열기 실패: {e}")
    
    def open_folder_location(self):
        """선택된 파일이 있는 폴더를 엽니다."""
        if not self.current_selected_file or not os.path.exists(self.current_selected_file):
            print(f"[오류] 폴더 열기 실패: 파일 경로가 없거나 존재하지 않습니다. {self.current_selected_file}")
            return
        
        try:
            import subprocess
            import sys
            
            # 절대 경로로 변환
            file_path = os.path.abspath(self.current_selected_file)
            folder_path = os.path.dirname(file_path)
            
            print(f"[폴더] 파일 경로: {file_path}")
            print(f"[경로] 폴더 경로: {folder_path}")
            
            if sys.platform == "win32":
                # Windows에서는 explorer의 /select 옵션을 사용하여 파일을 선택한 상태로 폴더 열기
                file_path_normalized = os.path.normpath(file_path)
                subprocess.run(['explorer', '/select,', file_path_normalized])
                print(f"[성공] Windows 폴더 열기 성공: {folder_path}")
            elif sys.platform == "darwin":
                # macOS에서는 open 명령 사용
                subprocess.call(["open", folder_path])
                print(f"[성공] macOS 폴더 열기 성공: {folder_path}")
            else:
                # Linux에서는 xdg-open 사용
                subprocess.call(["xdg-open", folder_path])
                print(f"[성공] Linux 폴더 열기 성공: {folder_path}")
            
        except Exception as e:
            print(f"[오류] 폴더 열기 실패: {e}")
            print(f"[오류] 파일 경로: {self.current_selected_file}")
            print(f"[오류] 폴더 경로: {os.path.dirname(self.current_selected_file)}")
    
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
        
        print(f"[로딩] 파일 뷰어에서 열기: {self.current_selected_file}")
        
        # 파일 선택 신호 발생
        self.file_selected.emit(self.current_selected_file)
    
    def close_loading_dialog(self):
        """로딩 알림창을 닫습니다."""
        if hasattr(self, 'loading_dialog') and self.loading_dialog:
            self.loading_dialog.close()
            self.loading_dialog = None
            print("[성공] 파일 로딩 완료 - 알림창 닫음")
        
        # 버튼 다시 활성화 (로딩 완료 후)
        if self.current_selected_file:
            self.open_viewer_button.setEnabled(True)
    
    def get_current_matching_pages(self):
        """
        현재 선택된 검색 결과의 매칭된 페이지 목록을 반환합니다.
        
        Returns:
            list: 매칭된 페이지 번호 목록
        """
        if self.current_selected_result:
            return self.current_selected_result.get('matching_pages', [])
        return []
    
    
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
            print(f"[오류] 파일명 검색 중 오류: {e}")
        
        # 관련성 점수로 정렬 (파일명 일치도)
        results.sort(key=lambda x: x['relevance_score'], reverse=True)
        
        return results