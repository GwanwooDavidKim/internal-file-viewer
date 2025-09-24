# -*- coding: utf-8 -*-
"""
검색 위젯 (Search Widget)

파일 내용 검색을 위한 UI 위젯입니다.
"""
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLineEdit, 
                            QPushButton, QListWidget, QListWidgetItem, QLabel,
                            QProgressBar, QFrame, QSplitter, QTextEdit)
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
        self.setup_ui()
        
        # 검색 지연 타이머 (타이핑 완료 후 검색)
        self.search_timer = QTimer()
        self.search_timer.setSingleShot(True)
        self.search_timer.timeout.connect(self.perform_search)
    
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
        self.search_input.setPlaceholderText("파일 내용 검색... (2글자 이상 입력)")
        self.search_input.textChanged.connect(self.on_search_text_changed)
        self.search_input.returnPressed.connect(self.perform_search)
        search_input_layout.addWidget(self.search_input)
        
        self.search_button = QPushButton("🔍 검색")
        self.search_button.clicked.connect(self.perform_search)
        search_input_layout.addWidget(self.search_button)
        
        search_layout.addLayout(search_input_layout)
        
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
        
        # 미리보기 영역
        preview_frame = QFrame()
        preview_layout = QVBoxLayout()
        preview_frame.setLayout(preview_layout)
        
        preview_title = QLabel("미리보기")
        preview_title.setFont(QFont(config.UI_FONTS["font_family"], 
                                  config.UI_FONTS["subtitle_size"], 
                                  QFont.Weight.Bold))
        preview_layout.addWidget(preview_title)
        
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setMaximumHeight(150)
        self.preview_text.setPlainText("검색 결과를 선택하면 미리보기가 표시됩니다.")
        preview_layout.addWidget(self.preview_text)
        
        results_splitter.addWidget(preview_frame)
        
        # 스플리터 비율 설정
        results_splitter.setSizes([300, 150])
        
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
        
        text_style = f"""
            QTextEdit {{
                background-color: white;
                border: 1px solid {config.UI_COLORS['secondary']};
                font-family: 'Consolas', 'Monaco', monospace;
                font-size: {config.UI_FONTS['body_size']}px;
                line-height: 1.4;
            }}
        """
        self.preview_text.setStyleSheet(text_style)
    
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
        self.preview_text.setPlainText("인덱스가 초기화되었습니다.")
        self.update_index_stats()
        self.results_label.setText("검색 결과 - 인덱스 초기화됨")
    
    def update_index_stats(self):
        """인덱스 통계를 업데이트합니다."""
        stats = self.indexer.get_index_statistics()
        self.index_stats_label.setText(f"인덱스: {stats['total_files']}개 파일, {stats['total_tokens']}개 토큰")
    
    def on_search_text_changed(self, text: str):
        """검색 텍스트 변경 시 호출됩니다."""
        # 타이핑 중이면 타이머 리셋
        self.search_timer.stop()
        
        if len(text.strip()) >= 2:
            # 500ms 후 자동 검색
            self.search_timer.start(500)
        else:
            self.results_list.clear()
            self.preview_text.setPlainText("2글자 이상 입력해주세요.")
            self.results_label.setText("검색 결과")
    
    def perform_search(self):
        """검색을 수행합니다."""
        query = self.search_input.text().strip()
        
        if len(query) < 2:
            self.results_label.setText("검색 결과 - 2글자 이상 입력해주세요")
            return
        
        # 검색 실행
        search_results = self.indexer.search_files(query, max_results=100)
        
        # 결과 표시
        self.results_list.clear()
        self.preview_text.setPlainText("")
        
        if not search_results:
            self.results_label.setText(f"검색 결과 - '{query}'에 대한 결과 없음")
            self.preview_text.setPlainText("검색 결과가 없습니다. 다른 검색어를 시도해보세요.")
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
            # 미리보기 표시
            preview_text = f"파일: {result['filename']}\\n"
            preview_text += f"경로: {result['file_path']}\\n"
            preview_text += f"타입: {result['file_type']}\\n"
            preview_text += f"크기: {result['file_size_mb']}MB\\n"
            preview_text += f"관련성: {result['relevance_score']:.2f}\\n\\n"
            preview_text += "내용 미리보기:\\n"
            preview_text += result.get('preview', '미리보기 없음')
            
            self.preview_text.setPlainText(preview_text)
            
            # 파일 선택 신호 발생
            self.file_selected.emit(result['file_path'])
    
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