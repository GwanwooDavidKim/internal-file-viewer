# -*- coding: utf-8 -*-
"""
검색 위젯 (Search Widget)

파일 내용 검색을 위한 UI 위젯입니다.
"""
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLineEdit, 
                            QPushButton, QTreeWidget, QTreeWidgetItem, QLabel,
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
    
    progress_updated = pyqtSignal(str, float)
    indexing_finished = pyqtSignal(int)
    
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
    
    file_selected = pyqtSignal(str)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.indexer = SearchIndexer()
        self.indexing_worker = None
        self.current_directory = ""
        self.current_selected_file = None
        self.current_selected_result = None
        
        self.current_search_results = []
        self.current_sort_mode = "[정렬] 관련성 순 (기본)"
        
        self.setup_ui()
    
    def setup_ui(self):
        """UI 구성 요소를 설정합니다."""
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        search_frame = QFrame()
        search_layout = QVBoxLayout()
        search_frame.setLayout(search_layout)
        
        content_search_layout = QHBoxLayout()
        
        content_label = QLabel("📄 [파일] 내용:")
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
        
        exclude_search_layout = QHBoxLayout()
        
        exclude_label = QLabel("🚫 제외:")
        exclude_label.setMinimumWidth(60)
        exclude_search_layout.addWidget(exclude_label)
        
        self.exclude_search_input = QLineEdit()
        self.exclude_search_input.setPlaceholderText("제외할 키워드 (쉼표로 구분, 예: Fundamental,기초)")
        self.exclude_search_input.returnPressed.connect(self.perform_search)
        exclude_search_layout.addWidget(self.exclude_search_input)
        
        search_layout.addLayout(exclude_search_layout)
        
        help_label = QLabel("💡 팁: 내용에 키워드를 입력하고, 제외에 입력하면 해당 단어가 포함된 파일은 결과에서 빠집니다")
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
        
        self.indexed_extensions_label = QLabel("인덱싱 대상: .pdf .doc .docx .txt (※ Excel, PPT 제외)")
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
        
        sort_layout = QHBoxLayout()
        
        sort_label = QLabel("정렬 순서:")
        sort_layout.addWidget(sort_label)
        
        self.sort_combo = QComboBox()
        self.sort_combo.addItems([
            "[정렬] 관련성 순 (기본)",
            "📁 [폴더] 파일명 (오름차순)", 
            "📁 [폴더] 파일명 (내림차순)",
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
        
        self.progress_bar = QProgressBar()
        self.progress_bar.hide()
        search_layout.addWidget(self.progress_bar)
        
        self.progress_label = QLabel("")
        self.progress_label.hide()
        search_layout.addWidget(self.progress_label)
        
        layout.addWidget(search_frame)
        
        results_splitter = QSplitter(Qt.Orientation.Vertical)
        
        results_frame = QFrame()
        results_layout = QVBoxLayout()
        results_frame.setLayout(results_layout)
        
        self.results_label = QLabel("검색 결과")
        self.results_label.setFont(QFont(config.UI_FONTS["font_family"], 
                                       config.UI_FONTS["subtitle_size"], 
                                       QFont.Weight.Bold))
        results_layout.addWidget(self.results_label)
        
        self.results_list = QTreeWidget()
        self.results_list.setHeaderHidden(True)
        self.results_list.setIndentation(20)
        self.results_list.itemClicked.connect(self.on_result_selected)
        self.results_list.setMinimumHeight(200)
        results_layout.addWidget(self.results_list)
        
        results_splitter.addWidget(results_frame)
        
        actions_frame = QFrame()
        actions_layout = QHBoxLayout()
        actions_frame.setLayout(actions_layout)
        
        actions_layout.addStretch()
        
        self.open_viewer_button = QPushButton("📄 뷰어에서 열기")
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
        self.open_original_button.setEnabled(False)
        actions_layout.addWidget(self.open_original_button)
        
        results_splitter.addWidget(actions_frame)
        
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
        
        tree_style = f"""
            QTreeWidget {{
                background-color: white;
                border: 1px solid {config.UI_COLORS['secondary']};
                font-size: {config.UI_FONTS['body_size']}px;
            }}
            QTreeWidget::item {{
                padding: 6px 4px;
                border-bottom: 1px solid #EEEEEE;
            }}
            QTreeWidget::item:hover {{
                background-color: {config.UI_COLORS['hover']};
            }}
            QTreeWidget::item:selected {{
                background-color: {config.UI_COLORS['accent']};
                color: white;
            }}
            QTreeWidget::branch {{
                background-color: white;
            }}
        """
        self.results_list.setStyleSheet(tree_style)
    
    def set_directory(self, directory_path: str):
        """
        검색 대상 디렉토리를 설정합니다.
        
        Args:
            directory_path (str): 디렉토리 경로
        """
        self.current_directory = directory_path
        self.index_button.setText(f"📂 [경로] '{os.path.basename(directory_path)}' 인덱싱")
        self.index_button.setEnabled(True)
    
    def start_indexing(self):
        """인덱싱을 시작합니다."""
        if not self.current_directory or not os.path.exists(self.current_directory):
            self.results_label.setText("검색 결과 - 디렉토리를 먼저 선택해주세요")
            return
        
        if self.indexing_worker and self.indexing_worker.isRunning():
            return
        
        self.index_button.setEnabled(False)
        self.progress_bar.show()
        self.progress_bar.setValue(0)
        self.progress_label.show()
        self.progress_label.setText("인덱싱 준비 중...")
        
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
        if len(text.strip()) < 2:
            self.results_list.clear()
            self.results_label.setText("검색 결과")
    
    def perform_search(self):
        """검색을 수행합니다 (제외 키워드 지원)."""
        exclude_query = self.exclude_search_input.text().strip()
        content_query = self.search_input.text().strip()
        
        if not content_query:
            self.results_label.setText("검색 결과 - 내용 검색어를 입력해주세요")
            return
        
        display_text = f"내용:{content_query}"
        if exclude_query:
            display_text += f", 제외:{exclude_query}"
        
        self.results_label.setText(f"🔍 '{display_text}' 조회 중...")
        self.results_list.clear()
        
        QApplication.processEvents()
        
        if not self.indexer or len(self.indexer.indexed_paths) == 0:
            QMessageBox.warning(self, "인덱싱 필요", 
                               "파일 내용 검색을 위해서는 먼저 인덱싱을 완료해야 합니다.\n\n'[경로] 폴더 인덱싱' 버튼을 클릭하여 인덱싱을 시작하세요.")
            self.results_list.clear()
            self.results_label.setText("검색 결과")
            return
        
        search_results = self.indexer.search_files(content_query, exclude_query=exclude_query)
        
        self.current_search_results = search_results
        self._display_sorted_results(display_text)
    
    def on_result_selected(self, item: QTreeWidgetItem):
        """검색 결과 선택 시 호출됩니다."""
        if item.childCount() > 0:
            item.setExpanded(not item.isExpanded())
            self.open_viewer_button.setEnabled(False)
            self.open_original_button.setEnabled(False)
            self.open_folder_button.setEnabled(False)
            self.current_selected_file = None
            self.current_selected_result = None
            return
        
        result = item.data(0, Qt.ItemDataRole.UserRole)
        
        if result is None:
            self.open_viewer_button.setEnabled(False)
            self.open_original_button.setEnabled(False)
            self.open_folder_button.setEnabled(False)
            self.current_selected_file = None
            self.current_selected_result = None
            return
        
        if result:
            self.current_selected_file = result['file_path']
            self.current_selected_result = result
            
            self.open_viewer_button.setEnabled(True)
            self.open_original_button.setEnabled(True)
            self.open_folder_button.setEnabled(True)
    
    def on_sort_changed(self, sort_text: str):
        """정렬 방식 변경 시 호출됩니다."""
        self.current_sort_mode = sort_text
        if self.current_search_results:
            content_query = self.search_input.text().strip()
            display_text = f"내용:{content_query}"
            exclude_query = self.exclude_search_input.text().strip()
            if exclude_query:
                display_text += f", 제외:{exclude_query}"
            self._display_sorted_results(display_text)
    
    def _sort_results(self, results: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """검색 결과를 현재 정렬 모드에 따라 정렬합니다."""
        if not results:
            return results
        
        sort_mode = self.current_sort_mode
        
        if "관련성" in sort_mode:
            return results
        elif "파일명 (오름차순)" in sort_mode:
            return sorted(results, key=lambda x: x['filename'].lower())
        elif "파일명 (내림차순)" in sort_mode:
            return sorted(results, key=lambda x: x['filename'].lower(), reverse=True)
        elif "최신 변경일" in sort_mode:
            return sorted(results, key=lambda x: self._get_file_mtime(x['file_path']), reverse=True)
        elif "오래된 변경일" in sort_mode:
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
    
    def _group_by_directory(self, results: List[Dict[str, Any]]) -> Dict[str, List[Dict[str, Any]]]:
        """결과를 디렉토리별로 그룹화합니다."""
        import os
        groups = {}
        
        for result in results:
            file_path = result.get('file_path', '')
            directory = os.path.dirname(file_path)
            
            if not directory:
                directory = "(루트)"
            
            if directory not in groups:
                groups[directory] = []
            groups[directory].append(result)
        
        sorted_groups = dict(sorted(groups.items()))
        return sorted_groups
    
    def _display_sorted_results(self, query: str):
        """정렬된 검색 결과를 QTreeWidget에 표시합니다."""
        self.results_list.clear()
        
        if not self.current_search_results:
            self.results_label.setText(f"검색 결과 - '{query}'에 대한 결과 없음")
            return
        
        sorted_results = self._sort_results(self.current_search_results)
        
        dir_groups = self._group_by_directory(sorted_results)
        
        total_count = len(sorted_results)
        self.results_label.setText(f"검색 결과 - '{query}' ({total_count}개) | {self.current_sort_mode}")
        
        for directory, dir_results in dir_groups.items():
            if directory == "(루트)":
                display_path = "(루트)"
            elif self.current_directory:
                try:
                    rel_path = os.path.relpath(directory, self.current_directory)
                    display_path = rel_path if rel_path != "." else "(루트)"
                except ValueError:
                    display_path = directory
            else:
                display_path = directory
            
            dir_item = QTreeWidgetItem(self.results_list)
            dir_item.setText(0, f"📁 {display_path} ({len(dir_results)}개)")
            font = dir_item.font(0)
            font.setBold(True)
            dir_item.setFont(0, font)
            dir_item.setToolTip(0, f"전체 경로: {directory}")
            dir_item.setExpanded(False)
            
            for result in dir_results:
                filename = result['filename']
                file_type = result['file_type'].upper()
                file_size = result['file_size_mb']
                matching_pages = result.get('matching_pages', [])
                
                page_info = ""
                if matching_pages:
                    if len(matching_pages) <= 5:
                        page_info = f" | 페이지: {', '.join(map(str, matching_pages))}"
                    else:
                        page_info = f" | 페이지: {', '.join(map(str, matching_pages[:5]))}... ({len(matching_pages)}개)"
                
                file_item = QTreeWidgetItem(dir_item)
                file_item.setText(0, f"📄 {filename} ({file_type}, {file_size}MB){page_info}")
                file_item.setData(0, Qt.ItemDataRole.UserRole, result)
                
                tooltip = f"전체 경로: {result.get('file_path', '')}"
                if matching_pages:
                    tooltip += f"\n검색어 포함 페이지: {', '.join(map(str, matching_pages))}"
                file_item.setToolTip(0, tooltip)
    
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
                os.startfile(self.current_selected_file)
            elif sys.platform == "darwin":
                subprocess.call(["open", self.current_selected_file])
            else:
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
            
            file_path = os.path.abspath(self.current_selected_file)
            folder_path = os.path.dirname(file_path)
            
            print(f"[폴더] 파일 경로: {file_path}")
            print(f"[경로] 폴더 경로: {folder_path}")
            
            if sys.platform == "win32":
                file_path_normalized = os.path.normpath(file_path)
                subprocess.run(['explorer', '/select,', file_path_normalized])
                print(f"[성공] Windows 폴더 열기 성공: {folder_path}")
            elif sys.platform == "darwin":
                subprocess.call(["open", folder_path])
                print(f"[성공] macOS 폴더 열기 성공: {folder_path}")
            else:
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
        
        self.open_viewer_button.setEnabled(False)
        
        from PyQt6.QtWidgets import QProgressDialog
        from PyQt6.QtCore import Qt
        
        self.loading_dialog = QProgressDialog("파일 로딩중입니다...", None, 0, 0, self)
        self.loading_dialog.setWindowTitle("파일 로딩 중")
        self.loading_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.loading_dialog.setAutoClose(False)
        self.loading_dialog.setAutoReset(False)
        self.loading_dialog.show()
        
        print(f"[로딩] 파일 뷰어에서 열기: {self.current_selected_file}")
        
        self.file_selected.emit(self.current_selected_file)
    
    def close_loading_dialog(self):
        """로딩 알림창을 닫습니다."""
        if hasattr(self, 'loading_dialog') and self.loading_dialog:
            self.loading_dialog.close()
            self.loading_dialog = None
            print("[성공] 파일 로딩 완료 - 알림창 닫음")
        
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
