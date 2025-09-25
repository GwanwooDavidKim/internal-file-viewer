# -*- coding: utf-8 -*-
"""
파일 브라우저 위젯 (File Browser Widget)

QTreeView와 QFileSystemModel을 사용하여 파일 시스템을 탐색하는 위젯입니다.
"""
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QTreeView, 
                            QLineEdit, QPushButton, QLabel, QComboBox, QFrame)
from PyQt6.QtGui import QFileSystemModel
from PyQt6.QtCore import Qt, QDir, QFileSystemWatcher, pyqtSignal, QModelIndex, QSortFilterProxyModel
from PyQt6.QtGui import QFont
import os
from typing import Optional
import config
from utils.file_manager import FileManager


class FileFilterProxyModel(QSortFilterProxyModel):
    """
    파일 형식 필터링을 지원하는 프록시 모델입니다.
    """
    
    def __init__(self, file_manager: FileManager):
        super().__init__()
        self.file_manager = file_manager
        self.show_all_files = False  # 기본적으로 지원되는 파일만 표시
        
    def set_show_all_files(self, show_all: bool):
        """
        모든 파일 표시 여부를 설정합니다.
        
        Args:
            show_all (bool): True면 모든 파일, False면 지원되는 파일만 표시
        """
        self.show_all_files = show_all
        self.invalidateFilter()  # 필터 다시 적용
    
    def filterAcceptsRow(self, source_row: int, source_parent: QModelIndex) -> bool:
        """행이 필터를 통과하는지 확인합니다."""
        if self.show_all_files:
            return True  # 모든 파일 표시
        
        source_model = self.sourceModel()
        if not source_model:
            return True
            
        index = source_model.index(source_row, 0, source_parent)
        if not index.isValid():
            return True
            
        # QFileSystemModel로 캐스팅하여 filePath 메서드 사용
        if hasattr(source_model, 'filePath'):
            file_path = source_model.filePath(index)
        else:
            return True
        
        # 디렉토리는 항상 표시
        if os.path.isdir(file_path):
            return True
        
        # 파일인 경우 지원되는 파일만 표시
        if os.path.isfile(file_path):
            return self.file_manager.is_supported_file(file_path)
        
        return True


class FileBrowser(QWidget):
    """
    파일 브라우저 위젯 클래스입니다.
    
    QTreeView를 사용하여 폴더 구조를 표시하고,
    파일 선택 시 신호를 발생시킵니다.
    """
    
    # 파일 선택 시 발생하는 신호
    file_selected = pyqtSignal(str)  # 파일 경로
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_manager = FileManager()
        self.current_path = ""
        self.setup_ui()
        self.setup_file_watcher()
    
    def setup_ui(self):
        """UI 구성 요소를 설정합니다."""
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        # 상단 컨트롤 패널
        controls_frame = QFrame()
        controls_layout = QVBoxLayout()
        controls_frame.setLayout(controls_layout)
        
        # 현재 경로 표시
        self.path_label = QLabel("경로: (폴더를 선택해주세요)")
        self.path_label.setStyleSheet(f"""
            QLabel {{
                color: {config.UI_COLORS['text']};
                font-size: {config.UI_FONTS['body_size']}px;
                padding: 4px;
                background-color: {config.UI_COLORS['background']};
                border: 1px solid {config.UI_COLORS['secondary']};
                border-radius: 3px;
            }}
        """)
        controls_layout.addWidget(self.path_label)
        
        # 필터 컨트롤
        filter_layout = QHBoxLayout()
        
        filter_label = QLabel("표시:")
        filter_layout.addWidget(filter_label)
        
        self.filter_combo = QComboBox()
        self.filter_combo.addItems(["지원되는 파일만", "모든 파일"])
        self.filter_combo.setCurrentText("지원되는 파일만")  # 기본값을 "지원되는 파일만"으로 설정
        self.filter_combo.currentTextChanged.connect(self.on_filter_changed)
        filter_layout.addWidget(self.filter_combo)
        
        filter_layout.addStretch()
        
        # 새로고침 버튼
        self.refresh_btn = QPushButton("🔄")
        self.refresh_btn.setToolTip("새로고침")
        self.refresh_btn.clicked.connect(self.refresh_view)
        self.refresh_btn.setMaximumWidth(30)
        filter_layout.addWidget(self.refresh_btn)
        
        controls_layout.addLayout(filter_layout)
        layout.addWidget(controls_frame)
        
        # 파일 트리 뷰
        self.tree_view = QTreeView()
        self.tree_view.setAlternatingRowColors(True)
        self.tree_view.setRootIsDecorated(True)
        self.tree_view.setSortingEnabled(True)
        self.tree_view.sortByColumn(0, Qt.SortOrder.AscendingOrder)
        
        # 파일 시스템 모델 설정
        self.file_model = QFileSystemModel()
        self.file_model.setReadOnly(True)
        self.file_model.setFilter(QDir.Filter.AllDirs | QDir.Filter.Files | QDir.Filter.NoDotAndDotDot)
        
        # 프록시 모델로 필터링
        self.model = FileFilterProxyModel(self.file_manager)
        self.model.setSourceModel(self.file_model)
        
        self.tree_view.setModel(self.model)
        
        # 열 설정 (파일명, 크기, 타입, 수정일)
        self.tree_view.setColumnWidth(0, 200)  # 파일명
        self.tree_view.setColumnWidth(1, 80)   # 크기
        self.tree_view.setColumnWidth(2, 80)   # 타입
        self.tree_view.setColumnWidth(3, 120)  # 수정일
        
        # 파일 선택 시그널 연결
        self.tree_view.clicked.connect(self.on_file_clicked)
        self.tree_view.doubleClicked.connect(self.on_file_double_clicked)
        
        layout.addWidget(self.tree_view)
        
        # 하단 정보 패널
        self.info_label = QLabel("파일을 선택하세요")
        self.info_label.setStyleSheet(f"""
            QLabel {{
                color: {config.UI_COLORS['text']};
                font-size: {config.UI_FONTS['small_size']}px;
                padding: 4px;
                background-color: {config.UI_COLORS['background']};
                border: 1px solid {config.UI_COLORS['secondary']};
            }}
        """)
        layout.addWidget(self.info_label)
        
        self.apply_styles()
    
    def apply_styles(self):
        """스타일을 적용합니다."""
        tree_style = f"""
            QTreeView {{
                background-color: white;
                alternate-background-color: #F8F9FA;
                border: 1px solid {config.UI_COLORS['secondary']};
                selection-background-color: {config.UI_COLORS['accent']};
                selection-color: white;
                font-size: {config.UI_FONTS['body_size']}px;
            }}
            QTreeView::item {{
                padding: 4px;
                border: none;
            }}
            QTreeView::item:hover {{
                background-color: {config.UI_COLORS['hover']};
            }}
            QTreeView::item:selected {{
                background-color: {config.UI_COLORS['accent']};
                color: white;
            }}
            QHeaderView::section {{
                background-color: {config.UI_COLORS['secondary']};
                color: {config.UI_COLORS['text']};
                padding: 6px;
                border: 1px solid {config.UI_COLORS['primary']};
                font-weight: bold;
            }}
        """
        self.tree_view.setStyleSheet(tree_style)
        
        button_style = f"""
            QPushButton {{
                background-color: {config.UI_COLORS['accent']};
                color: white;
                border: none;
                border-radius: 3px;
                padding: 4px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: {config.UI_COLORS['hover']};
            }}
        """
        self.refresh_btn.setStyleSheet(button_style)
    
    def setup_file_watcher(self):
        """파일 시스템 변경 감지를 설정합니다."""
        self.file_watcher = QFileSystemWatcher()
        self.file_watcher.directoryChanged.connect(self.on_directory_changed)
    
    def set_root_path(self, path: str):
        """
        루트 경로를 설정합니다.
        
        Args:
            path (str): 루트 폴더 경로
        """
        if not os.path.exists(path):
            self.path_label.setText("경로: (잘못된 경로)")
            return
        
        self.current_path = path
        self.path_label.setText(f"경로: {path}")
        
        # 파일 시스템 모델에 루트 경로 설정
        self.file_model.setRootPath(path)
        
        # 프록시 모델을 통해 루트 인덱스 설정
        source_root_index = self.file_model.index(path)
        proxy_root_index = self.model.mapFromSource(source_root_index)
        self.tree_view.setRootIndex(proxy_root_index)
        
        # 파일 와처에 경로 추가
        if self.file_watcher.directories():
            self.file_watcher.removePaths(self.file_watcher.directories())
        self.file_watcher.addPath(path)
        
        self.info_label.setText(f"폴더 로드됨: {os.path.basename(path)}")
    
    def on_filter_changed(self, filter_text: str):
        """필터 변경 시 호출됩니다."""
        show_all = (filter_text == "모든 파일")
        self.model.set_show_all_files(show_all)
    
    def on_file_clicked(self, index: QModelIndex):
        """파일 클릭 시 호출됩니다."""
        # 프록시 모델에서 소스 모델로 인덱스 변환
        source_index = self.model.mapToSource(index)
        file_path = self.file_model.filePath(source_index)
        
        if os.path.isfile(file_path):
            # 파일 정보 표시
            file_info = self.file_manager.get_file_info(file_path)
            
            if file_info.get('supported', False):
                info_text = f"📄 {file_info['filename']} ({file_info['file_size_mb']} MB)"
                if file_info.get('file_type'):
                    info_text += f" - {file_info['file_type'].upper()}"
            else:
                info_text = f"📄 {os.path.basename(file_path)} (지원되지 않는 형식)"
            
            self.info_label.setText(info_text)
            
            # 파일 선택 신호 발생
            self.file_selected.emit(file_path)
        
        elif os.path.isdir(file_path):
            # 폴더 정보 표시
            try:
                item_count = len(os.listdir(file_path))
                self.info_label.setText(f"📁 {os.path.basename(file_path)} ({item_count}개 항목)")
            except PermissionError:
                self.info_label.setText(f"📁 {os.path.basename(file_path)} (접근 권한 없음)")
    
    def on_file_double_clicked(self, index: QModelIndex):
        """파일 더블클릭 시 호출됩니다."""
        # 프록시 모델에서 소스 모델로 인덱스 변환
        source_index = self.model.mapToSource(index)
        file_path = self.file_model.filePath(source_index)
        
        if os.path.isdir(file_path):
            # 폴더인 경우 해당 폴더로 이동
            self.set_root_path(file_path)
        elif os.path.isfile(file_path):
            # 파일인 경우 상세 정보 표시
            self.on_file_clicked(index)
    
    def on_directory_changed(self, path: str):
        """디렉토리 변경 시 호출됩니다."""
        self.info_label.setText(f"폴더 내용이 변경됨: {os.path.basename(path)}")
    
    def refresh_view(self):
        """뷰를 새로고침합니다."""
        if self.current_path:
            current_index = self.tree_view.currentIndex()
            
            # 모델 새로고침을 위해 루트 경로 재설정
            self.file_model.setRootPath("")
            self.file_model.setRootPath(self.current_path)
            
            # 프록시 모델을 통해 루트 인덱스 재설정
            source_root_index = self.file_model.index(self.current_path)
            proxy_root_index = self.model.mapFromSource(source_root_index)
            self.tree_view.setRootIndex(proxy_root_index)
            
            # 선택 상태 복원
            if current_index.isValid():
                self.tree_view.setCurrentIndex(current_index)
            
            self.info_label.setText("새로고침 완료")
    
    def get_current_path(self) -> str:
        """현재 경로를 반환합니다."""
        return self.current_path
    
    def get_selected_file(self) -> Optional[str]:
        """현재 선택된 파일의 경로를 반환합니다."""
        current_index = self.tree_view.currentIndex()
        if current_index.isValid():
            # 프록시 모델에서 소스 모델로 인덱스 변환
            source_index = self.model.mapToSource(current_index)
            file_path = self.file_model.filePath(source_index)
            if os.path.isfile(file_path):
                return file_path
        return None