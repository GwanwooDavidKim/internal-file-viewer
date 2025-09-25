# -*- coding: utf-8 -*-
"""
메인 윈도우 (Main Window)

사내 파일 뷰어의 메인 사용자 인터페이스를 제공합니다.
"""
from PyQt6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QSplitter, QLabel, QMenuBar, QStatusBar, QPushButton,
                            QFileDialog, QMessageBox, QDialog)
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QFont, QAction
import config
from core.auth import AuthenticationManager
from ui.login_dialog import LoginDialog
from ui.file_browser import FileBrowser
from ui.content_viewer import ContentViewer
from ui.search_widget import SearchWidget


class MainWindow(QMainWindow):
    """
    사내 파일 뷰어의 메인 윈도우 클래스입니다.
    """
    
    def __init__(self):
        """MainWindow 인스턴스를 초기화합니다."""
        super().__init__()
        self.auth_manager = AuthenticationManager()
        self.current_folder_path = ""
        self.setup_ui()
        self.setup_session_timer()
        
        # 로그인 프로세스 시작
        if not self.show_login_dialog():
            self.close()
    
    def setup_ui(self):
        """UI 구성 요소를 설정합니다."""
        self.setWindowTitle(config.APP_SETTINGS["app_name"])
        self.setGeometry(100, 100, 
                        config.APP_SETTINGS["window_width"], 
                        config.APP_SETTINGS["window_height"])
        
        # 중앙 위젯
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 메인 레이아웃
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)
        
        # 상단 도구 모음
        toolbar_layout = QHBoxLayout()
        
        # 폴더 선택 버튼
        self.folder_button = QPushButton("📁 폴더 선택")
        self.folder_button.clicked.connect(self.select_folder)
        toolbar_layout.addWidget(self.folder_button)
        
        # 새로고침 버튼
        self.refresh_button = QPushButton("🔄 새로고침")
        self.refresh_button.clicked.connect(self.refresh_view)
        toolbar_layout.addWidget(self.refresh_button)
        
        toolbar_layout.addStretch()
        
        # 크레딧 표시 (사용자 요청: MCI팀 gwanwookim 크레딧)
        credit_label = QLabel("💎 Made by MCI Team • gwanwookim")
        credit_label.setStyleSheet(f"""
            QLabel {{
                color: {config.UI_COLORS['accent']};
                font-size: 11px;
                font-weight: bold;
                padding: 5px 10px;
                border-radius: 3px;
                background-color: rgba(63, 81, 181, 0.1);
            }}
        """)
        toolbar_layout.addWidget(credit_label)
        
        # 사용자 정보 레이블
        self.user_info_label = QLabel("")
        toolbar_layout.addWidget(self.user_info_label)
        
        # 로그아웃 버튼
        self.logout_button = QPushButton("로그아웃")
        self.logout_button.clicked.connect(self.logout)
        toolbar_layout.addWidget(self.logout_button)
        
        main_layout.addLayout(toolbar_layout)
        
        # 메인 스플리터 (2-Pane 구조)
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # 좌측 패널 (파일 탐색기)
        left_panel = QWidget()
        left_layout = QVBoxLayout()
        left_panel.setLayout(left_layout)
        
        explorer_title = QLabel("📂 파일 탐색기")
        explorer_title.setFont(QFont(config.UI_FONTS["font_family"], 
                                   config.UI_FONTS["subtitle_size"], 
                                   QFont.Weight.Bold))
        left_layout.addWidget(explorer_title)
        
        # 파일 브라우저 위젯
        self.file_browser = FileBrowser()
        self.file_browser.file_selected.connect(self.on_file_selected)
        left_layout.addWidget(self.file_browser)
        
        # 우측 패널 (콘텐츠 뷰어 및 검색)
        right_panel = QWidget()
        right_layout = QVBoxLayout()
        right_panel.setLayout(right_layout)
        
        # 탭 위젯으로 콘텐츠 뷰어와 검색 분리
        from PyQt6.QtWidgets import QTabWidget
        self.right_tabs = QTabWidget()
        
        # 콘텐츠 뷰어 탭
        content_widget = QWidget()
        content_layout = QVBoxLayout()
        content_widget.setLayout(content_layout)
        
        viewer_title = QLabel("👁️ 콘텐츠 뷰어")
        viewer_title.setFont(QFont(config.UI_FONTS["font_family"], 
                                 config.UI_FONTS["subtitle_size"], 
                                 QFont.Weight.Bold))
        content_layout.addWidget(viewer_title)
        
        self.content_viewer = ContentViewer()
        content_layout.addWidget(self.content_viewer)
        
        self.right_tabs.addTab(content_widget, "📄 파일 뷰어")
        
        # 검색 탭
        self.search_widget = SearchWidget()
        self.search_widget.file_selected.connect(self.on_file_selected)
        self.right_tabs.addTab(self.search_widget, "🔍 파일 검색")
        
        right_layout.addWidget(self.right_tabs)
        
        # 스플리터에 패널 추가
        self.splitter.addWidget(left_panel)
        self.splitter.addWidget(right_panel)
        
        # 스플리터 비율 설정
        splitter_ratio = config.APP_SETTINGS["splitter_ratio"]
        total_width = config.APP_SETTINGS["window_width"]
        left_width = int(total_width * splitter_ratio)
        right_width = total_width - left_width
        self.splitter.setSizes([left_width, right_width])
        
        main_layout.addWidget(self.splitter)
        
        # 상태 표시줄
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("준비 완료")
        
        # 스타일 적용
        self.apply_styles()
    
    def apply_styles(self):
        """애플리케이션 전체 스타일을 적용합니다."""
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
        
        self.folder_button.setStyleSheet(button_style)
        self.refresh_button.setStyleSheet(button_style)
        
        logout_style = f"""
            QPushButton {{
                background-color: {config.UI_COLORS['secondary']};
                color: {config.UI_COLORS['text']};
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-size: {config.UI_FONTS['body_size']}px;
            }}
            QPushButton:hover {{
                background-color: #A8B0B3;
            }}
        """
        self.logout_button.setStyleSheet(logout_style)
    
    def setup_session_timer(self):
        """세션 유효성 검사 타이머를 설정합니다."""
        self.session_timer = QTimer()
        self.session_timer.timeout.connect(self.check_session)
        self.session_timer.start(60000)  # 1분마다 검사
    
    def show_login_dialog(self):
        """
        로그인 다이얼로그를 표시합니다.
        
        Returns:
            bool: 로그인 성공 여부
        """
        login_dialog = LoginDialog(self)
        
        while True:
            if login_dialog.exec() == QDialog.DialogCode.Accepted:
                username, password = login_dialog.get_credentials()
                
                if not username or not password:
                    QMessageBox.warning(self, "입력 오류", "사용자명과 비밀번호를 모두 입력해주세요.")
                    login_dialog.clear_fields()
                    continue
                
                success, message = self.auth_manager.authenticate(username, password)
                
                if success:
                    QMessageBox.information(self, "로그인 성공", message)
                    self.update_user_info()
                    return True
                else:
                    QMessageBox.warning(self, "로그인 실패", message)
                    login_dialog.clear_fields()
            else:
                return False
    
    def update_user_info(self):
        """사용자 정보를 업데이트합니다."""
        user_info = self.auth_manager.get_user_info()
        if user_info:
            if user_info["is_admin"]:
                info_text = f"👤 {user_info['username']} (관리자)"
            else:
                remaining_days = user_info.get("remaining_days", 0)
                info_text = f"👤 {user_info['username']} (남은 일수: {remaining_days}일)"
            
            self.user_info_label.setText(info_text)
            self.user_info_label.setStyleSheet(f"color: {config.UI_COLORS['primary']}; font-weight: bold;")
    
    def check_session(self):
        """세션 유효성을 확인합니다."""
        valid, message = self.auth_manager.check_session_validity()
        if not valid:
            QMessageBox.warning(self, "세션 만료", message)
            self.close()
    
    def select_folder(self):
        """폴더 선택 다이얼로그를 표시합니다."""
        folder_path = QFileDialog.getExistingDirectory(
            self, "탐색할 폴더 선택", self.current_folder_path
        )
        
        if folder_path:
            self.current_folder_path = folder_path
            self.file_browser.set_root_path(folder_path)
            self.search_widget.set_directory(folder_path)
            self.status_bar.showMessage(f"폴더 로드됨: {folder_path}")
    
    def refresh_view(self):
        """뷰를 새로고침합니다."""
        if self.current_folder_path:
            self.status_bar.showMessage("새로고침 중...")
            self.file_browser.refresh_view()
            self.status_bar.showMessage("새로고침 완료")
        else:
            QMessageBox.information(self, "알림", "먼저 폴더를 선택해주세요.")
    
    def on_file_selected(self, file_path: str):
        """
        파일이 선택되었을 때 호출됩니다.
        
        Args:
            file_path (str): 선택된 파일의 경로
        """
        self.status_bar.showMessage(f"파일 로딩 중: {file_path}")
        self.content_viewer.load_file(file_path)
        self.status_bar.showMessage(f"파일 로드됨: {file_path}")
    
    def logout(self):
        """로그아웃을 수행합니다."""
        reply = QMessageBox.question(
            self, "로그아웃", "정말 로그아웃하시겠습니까?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            self.auth_manager.logout()
            self.close()
    
    def closeEvent(self, event):
        """창 닫기 이벤트를 처리합니다."""
        if self.auth_manager.is_logged_in():
            reply = QMessageBox.question(
                self, "종료 확인", "프로그램을 종료하시겠습니까?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            
            if reply == QMessageBox.StandardButton.Yes:
                self.auth_manager.logout()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()