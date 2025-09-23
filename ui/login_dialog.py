# -*- coding: utf-8 -*-
"""
로그인 다이얼로그 (Login Dialog)

사용자 인증을 위한 로그인 화면을 제공합니다.
"""
from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
                            QLineEdit, QPushButton, QMessageBox, QFrame)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
import config


class LoginDialog(QDialog):
    """
    사용자 로그인을 위한 다이얼로그 클래스입니다.
    """
    
    def __init__(self, parent=None):
        """
        LoginDialog 인스턴스를 초기화합니다.
        
        Args:
            parent: 부모 위젯
        """
        super().__init__(parent)
        self.username = None
        self.password = None
        self.setup_ui()
        
    def setup_ui(self):
        """UI 구성 요소를 설정합니다."""
        self.setWindowTitle(config.APP_SETTINGS["app_name"] + " - 로그인")
        self.setFixedSize(350, 200)
        self.setModal(True)
        
        # 메인 레이아웃
        layout = QVBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # 제목
        title_label = QLabel(config.APP_SETTINGS["app_name"])
        title_font = QFont()
        title_font.setPointSize(config.UI_FONTS["title_size"])
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet(f"color: {config.UI_COLORS['primary']};")
        layout.addWidget(title_label)
        
        # 구분선
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        layout.addWidget(line)
        
        # 사용자명 입력
        username_layout = QHBoxLayout()
        username_label = QLabel("사용자명:")
        username_label.setMinimumWidth(70)
        self.username_edit = QLineEdit()
        self.username_edit.setPlaceholderText("사용자명을 입력하세요")
        username_layout.addWidget(username_label)
        username_layout.addWidget(self.username_edit)
        layout.addLayout(username_layout)
        
        # 비밀번호 입력
        password_layout = QHBoxLayout()
        password_label = QLabel("비밀번호:")
        password_label.setMinimumWidth(70)
        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_edit.setPlaceholderText("비밀번호를 입력하세요")
        password_layout.addWidget(password_label)
        password_layout.addWidget(self.password_edit)
        layout.addLayout(password_layout)
        
        # 버튼
        button_layout = QHBoxLayout()
        self.login_button = QPushButton("로그인")
        self.cancel_button = QPushButton("취소")
        
        # 버튼 스타일
        button_style = f"""
            QPushButton {{
                background-color: {config.UI_COLORS['accent']};
                color: white;
                border: none;
                padding: 8px;
                border-radius: 4px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: {config.UI_COLORS['hover']};
            }}
            QPushButton:pressed {{
                background-color: {config.UI_COLORS['primary']};
            }}
        """
        self.login_button.setStyleSheet(button_style)
        
        cancel_button_style = f"""
            QPushButton {{
                background-color: {config.UI_COLORS['secondary']};
                color: {config.UI_COLORS['text']};
                border: none;
                padding: 8px;
                border-radius: 4px;
            }}
            QPushButton:hover {{
                background-color: #A8B0B3;
            }}
        """
        self.cancel_button.setStyleSheet(cancel_button_style)
        
        button_layout.addWidget(self.cancel_button)
        button_layout.addWidget(self.login_button)
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
        
        # 이벤트 연결
        self.login_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.reject)
        self.password_edit.returnPressed.connect(self.accept)
        self.username_edit.returnPressed.connect(self.password_edit.setFocus)
        
        # 초기 포커스
        self.username_edit.setFocus()
    
    def get_credentials(self):
        """
        입력된 로그인 정보를 반환합니다.
        
        Returns:
            tuple: (사용자명, 비밀번호)
        """
        return self.username_edit.text().strip(), self.password_edit.text()
    
    def clear_fields(self):
        """입력 필드를 모두 지웁니다."""
        self.username_edit.clear()
        self.password_edit.clear()
        self.username_edit.setFocus()