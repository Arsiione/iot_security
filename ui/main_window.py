from PyQt6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QLabel, QPushButton, QStackedWidget, QFrame, 
                            QSizePolicy, QGraphicsDropShadowEffect)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QFont, QIcon, QColor, QPalette
from .pages import setup_home_page, setup_scan_page, setup_dashboard_page, setup_results_page, setup_settings_page, setup_history_page

class MainWindow(QMainWindow):
    theme_changed = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("IoT Howpsuzlyk Skaneri")
        self.setGeometry(100, 100, 1200, 800)
        self.setMinimumSize(1000, 700)
        
        # Установка тёмной темы по умолчанию
        self.current_theme = "dark"
        self.apply_theme(self.current_theme)
        
        self.setup_ui()
        
    def apply_theme(self, theme):
        self.current_theme = theme
        palette = QPalette()
        
        if theme == "dark":
            # Тёмная тема
            palette.setColor(QPalette.ColorRole.Window, QColor(30, 30, 40))
            palette.setColor(QPalette.ColorRole.WindowText, QColor(220, 220, 220))
            palette.setColor(QPalette.ColorRole.Base, QColor(45, 45, 55))
            palette.setColor(QPalette.ColorRole.AlternateBase, QColor(50, 50, 60))
            palette.setColor(QPalette.ColorRole.Text, QColor(220, 220, 220))
            palette.setColor(QPalette.ColorRole.Button, QColor(60, 60, 70))
            palette.setColor(QPalette.ColorRole.ButtonText, QColor(220, 220, 220))
            palette.setColor(QPalette.ColorRole.Highlight, QColor(0, 150, 200))
        else:
            # Светлая тема
            palette.setColor(QPalette.ColorRole.Window, QColor(240, 240, 245))
            palette.setColor(QPalette.ColorRole.WindowText, QColor(50, 50, 50))
            palette.setColor(QPalette.ColorRole.Base, QColor(255, 255, 255))
            palette.setColor(QPalette.ColorRole.AlternateBase, QColor(245, 245, 245))
            palette.setColor(QPalette.ColorRole.Text, QColor(50, 50, 50))
            palette.setColor(QPalette.ColorRole.Button, QColor(230, 230, 235))
            palette.setColor(QPalette.ColorRole.ButtonText, QColor(50, 50, 50))
            palette.setColor(QPalette.ColorRole.Highlight, QColor(0, 120, 215))
        
        self.setPalette(palette)
        self.theme_changed.emit(theme)

    def setup_ui(self):
        # Главный контейнер
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Основной layout
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Боковая панель
        self.sidebar = self.create_sidebar()
        main_layout.addWidget(self.sidebar, 1)
        
        # Основная область контента
        content_area = QFrame()
        content_area.setObjectName("contentArea")
        content_layout = QVBoxLayout(content_area)
        content_layout.setContentsMargins(20, 20, 20, 20)
        content_layout.setSpacing(20)
        
        # Заголовок
        self.header_label = QLabel("Baş sahypa")
        self.header_label.setObjectName("headerLabel")
        self.header_label.setFont(QFont("Segoe UI", 24, QFont.Weight.Bold))
        content_layout.addWidget(self.header_label)
        
        # Область страниц
        self.stacked_widget = QStackedWidget()
        self.stacked_widget.setObjectName("stackedWidget")
        content_layout.addWidget(self.stacked_widget)
        
        main_layout.addWidget(content_area, 4)
        
        # Настройка страниц
        self.setup_pages()
        
        # Применяем стили
        self.apply_styles()

    def create_sidebar(self):
        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(250)
        
        layout = QVBoxLayout(sidebar)
        layout.setContentsMargins(0, 20, 0, 20)
        layout.setSpacing(10)
        
        # Логотип и заголовок
        logo_frame = QFrame()
        logo_layout = QVBoxLayout(logo_frame)
        logo_layout.setContentsMargins(20, 10, 20, 30)
        
        logo_text = QLabel("IoT Security")
        logo_text.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        logo_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        logo_text.setStyleSheet("color: #00c8ff;")
        logo_layout.addWidget(logo_text)
        
        version = QLabel("v2.0")
        version.setFont(QFont("Segoe UI", 10))
        version.setAlignment(Qt.AlignmentFlag.AlignCenter)
        version.setStyleSheet("color: #888;")
        logo_layout.addWidget(version)
        
        layout.addWidget(logo_frame)
        
        # Кнопки навигации
        nav_buttons = [
            ("🏠 Baş sahypa", "home"),
            ("🔍 Skanirleme", "scan"),
            ("📊 Panel", "dashboard"),
            ("📋 Netijeler", "results"),
            ("📜 Taryh", "history"),
            ("⚙️ Sazlamalar", "settings")
        ]
        
        self.nav_buttons = {}
        
        for text, page_name in nav_buttons:
            btn = QPushButton(text)
            btn.setObjectName(f"nav_{page_name}")
            btn.setFont(QFont("Segoe UI", 12))
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setCheckable(True)
            btn.clicked.connect(lambda checked, p=page_name: self.switch_page(p))
            layout.addWidget(btn)
            self.nav_buttons[page_name] = btn
        
        # Spacer
        layout.addStretch()
        
        # Кнопка смены темы
        theme_btn = QPushButton("🌙 Temany çalyş")
        theme_btn.setFont(QFont("Segoe UI", 11))
        theme_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        theme_btn.clicked.connect(self.toggle_theme)
        layout.addWidget(theme_btn)
        
        # Статус
        status_label = QLabel("🟢 Ulgam işjeň")
        status_label.setFont(QFont("Segoe UI", 10))
        status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        status_label.setStyleSheet("color: #4CAF50;")
        layout.addWidget(status_label)
        
        return sidebar

    def setup_pages(self):
        # Создаем страницы
        self.pages = {
            "home": setup_home_page(),
            "scan": setup_scan_page(),
            "dashboard": setup_dashboard_page(),
            "results": setup_results_page(),
            "history": setup_history_page(),
            "settings": setup_settings_page()
        }
        
        # Добавляем страницы в stacked widget
        for page_name, page_widget in self.pages.items():
            self.stacked_widget.addWidget(page_widget)
        
        # Устанавливаем первую страницу активной
        self.switch_page("home")

    def switch_page(self, page_name):
        # Обновляем активную кнопку
        for btn_name, btn in self.nav_buttons.items():
            btn.setChecked(btn_name == page_name)
        
        # Переключаем страницу
        if page_name in self.pages:
            page_widget = self.pages[page_name]
            if hasattr(page_widget, "refresh") and callable(page_widget.refresh):
                page_widget.refresh()
            self.stacked_widget.setCurrentWidget(page_widget)
            
            # Обновляем заголовок
            titles = {
                "home": "Baş sahypa",
                "scan": "Tor skanirlemesi",
                "dashboard": "Howpsuzlyk paneli",
                "results": "Skanirleme netijeleri",
                "history": "Barlaglaryň taryhy",
                "settings": "Ulgam sazlamalary"
            }
            self.header_label.setText(titles.get(page_name, "Baş sahypa"))

    def toggle_theme(self):
        new_theme = "light" if self.current_theme == "dark" else "dark"
        self.apply_theme(new_theme)
        self.apply_styles()

    def apply_styles(self):
        style = f"""
        #sidebar {{
            background-color: {'#2a2a3a' if self.current_theme == 'dark' else '#f0f0f5'};
            border-right: 1px solid {'#444' if self.current_theme == 'dark' else '#ddd'};
        }}
        
        #contentArea {{
            background-color: {'#1e1e2a' if self.current_theme == 'dark' else '#ffffff'};
        }}
        
        #headerLabel {{
            color: {'#ffffff' if self.current_theme == 'dark' else '#333333'};
            padding: 10px 0;
            border-bottom: 2px solid {'#00c8ff' if self.current_theme == 'dark' else '#0078d7'};
        }}
        
        QPushButton#nav_home, 
        QPushButton#nav_scan, 
        QPushButton#nav_dashboard, 
        QPushButton#nav_results, 
        QPushButton#nav_history, 
        QPushButton#nav_settings {{
            text-align: left;
            padding: 15px 20px;
            border: none;
            border-radius: 8px;
            margin: 0 10px;
            background-color: transparent;
            color: {'#cccccc' if self.current_theme == 'dark' else '#666666'};
        }}
        
        QPushButton#nav_home:hover, 
        QPushButton#nav_scan:hover, 
        QPushButton#nav_dashboard:hover, 
        QPushButton#nav_results:hover, 
        QPushButton#nav_history:hover, 
        QPushButton#nav_settings:hover {{
            background-color: {'#3a3a4a' if self.current_theme == 'dark' else '#e8e8e8'};
        }}
        
        QPushButton#nav_home:checked, 
        QPushButton#nav_scan:checked, 
        QPushButton#nav_dashboard:checked, 
        QPushButton#nav_results:checked, 
        QPushButton#nav_history:checked, 
        QPushButton#nav_settings:checked {{
            background-color: {'#00c8ff' if self.current_theme == 'dark' else '#0078d7'};
            color: white;
            font-weight: bold;
        }}
        
        QPushButton {{
            background-color: {'#3a3a4a' if self.current_theme == 'dark' else '#f8f8f8'};
            color: {'#ffffff' if self.current_theme == 'dark' else '#333333'};
            border: 1px solid {'#555' if self.current_theme == 'dark' else '#ccc'};
            border-radius: 6px;
            padding: 10px 15px;
            font-size: 12px;
        }}
        
        QPushButton:hover {{
            background-color: {'#4a4a5a' if self.current_theme == 'dark' else '#e8e8e8'};
        }}
        
        QPushButton:pressed {{
            background-color: {'#5a5a6a' if self.current_theme == 'dark' else '#d8d8d8'};
        }}
        
        QPushButton:disabled {{
            background-color: {'#2a2a3a' if self.current_theme == 'dark' else '#f0f0f0'};
            color: {'#777' if self.current_theme == 'dark' else '#999'};
        }}
        
        QLineEdit, QTextEdit, QComboBox {{
            background-color: {'#2a2a3a' if self.current_theme == 'dark' else '#ffffff'};
            color: {'#ffffff' if self.current_theme == 'dark' else '#333333'};
            border: 1px solid {'#5d5d70' if self.current_theme == 'dark' else '#c8c8d0'};
            border-radius: 6px;
            padding: 5px 8px;
            font-size: 12px;
            min-height: 20px;
        }}

        QLineEdit:focus, QComboBox:focus {{
            border: 1px solid {'#00c8ff' if self.current_theme == 'dark' else '#0078d7'};
        }}

        QComboBox {{
            padding-right: 24px;
        }}

        QComboBox::drop-down {{
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 24px;
            border-left: 1px solid {'#555' if self.current_theme == 'dark' else '#d0d0d0'};
            border-top-right-radius: 6px;
            border-bottom-right-radius: 6px;
        }}

        QComboBox::down-arrow {{
            width: 0;
            height: 0;
        }}

        QComboBox QAbstractItemView {{
            background-color: {'#2a2a3a' if self.current_theme == 'dark' else '#ffffff'};
            color: {'#ffffff' if self.current_theme == 'dark' else '#333333'};
            selection-background-color: {'#00c8ff' if self.current_theme == 'dark' else '#0078d7'};
            selection-color: white;
            outline: none;
        }}

        QTableWidget {{
            background-color: {'#2a2a2a' if self.current_theme == 'dark' else '#ffffff'};
            color: {'#ffffff' if self.current_theme == 'dark' else '#222222'};
            gridline-color: {'#3f3f46' if self.current_theme == 'dark' else '#dddddd'};
            selection-background-color: {'#00a6d6' if self.current_theme == 'dark' else '#cfe8ff'};
            selection-color: {'#ffffff' if self.current_theme == 'dark' else '#111111'};
        }}

        QHeaderView::section {{
            background-color: {'#3a3a3a' if self.current_theme == 'dark' else '#eeeeee'};
            color: {'#ffffff' if self.current_theme == 'dark' else '#222222'};
            padding: 6px;
            border: 1px solid {'#4a4a4a' if self.current_theme == 'dark' else '#d0d0d0'};
        }}
        
        QProgressBar {{
            border: 1px solid {'#555' if self.current_theme == 'dark' else '#ccc'};
            border-radius: 6px;
            text-align: center;
            background-color: {'#2a2a3a' if self.current_theme == 'dark' else '#f8f8f8'};
        }}
        
        QProgressBar::chunk {{
            background-color: {'#00c8ff' if self.current_theme == 'dark' else '#0078d7'};
            border-radius: 5px;
        }}
        
        QLabel {{
            color: {'#ffffff' if self.current_theme == 'dark' else '#333333'};
        }}
        
        .card {{
            background-color: {'#2a2a3a' if self.current_theme == 'dark' else '#ffffff'};
            border: 1px solid {'#444' if self.current_theme == 'dark' else '#eee'};
            border-radius: 12px;
            padding: 20px;
        }}
        
        .critical {{
            color: #ff4757;
            font-weight: bold;
        }}
        
        .high {{
            color: #ff6b81;
            font-weight: bold;
        }}
        
        .medium {{
            color: #ffa502;
            font-weight: bold;
        }}
        
        .low {{
            color: #2ed573;
            font-weight: bold;
        }}
        """
        
        self.setStyleSheet(style)
