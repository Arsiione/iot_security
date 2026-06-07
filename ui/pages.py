# ui/pages.py
from collections import Counter

from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                            QPushButton, QLineEdit, QProgressBar, QTextEdit,
                            QTableWidget, QTableWidgetItem, QHeaderView,
                            QFrame, QComboBox, QCheckBox, QGroupBox, QGridLayout,
                            QScrollArea, QFileDialog, QMessageBox,
                            QAbstractItemView, QSizePolicy)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
import ipaddress

# ЗАМЕНИТЕ относительные импорты на абсолютные:
from core.scanner import ScanThread, get_default_ip_range, get_network_adapters
from database import save_scan, load_history, get_scan_stats, clear_history
from reports import generate_pdf
from notifications import send_email


DEVICE_COLUMNS = [
    "IP", "Ady", "MAC", "Öndüriji", "Görnüşi", "Tapylan usul",
    "Portlar", "Töwekgelçilik", "Gowşaklyk", "Maslahat"
]

DEVICE_COLUMN_WIDTHS = [105, 160, 135, 110, 130, 180, 120, 120, 95, 320]


def localize_unknown(value):
    return "Näbelli" if not value or value == "Unknown" else value


def risk_label(value):
    labels = {
        "low": "pes",
        "medium": "orta",
        "high": "ýokary",
        "critical": "kritiki",
    }
    return labels.get(value or "low", value or "pes")


def format_ports(ports):
    return ", ".join(map(str, ports or [])) or "ýok"


def fill_devices_table(table, devices):
    table.setRowCount(0)
    for device in devices:
        row = table.rowCount()
        table.insertRow(row)
        values = [
            localize_unknown(device.get("ip")),
            localize_unknown(device.get("name") or device.get("hostname")),
            localize_unknown(device.get("mac")),
            localize_unknown(device.get("vendor")),
            localize_unknown(device.get("device_type")),
            localize_unknown(device.get("discovery_method")),
            format_ports(device.get("ports")),
            risk_label(device.get("risk_level", "low")),
            "Hawa" if device.get("vulnerable") else "Ýok",
            device.get("recommendation") or "Kritiki maslahat ýok.",
        ]
        for column, value in enumerate(values):
            item = QTableWidgetItem(str(value))
            item.setToolTip(str(value))
            if column in (7, 8):
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            table.setItem(row, column, item)


def configure_devices_table(table, fit_to_width=False):
    header = table.horizontalHeader()
    header.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)
    if fit_to_width:
        header.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        header.setStretchLastSection(False)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
    else:
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(False)
        for column, width in enumerate(DEVICE_COLUMN_WIDTHS):
            table.setColumnWidth(column, width)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

    table.setWordWrap(False)
    table.setTextElideMode(Qt.TextElideMode.ElideRight)
    table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
    table.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
    table.verticalHeader().setDefaultSectionSize(32)
    table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)


def create_devices_table():
    table = QTableWidget(0, len(DEVICE_COLUMNS))
    table.setHorizontalHeaderLabels(DEVICE_COLUMNS)
    configure_devices_table(table)
    table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
    table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
    table.setAlternatingRowColors(True)
    return table


def configure_combo(combo, minimum_contents=24):
    combo.setMinimumHeight(30)
    combo.setMinimumContentsLength(minimum_contents)
    combo.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
    combo.view().setTextElideMode(Qt.TextElideMode.ElideRight)


def highest_risk(devices):
    order = {"low": 0, "medium": 1, "high": 2, "critical": 3}
    highest = "low"
    for device in devices:
        level = device.get("risk_level", "low")
        if order.get(level, 0) > order.get(highest, 0):
            highest = level
    return highest


def count_vulnerable(devices):
    return sum(1 for device in devices if device.get("vulnerable"))


def count_with_ports(devices):
    return sum(1 for device in devices if device.get("ports"))


def first_scan_value(scan, key, default="Näbelli"):
    for device in scan.get("devices", []):
        value = device.get(key)
        if value:
            return value
    return default


def fill_summary_table(table, rows):
    table.setRowCount(0)
    for values in rows:
        row = table.rowCount()
        table.insertRow(row)
        for column, value in enumerate(values):
            item = QTableWidgetItem(str(value))
            if column:
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item.setToolTip(str(value))
            table.setItem(row, column, item)


def configure_summary_table(table, headers):
    table.setColumnCount(len(headers))
    table.setHorizontalHeaderLabels(headers)
    table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
    table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
    table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
    table.setAlternatingRowColors(True)


def scan_summary_text(scan):
    devices = scan.get("devices", [])
    if not devices:
        return "Bu skanirlemede gurluş tapylmady."

    lines = [
        f"Wagty: {scan.get('timestamp', 'Näbelli')}",
        f"Tor: {first_scan_value(scan, 'network')}",
        f"Adapter: {first_scan_value(scan, 'adapter')}",
        f"Gurluşlar: {len(devices)} | Gowşak: {count_vulnerable(devices)} | Portly: {count_with_ports(devices)}",
        "",
        "Esasy tapyndylar:",
    ]
    for device in devices[:6]:
        lines.append(
            f"- {device.get('ip', 'Näbelli')} | {device.get('device_type', 'Unknown')} | "
            f"{format_ports(device.get('ports'))} | {risk_label(device.get('risk_level', 'low'))}"
        )
    if len(devices) > 6:
        lines.append(f"... ýene {len(devices) - 6} gurluş")
    return "\n".join(lines)


def device_details_text(device):
    if not device:
        return "Gurluş saýlanmady."
    return "\n".join([
        f"IP: {localize_unknown(device.get('ip'))}",
        f"Ady: {localize_unknown(device.get('name') or device.get('hostname'))}",
        f"MAC: {localize_unknown(device.get('mac'))}",
        f"Öndüriji: {localize_unknown(device.get('vendor'))}",
        f"Görnüşi: {localize_unknown(device.get('device_type'))}",
        f"Tapylan usul: {localize_unknown(device.get('discovery_method'))}",
        f"Portlar: {format_ports(device.get('ports'))}",
        f"Töwekgelçilik: {risk_label(device.get('risk_level', 'low'))}",
        f"Gowşaklyk: {'Hawa' if device.get('vulnerable') else 'Ýok'}",
        f"Maslahat: {device.get('recommendation') or 'Kritiki maslahat ýok.'}",
    ])


def detect_local_ip_range():
    return get_default_ip_range()

def setup_home_page():
    widget = QWidget()
    layout = QVBoxLayout(widget)
    layout.setAlignment(Qt.AlignmentFlag.AlignTop)
    
    # Заголовок
    title = QLabel("IoT Howpsuzlyk Skanerine hoş geldiňiz")
    title.setFont(QFont("Segoe UI", 20, QFont.Weight.Bold))
    title.setAlignment(Qt.AlignmentFlag.AlignCenter)
    layout.addWidget(title)
    
    # Статистика
    stats_frame = QFrame()
    stats_frame.setObjectName("card")
    stats_layout = QHBoxLayout(stats_frame)
    
    stats = [
        ("📊", "Ähli gurluşlar", "0"),
        ("⚠️", "Gowşak gurluşlar", "0"),
        ("🛡️", "Goralan gurluşlar", "0"),
        ("🔍", "Soňky skanirleme", "Hiç haçan")
    ]
    
    for icon, text, value in stats:
        stat_widget = QWidget()
        stat_layout = QVBoxLayout(stat_widget)
        
        icon_label = QLabel(icon)
        icon_label.setFont(QFont("Segoe UI", 24))
        icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        text_label = QLabel(text)
        text_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        text_label.setFont(QFont("Segoe UI", 10))
        
        value_label = QLabel(value)
        value_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        value_label.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        
        stat_layout.addWidget(icon_label)
        stat_layout.addWidget(text_label)
        stat_layout.addWidget(value_label)
        stats_layout.addWidget(stat_widget)
    
    layout.addWidget(stats_frame)
    
    # Быстрые действия
    actions_frame = QFrame()
    actions_frame.setObjectName("card")
    actions_layout = QHBoxLayout(actions_frame)
    
    quick_actions = [
        ("🚀 Çalt skanirleme", "Tory çalt skanirlemegi başlat"),
        ("📋 Hasabat", "Soňky skanirleme boýunça hasabat döret"),
        ("⚡ Gowşaklyklary barla", "Belli gowşaklyklary barlamagy başlat")
    ]
    
    for btn_text, tooltip in quick_actions:
        btn = QPushButton(btn_text)
        btn.setFont(QFont("Segoe UI", 11))
        btn.setToolTip(tooltip)
        btn.setMinimumHeight(60)
        actions_layout.addWidget(btn)
    
    layout.addWidget(actions_frame)
    
    # Последние уведомления
    notifications_group = QGroupBox("Soňky duýduryşlar")
    notifications_group.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
    notifications_layout = QVBoxLayout(notifications_group)
    
    notifications = [
        ("🟢", "Ulgam işe girizildi we taýýar", "2 min öň"),
        ("🔵", "Gowşaklyk maglumat bazasy täzelendi", "5 min öň"),
        ("🟠", "Skanirleme düzgünlerini täzelemek maslahat berilýär", "10 min öň")
    ]
    
    for icon, text, time in notifications:
        notif_frame = QFrame()
        notif_layout = QHBoxLayout(notif_frame)
        
        icon_label = QLabel(icon)
        text_label = QLabel(text)
        time_label = QLabel(time)
        time_label.setStyleSheet("color: #888;")
        
        notif_layout.addWidget(icon_label)
        notif_layout.addWidget(text_label)
        notif_layout.addStretch()
        notif_layout.addWidget(time_label)
        
        notifications_layout.addWidget(notif_frame)
    
    layout.addWidget(notifications_group)
    layout.addStretch()
    
    return widget

def setup_scan_page():
    widget = QWidget()
    widget.scan_thread = None
    widget.last_devices = []
    layout = QVBoxLayout(widget)
    layout.setSpacing(8)
    layout.setContentsMargins(0, 0, 0, 0)
    
    # Настройки сканирования
    settings_group = QGroupBox("Skanirleme sazlamalary")
    settings_group.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
    settings_layout = QGridLayout(settings_group)
    settings_layout.setContentsMargins(10, 14, 10, 10)
    settings_layout.setColumnStretch(0, 0)
    settings_layout.setColumnStretch(1, 1)
    settings_layout.setColumnStretch(2, 0)
    settings_layout.setColumnStretch(3, 1)
    settings_layout.setHorizontalSpacing(12)
    settings_layout.setVerticalSpacing(6)

    adapters = get_network_adapters()
    settings_layout.addWidget(QLabel("Tor adapteri:"), 0, 0)
    adapter_combo = QComboBox()
    configure_combo(adapter_combo, 34)
    if adapters:
        for adapter in adapters:
            adapter_combo.addItem(adapter["label"], adapter)
    else:
        adapter_combo.addItem("Adapter tapylmady", None)
    settings_layout.addWidget(adapter_combo, 0, 1, 1, 3)
    
    settings_layout.addWidget(QLabel("IP aralygy:"), 1, 0)
    ip_range_input = QLineEdit(detect_local_ip_range())
    ip_range_input.setMinimumHeight(30)
    if adapters:
        ip_range_input.setText(adapters[0]["network"])
    settings_layout.addWidget(ip_range_input, 1, 1)
    
    settings_layout.addWidget(QLabel("Güýç derejesi:"), 1, 2)
    intensity_combo = QComboBox()
    configure_combo(intensity_combo, 12)
    intensity_combo.addItems(["Pes", "Orta", "Ýokary", "Agresiw"])
    intensity_combo.setCurrentIndex(1)
    settings_layout.addWidget(intensity_combo, 1, 3)
    
    settings_layout.addWidget(QLabel("Skanirleme görnüşi:"), 2, 0)
    scan_type_combo = QComboBox()
    configure_combo(scan_type_combo, 18)
    scan_type_combo.addItems(["Çalt", "Doly", "Diňe gowşaklyklar", "Ýörite"])
    settings_layout.addWidget(scan_type_combo, 2, 1)
    
    layout.addWidget(settings_group)
    
    # Дополнительные опции
    options_group = QGroupBox("Goşmaça mümkinçilikler")
    options_layout = QHBoxLayout(options_group)
    options_layout.setContentsMargins(10, 14, 10, 8)
    options_layout.setSpacing(18)
    
    options = [
        ("Kritiki gowşaklyklary awtomatik düzetmek", False),
        ("Jikme-jik loglary saklamak", True),
        ("Email arkaly duýduryş ibermek", False),
        ("Ätiýaçlyk nusgalary döretmek", True)
    ]
    
    for text, default in options:
        checkbox = QCheckBox(text)
        checkbox.setChecked(default)
        options_layout.addWidget(checkbox)
    options_layout.addStretch()
    
    layout.addWidget(options_group)
    
    # Прогресс и кнопки
    progress_bar = QProgressBar()
    progress_bar.setVisible(False)
    layout.addWidget(progress_bar)

    results_table = QTableWidget(0, len(DEVICE_COLUMNS))
    results_table.setHorizontalHeaderLabels(DEVICE_COLUMNS)
    configure_devices_table(results_table, fit_to_width=True)
    results_table.setMinimumHeight(285)
    results_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
    results_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
    results_table.setAlternatingRowColors(True)
    layout.addWidget(results_table, 1)
    
    log_output = QTextEdit()
    log_output.setMaximumHeight(58)
    log_output.setMinimumHeight(46)
    log_output.setReadOnly(True)
    layout.addWidget(log_output)
    
    buttons_layout = QHBoxLayout()
    start_btn = QPushButton("🚀 Skanirlemäni başlat")
    stop_btn = QPushButton("⏹️ Sakla")
    stop_btn.setEnabled(False)
    export_btn = QPushButton("📊 Netijeleri eksport et")
    
    buttons_layout.addWidget(start_btn)
    buttons_layout.addWidget(stop_btn)
    buttons_layout.addWidget(export_btn)
    buttons_layout.addStretch()
    
    layout.addLayout(buttons_layout)

    def append_log(message):
        log_output.append(message)

    def validate_ip_range(value):
        try:
            ipaddress.ip_network(value, strict=False)
            return True
        except ValueError:
            return False

    def workers_for_intensity():
        return {
            0: 5,
            1: 10,
            2: 20,
            3: 40,
        }.get(intensity_combo.currentIndex(), 10)

    def selected_adapter():
        return adapter_combo.currentData()

    def on_adapter_changed():
        adapter = selected_adapter()
        if adapter:
            ip_range_input.setText(adapter["network"])
            adapter_combo.setToolTip(adapter["label"])
            ip_range_input.setToolTip(adapter["network"])

    def add_device_to_table(device):
        row = results_table.rowCount()
        results_table.insertRow(row)

        ports = format_ports(device.get("ports"))
        vulnerable = "Hawa" if device.get("vulnerable") else "Ýok"
        recommendation = device.get("recommendation") or "Kritiki maslahat ýok."
        values = [
            localize_unknown(device.get("ip")),
            localize_unknown(device.get("name") or device.get("hostname")),
            localize_unknown(device.get("mac")),
            localize_unknown(device.get("vendor")),
            localize_unknown(device.get("device_type")),
            localize_unknown(device.get("discovery_method")),
            ports,
            risk_label(device.get("risk_level", "low")),
            vulnerable,
            recommendation,
        ]

        for column, value in enumerate(values):
            item = QTableWidgetItem(str(value))
            item.setToolTip(str(value))
            if column in (7, 8):
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            results_table.setItem(row, column, item)

        append_log(
            f"Gurluş tapyldy: {values[0]} | görnüşi: {values[4]} | "
            f"usul: {values[5]} | portlar: {ports} | töwekgelçilik: {values[7]}"
        )

    def on_progress(value, message):
        progress_bar.setValue(value)
        append_log(message)

    def on_error(message):
        append_log(f"Ýalňyşlyk: {message}")

    def on_finished(devices):
        widget.last_devices = devices
        start_btn.setEnabled(True)
        stop_btn.setEnabled(False)
        progress_bar.setValue(100)

        if devices:
            try:
                save_scan(devices)
                append_log(f"Skanirleme tamamlandy. Saklanan gurluşlar: {len(devices)}")
            except Exception as exc:
                QMessageBox.critical(widget, "Maglumat bazasynyň ýalňyşlygy", str(exc))
                append_log(f"Netijeleri saklamak başartmady: {exc}")
        else:
            append_log("Skanirleme tamamlandy, gurluş tapylmady.")

        widget.scan_thread = None

    def start_scan():
        ip_range = ip_range_input.text().strip()
        if not validate_ip_range(ip_range):
            QMessageBox.warning(widget, "Nädogry aralyk", "IP ýa-da tory 192.168.0.0/24 görnüşinde giriziň.")
            return

        network = ipaddress.ip_network(ip_range, strict=False)
        if not network.is_private:
            QMessageBox.warning(
                widget,
                "Daşarky tor",
                "Bu aralyk lokal içki tor däl. Öz Wi-Fi toruňyz üçin, mysal üçin, 192.168.0.0/24 ýa-da 192.168.1.0/24 ulanyň.",
            )
            return

        if widget.scan_thread and widget.scan_thread.isRunning():
            return

        widget.last_devices = []
        results_table.setRowCount(0)
        log_output.clear()
        progress_bar.setVisible(True)
        progress_bar.setValue(0)
        start_btn.setEnabled(False)
        stop_btn.setEnabled(True)

        adapter = selected_adapter()
        if adapter:
            append_log(f"Saýlanan adapter: {adapter['label']}")
            if adapter.get("is_vpn"):
                append_log("Üns beriň: VPN adapter saýlandy. Wi-Fi gurluşlaryny görmek üçin Wi-Fi adapterini saýlaň.")
        append_log(f"Skanirleme başlady: {ip_range}")
        thread = ScanThread(ip_range=ip_range, max_workers=workers_for_intensity(), adapter=adapter)
        thread.progress.connect(on_progress)
        thread.device_found.connect(add_device_to_table)
        thread.error.connect(on_error)
        thread.scan_finished.connect(on_finished)
        widget.scan_thread = thread
        thread.start()

    def stop_scan():
        if widget.scan_thread and widget.scan_thread.isRunning():
            widget.scan_thread.stop()
            stop_btn.setEnabled(False)
            append_log("Skanirleme saklanýar...")

    def export_results():
        if not widget.last_devices:
            QMessageBox.information(widget, "Maglumat ýok", "Ilki skanirlemäni ýerine ýetiriň.")
            return

        filename, _ = QFileDialog.getSaveFileName(
            widget,
            "Hasabaty sakla",
            "security_report.pdf",
            "PDF faýllar (*.pdf)"
        )
        if not filename:
            return

        try:
            generate_pdf(widget.last_devices, filename)
            append_log(f"Hasabat saklandy: {filename}")
            QMessageBox.information(widget, "Taýýar", "PDF hasabat üstünlikli döredildi.")
        except Exception as exc:
            QMessageBox.critical(widget, "Hasabat ýalňyşlygy", str(exc))
            append_log(f"Hasabat döretmek başartmady: {exc}")

    adapter_combo.currentIndexChanged.connect(on_adapter_changed)
    on_adapter_changed()
    start_btn.clicked.connect(start_scan)
    stop_btn.clicked.connect(stop_scan)
    export_btn.clicked.connect(export_results)
    
    return widget

def setup_dashboard_page():
    widget = QWidget()
    layout = QVBoxLayout(widget)
    layout.setSpacing(12)
    
    title = QLabel("Howpsuzlyk paneli")
    title.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
    layout.addWidget(title)

    stats_frame = QFrame()
    stats_frame.setObjectName("card")
    stats_layout = QGridLayout(stats_frame)
    layout.addWidget(stats_frame)

    overview_layout = QHBoxLayout()
    layout.addLayout(overview_layout, 1)

    latest_group = QGroupBox("Soňky skanirleme")
    latest_layout = QVBoxLayout(latest_group)
    latest_info = QLabel()
    latest_info.setWordWrap(True)
    latest_layout.addWidget(latest_info)
    latest_notes = QTextEdit()
    latest_notes.setReadOnly(True)
    latest_notes.setMaximumHeight(150)
    latest_layout.addWidget(latest_notes)
    overview_layout.addWidget(latest_group, 2)

    risk_group = QGroupBox("Töwekgelçilik paýlanyşy")
    risk_layout = QVBoxLayout(risk_group)
    risk_table = QTableWidget()
    configure_summary_table(risk_table, ["Dereje", "Sany"])
    risk_layout.addWidget(risk_table)
    overview_layout.addWidget(risk_group, 1)

    type_group = QGroupBox("Gurluş görnüşleri")
    type_layout = QVBoxLayout(type_group)
    type_table = QTableWidget()
    configure_summary_table(type_table, ["Görnüşi", "Sany"])
    type_layout.addWidget(type_table)
    overview_layout.addWidget(type_group, 1)

    refresh_btn = QPushButton("Täzele")
    layout.addWidget(refresh_btn, alignment=Qt.AlignmentFlag.AlignRight)

    def refresh():
        stats = get_scan_stats()
        history = load_history()
        latest = history[0] if history else None
        latest_devices = latest["devices"] if latest else []
        while stats_layout.count():
            item = stats_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        protected = max(len(latest_devices) - count_vulnerable(latest_devices), 0)
        cards = [
            ("Ähli skanirlemeler", str(stats["total_scans"])),
            ("Soňky gurluşlar", str(len(latest_devices))),
            ("Soňky gowşak", str(count_vulnerable(latest_devices))),
            ("Portly gurluşlar", str(count_with_ports(latest_devices))),
            ("Howpsuz gurluşlar", str(protected)),
            ("Soňky skanirleme", stats["last_scan"] or "Ýok"),
        ]
        for index, (card_title, value) in enumerate(cards):
            stats_layout.addWidget(create_card(card_title, value), index // 3, index % 3)

        if not latest:
            latest_info.setText("Heniz skanirleme ýok.")
            latest_notes.setPlainText("Skanirleme sahypasynda täze barlag başladyň.")
            risk_table.setRowCount(0)
            type_table.setRowCount(0)
            return

        latest_info.setText(
            f"Wagty: {latest['timestamp']} | Tor: {first_scan_value(latest, 'network')} | "
            f"Adapter: {first_scan_value(latest, 'adapter')}"
        )
        notes = []
        if count_vulnerable(latest_devices):
            notes.append("Gowşak gurluş bar: Netijeler sahypasynda maslahatlary barlaň.")
        if count_with_ports(latest_devices):
            notes.append("Açyk portly gurluşlar tapyldy, hyzmatlary aýratyn barlamak maslahat berilýär.")
        if len(latest_devices) <= 1:
            notes.append("Diňe bir gurluş tapyldy: Wi-Fi, Guest network ýa-da AP isolation sazlamasyny barlaň.")
        if not notes:
            notes.append("Soňky skanirlemede kritiki ýagdaý görünmeýär.")
        latest_notes.setPlainText("\n".join(f"- {note}" for note in notes))

        risk_counts = Counter(risk_label(device.get("risk_level", "low")) for device in latest_devices)
        fill_summary_table(risk_table, risk_counts.most_common())

        type_counts = Counter(device.get("device_type") or "Unknown" for device in latest_devices)
        fill_summary_table(type_table, type_counts.most_common())

    refresh_btn.clicked.connect(refresh)
    widget.refresh = refresh
    refresh()
    return widget

def setup_results_page():
    widget = QWidget()
    layout = QVBoxLayout(widget)
    layout.setSpacing(10)
    
    title = QLabel("Skanirleme netijeleri")
    title.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
    layout.addWidget(title)

    info_label = QLabel()
    info_label.setStyleSheet("color: #9aa0aa;")
    layout.addWidget(info_label)

    tools_layout = QHBoxLayout()
    tools_layout.addWidget(QLabel("Filtr:"))
    filter_combo = QComboBox()
    configure_combo(filter_combo, 18)
    filter_combo.addItems(["Ähli", "Gowşak", "Portly", "Router", "PC", "IoT kandidat", "Telefon/Unknown"])
    tools_layout.addWidget(filter_combo)
    tools_layout.addStretch()
    refresh_btn = QPushButton("Täzele")
    export_btn = QPushButton("PDF hasabat")
    tools_layout.addWidget(refresh_btn)
    tools_layout.addWidget(export_btn)
    layout.addLayout(tools_layout)

    table = create_devices_table()
    table.setMinimumHeight(360)
    layout.addWidget(table, 1)

    details_group = QGroupBox("Saýlanan gurluşyň seljermesi")
    details_layout = QVBoxLayout(details_group)
    details_box = QTextEdit()
    details_box.setReadOnly(True)
    details_box.setMaximumHeight(145)
    details_layout.addWidget(details_box)
    layout.addWidget(details_group)

    widget.latest_devices = []
    widget.visible_devices = []

    def apply_filter():
        selected = filter_combo.currentText()
        devices = widget.latest_devices
        if selected == "Gowşak":
            devices = [device for device in devices if device.get("vulnerable")]
        elif selected == "Portly":
            devices = [device for device in devices if device.get("ports")]
        elif selected != "Ähli":
            devices = [device for device in devices if device.get("device_type") == selected]

        widget.visible_devices = devices
        fill_devices_table(table, widget.visible_devices)
        if widget.visible_devices:
            table.selectRow(0)
            details_box.setPlainText(device_details_text(widget.visible_devices[0]))
        else:
            details_box.setPlainText("Bu filtr boýunça gurluş tapylmady.")

    def show_selected_device():
        row = table.currentRow()
        if row < 0 or row >= len(widget.visible_devices):
            details_box.setPlainText("Gurluş saýlanmady.")
            return
        details_box.setPlainText(device_details_text(widget.visible_devices[row]))

    def refresh():
        history = load_history()
        if not history:
            widget.latest_devices = []
            widget.visible_devices = []
            table.setRowCount(0)
            info_label.setText("Heniz saklanan skanirleme ýok. Ilki tor skanirlemesini ýerine ýetiriň.")
            details_box.setPlainText("Maglumat ýok.")
            export_btn.setEnabled(False)
            return

        latest = history[0]
        widget.latest_devices = latest["devices"]
        vulnerable_count = sum(1 for device in widget.latest_devices if device.get("vulnerable"))
        info_label.setText(
            f"Soňky skanirleme: {latest['timestamp']} | "
            f"tor: {first_scan_value(latest, 'network')} | "
            f"gurluşlar: {len(widget.latest_devices)} | gowşak: {vulnerable_count}"
        )
        export_btn.setEnabled(bool(widget.latest_devices))
        apply_filter()

    def export_latest():
        if not widget.latest_devices:
            QMessageBox.information(widget, "Maglumat ýok", "Eksport etmek üçin saklanan netije ýok.")
            return

        filename, _ = QFileDialog.getSaveFileName(
            widget,
            "Hasabaty sakla",
            "security_report.pdf",
            "PDF faýllar (*.pdf)",
        )
        if not filename:
            return

        try:
            generate_pdf(widget.latest_devices, filename)
            QMessageBox.information(widget, "Taýýar", "PDF hasabat üstünlikli döredildi.")
        except Exception as exc:
            QMessageBox.critical(widget, "Hasabat ýalňyşlygy", str(exc))

    refresh_btn.clicked.connect(refresh)
    export_btn.clicked.connect(export_latest)
    filter_combo.currentIndexChanged.connect(apply_filter)
    table.itemSelectionChanged.connect(show_selected_device)
    widget.refresh = refresh
    refresh()
    return widget

def setup_history_page():
    widget = QWidget()
    layout = QVBoxLayout(widget)
    layout.setSpacing(10)
    
    title = QLabel("Skanirleme taryhy")
    title.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
    layout.addWidget(title)

    top_layout = QHBoxLayout()
    info_label = QLabel()
    info_label.setStyleSheet("color: #9aa0aa;")
    refresh_btn = QPushButton("Täzele")
    clear_btn = QPushButton("Taryhy arassala")
    top_layout.addWidget(info_label)
    top_layout.addStretch()
    top_layout.addWidget(refresh_btn)
    top_layout.addWidget(clear_btn)
    layout.addLayout(top_layout)

    scans_table = QTableWidget(0, 7)
    scans_table.setHorizontalHeaderLabels(["Wagty", "Tor", "Adapter", "Gurluşlar", "Gowşak", "Portly", "Töwekgelçilik"])
    scans_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
    scans_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
    scans_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
    scans_table.setAlternatingRowColors(True)
    layout.addWidget(scans_table, 1)

    selected_group = QGroupBox("Saýlanan skanirlemäniň gysgaça mazmuny")
    selected_layout = QVBoxLayout(selected_group)
    selected_summary = QTextEdit()
    selected_summary.setReadOnly(True)
    selected_layout.addWidget(selected_summary)

    selected_actions = QHBoxLayout()
    export_selected_btn = QPushButton("Saýlanan skany PDF et")
    selected_actions.addWidget(export_selected_btn)
    selected_actions.addStretch()
    selected_layout.addLayout(selected_actions)
    layout.addWidget(selected_group)

    widget.history = []

    def show_scan_details(row=None):
        if row is None:
            row = scans_table.currentRow()
        if row < 0 or row >= len(widget.history):
            selected_summary.setPlainText("Skanirleme saýlanmady.")
            export_selected_btn.setEnabled(False)
            return
        selected_summary.setPlainText(scan_summary_text(widget.history[row]))
        export_selected_btn.setEnabled(bool(widget.history[row].get("devices")))

    def refresh():
        widget.history = load_history()
        scans_table.setRowCount(0)
        selected_summary.clear()

        if not widget.history:
            info_label.setText("Taryh boş. Skanirleme tamamlanandan soň ýazgylar şu ýerde peýda bolar.")
            clear_btn.setEnabled(False)
            export_selected_btn.setEnabled(False)
            return

        info_label.setText(f"Saklanan skanirlemeler: {len(widget.history)}")
        clear_btn.setEnabled(True)
        for scan in widget.history:
            devices = scan["devices"]
            row = scans_table.rowCount()
            scans_table.insertRow(row)
            values = [
                scan["timestamp"],
                first_scan_value(scan, "network"),
                first_scan_value(scan, "adapter"),
                str(len(devices)),
                str(sum(1 for device in devices if device.get("vulnerable"))),
                str(sum(1 for device in devices if device.get("ports"))),
                risk_label(highest_risk(devices)),
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                if column:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                scans_table.setItem(row, column, item)

        scans_table.selectRow(0)
        show_scan_details(0)

    def export_selected_scan():
        row = scans_table.currentRow()
        if row < 0 or row >= len(widget.history):
            QMessageBox.information(widget, "Maglumat ýok", "Eksport etmek üçin skanirleme saýlaň.")
            return

        devices = widget.history[row].get("devices", [])
        if not devices:
            QMessageBox.information(widget, "Maglumat ýok", "Saýlanan skanirlemede gurluş ýok.")
            return

        filename, _ = QFileDialog.getSaveFileName(
            widget,
            "Hasabaty sakla",
            f"security_report_{widget.history[row]['timestamp'].replace(':', '-')}.pdf",
            "PDF faýllar (*.pdf)",
        )
        if not filename:
            return

        try:
            generate_pdf(devices, filename)
            QMessageBox.information(widget, "Taýýar", "Saýlanan skan boýunça PDF döredildi.")
        except Exception as exc:
            QMessageBox.critical(widget, "Hasabat ýalňyşlygy", str(exc))

    def clear_all_history():
        if QMessageBox.question(
            widget,
            "Taryhy arassalamak",
            "Ähli skanirleme taryhyny pozmalymy?",
        ) != QMessageBox.StandardButton.Yes:
            return
        if clear_history():
            refresh()
        else:
            QMessageBox.warning(widget, "Ýalňyşlyk", "Taryhy arassalamak başartmady.")

    refresh_btn.clicked.connect(refresh)
    clear_btn.clicked.connect(clear_all_history)
    export_selected_btn.clicked.connect(export_selected_scan)
    scans_table.itemSelectionChanged.connect(show_scan_details)
    widget.refresh = refresh
    refresh()
    return widget

def setup_settings_page():
    widget = QWidget()
    layout = QVBoxLayout(widget)
    
    title = QLabel("Ulgam sazlamalary")
    title.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
    layout.addWidget(title)
    
    # Здесь будут настройки
    settings_content = QLabel("Ulgam sazlamalary şu ýerde elýeterli bolar...")
    settings_content.setAlignment(Qt.AlignmentFlag.AlignCenter)
    layout.addWidget(settings_content)
    
    return widget

def create_card(title, content, color=None):
    card = QFrame()
    card.setObjectName("card")
    card.setMinimumHeight(120)
    
    layout = QVBoxLayout(card)
    
    title_label = QLabel(title)
    title_label.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
    
    content_label = QLabel(content)
    content_label.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
    
    if color:
        content_label.setStyleSheet(f"color: {color};")
    
    layout.addWidget(title_label)
    layout.addWidget(content_label)
    layout.addStretch()
    
    return card
