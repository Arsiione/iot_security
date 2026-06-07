# database.py
import sqlite3
from datetime import datetime
import os

DB_PATH = "iot_security.db"

SCAN_COLUMNS = {
    "timestamp": "TEXT",
    "ip": "TEXT",
    "name": "TEXT",
    "mac": "TEXT",
    "ports": "TEXT",
    "vulnerable": "BOOLEAN",
    "risk_level": "TEXT",
    "recommendation": "TEXT",
    "discovery_method": "TEXT",
    "device_type": "TEXT",
    "vendor": "TEXT",
    "adapter": "TEXT",
    "network": "TEXT",
}

def init_db():
    """Инициализация базы данных"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS scans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT,
            ip TEXT,
            name TEXT,
            mac TEXT,
            ports TEXT,
            vulnerable BOOLEAN,
            risk_level TEXT,
            recommendation TEXT,
            discovery_method TEXT,
            device_type TEXT,
            vendor TEXT,
            adapter TEXT,
            network TEXT
        )
    """)
    migrate_scans_table(c)
    conn.commit()
    conn.close()

def migrate_scans_table(cursor):
    """Добавляет недостающие колонки, если база была создана старой версией."""
    cursor.execute("PRAGMA table_info(scans)")
    existing_columns = {row[1] for row in cursor.fetchall()}

    for column, column_type in SCAN_COLUMNS.items():
        if column not in existing_columns:
            cursor.execute(f"ALTER TABLE scans ADD COLUMN {column} {column_type}")

def save_scan(devices):
    """Сохранение результатов сканирования"""
    init_db()
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    for device in devices:
        # Преобразуем список портов в строку для хранения
        ports_str = ",".join(map(str, device.get("ports", [])))
        name = device.get("name") or device.get("hostname") or device.get("ip", "Unknown")
        recommendations = device.get("recommendations", [])
        recommendation = device.get("recommendation", "")
        if not recommendation and recommendations:
            recommendation = "; ".join(
                rec.get("details") or rec.get("action") or str(rec)
                for rec in recommendations
            )
        
        c.execute(
            """INSERT INTO scans 
            (timestamp, ip, name, mac, ports, vulnerable, risk_level, recommendation,
             discovery_method, device_type, vendor, adapter, network) 
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                timestamp, 
                device.get("ip", "Unknown"),
                name,
                device.get("mac", "Unknown"),
                ports_str,
                device.get("vulnerable", False),
                device.get("risk_level", "low"),
                recommendation,
                device.get("discovery_method", ""),
                device.get("device_type", "Unknown"),
                device.get("vendor", "Näbelli"),
                device.get("adapter", ""),
                device.get("network", "")
            )
        )
    
    conn.commit()
    conn.close()

def load_history():
    """Загрузка истории сканирований"""
    init_db()
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    c.execute("SELECT DISTINCT timestamp FROM scans ORDER BY timestamp DESC")
    timestamps = [row[0] for row in c.fetchall()]
    
    history = []
    for ts in timestamps:
        c.execute("""
            SELECT ip, name, mac, ports, vulnerable, risk_level, recommendation,
                   discovery_method, device_type, vendor, adapter, network
            FROM scans WHERE timestamp = ? ORDER BY ip
        """, (ts,))
        
        devices = []
        for row in c.fetchall():
            # Преобразуем строку портов обратно в список чисел
            ports = [int(port) for port in row[3].split(',')] if row[3] else []
            
            devices.append({
                "ip": row[0],
                "name": row[1],
                "mac": row[2],
                "ports": ports,
                "vulnerable": bool(row[4]),
                "risk_level": row[5],
                "recommendation": row[6],
                "discovery_method": row[7] or "",
                "device_type": row[8] or "Unknown",
                "vendor": row[9] or "Näbelli",
                "adapter": row[10] or "",
                "network": row[11] or "",
            })
        
        history.append({"timestamp": ts, "devices": devices})
    
    conn.close()
    return history

def get_scan_stats():
    """Получение статистики по сканированиям"""
    init_db()
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    stats = {
        "total_scans": 0,
        "total_devices": 0,
        "vulnerable_devices": 0,
        "last_scan": None
    }
    
    # Количество сканирований
    c.execute("SELECT COUNT(DISTINCT timestamp) FROM scans")
    stats["total_scans"] = c.fetchone()[0] or 0
    
    # Общее количество устройств
    c.execute("SELECT COUNT(*) FROM scans")
    stats["total_devices"] = c.fetchone()[0] or 0
    
    # Уязвимые устройства
    c.execute("SELECT COUNT(*) FROM scans WHERE vulnerable = 1")
    stats["vulnerable_devices"] = c.fetchone()[0] or 0
    
    # Последнее сканирование
    c.execute("SELECT MAX(timestamp) FROM scans")
    stats["last_scan"] = c.fetchone()[0]
    
    conn.close()
    return stats

def clear_history():
    """Очистка истории сканирований"""
    try:
        if os.path.exists(DB_PATH):
            os.remove(DB_PATH)
        return True
    except:
        return False
