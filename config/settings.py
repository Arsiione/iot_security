# config/settings.py
import json
import os
from typing import Dict, Any

class ConfigManager:
    def __init__(self, config_file: str = "config.json"):
        self.config_file = config_file
        self.default_config = {
            "scanning": {
                "default_ip_range": "192.168.0.0/24",
                "scan_timeout": 5,
                "max_threads": 15,
                "common_ports": [21, 22, 23, 80, 443, 554, 8000, 8080, 8888, 1883, 8883]
            },
            "notifications": {
                "email_enabled": False,
                "smtp_server": "smtp.gmail.com",
                "smtp_port": 587,
                "smtp_user": "",
                "smtp_pass": "",
                "alert_email": "",
                "webhook_enabled": False,
                "webhook_url": ""
            },
            "remediation": {
                "auto_fix": False,
                "allowed_fixes": ["telnet_weak_auth", "ftp_plaintext"],
                "backup_before_fix": True
            },
            "database": {
                "path": "iot_security.db",
                "keep_scans_days": 30
            }
        }
        self.config = self.load_config()

    def load_config(self) -> Dict[str, Any]:
        """Загрузка конфигурации из файла"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    return json.load(f)
            except:
                return self.default_config
        return self.default_config

    def save_config(self):
        """Сохранение конфигурации в файл"""
        try:
            with open(self.config_file, 'w') as f:
                json.dump(self.config, f, indent=2)
            return True
        except:
            return False

    def get(self, key: str, default=None):
        """Получение значения конфигурации"""
        keys = key.split('.')
        value = self.config
        for k in keys:
            if isinstance(value, dict) and k in value:
                value = value[k]
            else:
                return default
        return value

    def set(self, key: str, value):
        """Установка значения конфигурации"""
        keys = key.split('.')
        config = self.config
        for k in keys[:-1]:
            if k not in config:
                config[k] = {}
            config = config[k]
        config[keys[-1]] = value
        self.save_config()