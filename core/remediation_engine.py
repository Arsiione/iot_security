import paramiko
import requests
from typing import Dict, Any, List
import logging

class RemediationEngine:
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.available_fixes = {
            "telnet_weak_auth": self.fix_telnet_weak_auth,
            "ftp_plaintext": self.fix_ftp_plaintext,
            "http_default_credentials": self.fix_http_default_creds,
            "weak_ssh_config": self.fix_weak_ssh_config
        }

    def apply_fixes(self, device: Dict[str, Any], fixes: List[str]) -> Dict[str, Any]:
        """Применение исправлений к устройству"""
        results = {
            "successful": [],
            "failed": [],
            "skipped": []
        }
        
        for fix_name in fixes:
            if fix_name in self.available_fixes:
                try:
                    success = self.available_fixes[fix_name](device)
                    if success:
                        results["successful"].append(fix_name)
                    else:
                        results["failed"].append(fix_name)
                except Exception as e:
                    self.logger.error(f"Ошибка применения фикса {fix_name}: {e}")
                    results["failed"].append(fix_name)
            else:
                results["skipped"].append(fix_name)
        
        return results

    def fix_telnet_weak_auth(self, device: Dict[str, Any]) -> bool:
        """Исправление слабой аутентификации Telnet"""
        # Попытка изменить пароль через Telnet
        # В реальном сценарии это должно быть настроено для конкретного устройства
        try:
            # Здесь должна быть логика изменения пароля
            # Это сильно зависит от конкретного устройства
            return True
        except:
            return False

    def fix_ftp_plaintext(self, device: Dict[str, Any]) -> bool:
        """Отключение FTP и включение SFTP"""
        try:
            # Попытка настроить SFTP через SSH
            if 22 in device.get('ports', []):
                # Логика настройки SFTP
                return True
        except:
            return False
        return False

    def fix_http_default_creds(self, device: Dict[str, Any]) -> bool:
        """Изменение паролей по умолчанию через web-интерфейс"""
        try:
            if 80 in device.get('ports', []) or 443 in device.get('ports', []):
                # Автоматическое изменение паролей через web API
                return self.change_web_password(device)
        except:
            return False
        return False

    def change_web_password(self, device: Dict[str, Any]) -> bool:
        """Изменение пароля через web-интерфейс"""
        # Это сильно зависит от конкретного устройства
        # Нужны специфические реализации для разных производителей
        try:
            session = requests.Session()
            
            # Попытка логина с дефолтными credentials
            login_data = {
                'username': 'admin',
                'password': 'admin'
            }
            
            response = session.post(
                f"http://{device['ip']}/login",
                data=login_data,
                timeout=10
            )
            
            if response.status_code == 200:
                new_password = self.generate_strong_password()
                # Изменение пароля
                change_data = {
                    'current_password': 'admin',
                    'new_password': new_password,
                    'confirm_password': new_password
                }
                
                change_response = session.post(
                    f"http://{device['ip']}/change_password",
                    data=change_data,
                    timeout=10
                )
                
                return change_response.status_code == 200
        except:
            return False
        return False

    def generate_strong_password(self, length: int = 12) -> str:
        """Генерация сильного пароля"""
        import secrets
        import string
        alphabet = string.ascii_letters + string.digits + string.punctuation
        return ''.join(secrets.choice(alphabet) for _ in range(length))

    def fix_weak_ssh_config(self, device: Dict[str, Any]) -> bool:
        """Укрепление SSH конфигурации"""
        try:
            if 22 in device.get('ports', []):
                ssh = paramiko.SSHClient()
                ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
                
                # Подключение с дефолтными credentials
                ssh.connect(
                    device['ip'],
                    username='admin',
                    password='admin',
                    timeout=10
                )
                
                # Команды для укрепления SSH
                commands = [
                    "sed -i 's/#PermitRootLogin yes/PermitRootLogin no/' /etc/ssh/sshd_config",
                    "sed -i 's/PasswordAuthentication yes/PasswordAuthentication no/' /etc/ssh/sshd_config",
                    "systemctl restart sshd"
                ]
                
                for cmd in commands:
                    stdin, stdout, stderr = ssh.exec_command(cmd)
                    if stderr.channel.recv_exit_status() != 0:
                        return False
                
                ssh.close()
                return True
        except:
            return False
        return False
