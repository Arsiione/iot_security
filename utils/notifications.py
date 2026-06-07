# utils/notifications.py
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import json
from typing import List, Dict, Any
import logging
from datetime import datetime
import requests

class NotificationManager:
    def __init__(self, config: Dict[str, Any]):
        self.config = config
        self.logger = logging.getLogger(__name__)

    def send_security_alert(self, devices: List[Dict[str, Any]], scan_id: str):
        """Отправка уведомления о безопасности"""
        vulnerable_devices = [d for d in devices if d.get('vulnerabilities')]
        
        if not vulnerable_devices:
            return
        
        # Email уведомление
        if self.config.get('email_enabled', False):
            self.send_email_alert(vulnerable_devices, scan_id)
        
        # Webhook уведомление
        if self.config.get('webhook_enabled', False):
            self.send_webhook_alert(vulnerable_devices, scan_id)
        
        # SMS уведомление (если настроено)
        if self.config.get('sms_enabled', False):
            self.send_sms_alert(vulnerable_devices, scan_id)

    def send_email_alert(self, devices: List[Dict[str, Any]], scan_id: str):
        """Отправка email уведомления"""
        try:
            msg = MIMEMultipart()
            msg['Subject'] = f"IoT Security Alert - {len(devices)} vulnerable devices found"
            msg['From'] = self.config.get('smtp_user')
            msg['To'] = self.config.get('alert_email')
            
            # HTML содержимое
            html = self.generate_email_content(devices, scan_id)
            msg.attach(MIMEText(html, 'html'))
            
            with smtplib.SMTP(self.config.get('smtp_server'), self.config.get('smtp_port')) as server:
                server.starttls()
                server.login(self.config.get('smtp_user'), self.config.get('smtp_pass'))
                server.send_message(msg)
                
        except Exception as e:
            self.logger.error(f"Failed to send email alert: {e}")

    def generate_email_content(self, devices: List[Dict[str, Any]], scan_id: str) -> str:
        """Генерация HTML содержимого для email"""
        html = f"""
        <html>
        <body>
            <h2>IoT Security Scan Report</h2>
            <p>Scan ID: {scan_id}</p>
            <p>Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            <p>Found {len(devices)} vulnerable devices:</p>
            <table border="1">
                <tr>
                    <th>Device</th>
                    <th>IP</th>
                    <th>Vulnerabilities</th>
                    <th>Risk Level</th>
                </tr>
        """
        
        for device in devices:
            html += f"""
                <tr>
                    <td>{device.get('hostname', 'Unknown')}</td>
                    <td>{device.get('ip')}</td>
                    <td>{len(device.get('vulnerabilities', []))}</td>
                    <td style="color: {self.get_risk_color(device.get('risk_level', 'low'))}">
                        {device.get('risk_level', 'low')}
                    </td>
                </tr>
            """
        
        html += """
            </table>
            </body>
            </html>
        """
        
        return html

    def get_risk_color(self, risk_level: str) -> str:
        """Получение цвета для уровня риска"""
        colors = {
            'low': 'green',
            'medium': 'orange',
            'high': 'red',
            'critical': 'darkred'
        }
        return colors.get(risk_level, 'black')

    def send_webhook_alert(self, devices: List[Dict[str, Any]], scan_id: str):
        """Отправка webhook-уведомления."""
        webhook_url = self.config.get('webhook_url')
        if not webhook_url:
            self.logger.warning("Webhook enabled, but webhook_url is empty")
            return

        payload = {
            "scan_id": scan_id,
            "timestamp": datetime.now().isoformat(timespec="seconds"),
            "vulnerable_devices": [
                {
                    "ip": device.get("ip"),
                    "name": device.get("name") or device.get("hostname"),
                    "risk_level": device.get("risk_level", "low"),
                    "vulnerabilities": device.get("vulnerabilities", []),
                }
                for device in devices
            ],
        }

        try:
            response = requests.post(webhook_url, json=payload, timeout=10)
            response.raise_for_status()
        except Exception as e:
            self.logger.error(f"Failed to send webhook alert: {e}")

    def send_sms_alert(self, devices: List[Dict[str, Any]], scan_id: str):
        """Заглушка для SMS: провайдера SMS пока нет в настройках."""
        self.logger.info(
            "SMS alert requested for scan %s with %s devices, but SMS provider is not configured",
            scan_id,
            len(devices),
        )
