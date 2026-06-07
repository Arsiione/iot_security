# IoT Security Scanner (IoT Howpsuzlyk Skaneri)

A modular, multi-threaded desktop application designed to scan local networks, detect active devices, analyze open ports, and assess security vulnerabilities in IoT devices. Built with **Python 3**, **PyQt6**, **Nmap**, **SQLite**, and **ReportLab**.

The application interface is fully localized in the **Turkmen** language, making it suitable for educational projects, research, and local network audits.

---

## 🌟 Key Features

- 🔍 **Hybrid Host Discovery**:
  - Leverages **Nmap** (via `python-nmap`) for high-speed network scanning.
  - Gracefully falls back to a custom multi-threaded **Ping Sweep** (`ThreadPoolExecutor`) and system **ARP cache parsing** if Nmap is not installed.
  - Intelligent local network adapter detection (auto-filters VPNs and virtual interfaces, prioritizes Wi-Fi/Ethernet).
- 🛡️ **Vulnerability Assessment (Plugin-Based)**:
  - Dynamic loading of checks from the `plugins/` directory.
  - **Weak Credentials Check**: Tests Telnet (port 23) using a dictionary of common default passwords.
  - **HTTP Interface Audit**: Identifies open web panels on IoT devices (e.g., IP cameras).
- 🛠️ **Remediation Suggestions & Engine**:
  - Provides actionable tips for fixing detected issues.
  - Includes a prototype remediation engine for automated credential updates (SSH/Web API hardening).
- 📊 **Interactive GUI & Dashboard**:
  - Clean and modern PyQt6-based user interface with native Dark and Light modes.
  - Live statistics and charts showing risk distribution and device types.
- 📋 **Reporting & Notifications**:
  - Exports comprehensive security reports to **PDF** (including Matplotlib-generated status charts).
  - Sends immediate notifications via **Email (SMTP)** and **Webhooks** when vulnerabilities are found.
- 💾 **Scan History**:
  - Stores all scan history in a local SQLite database (`iot_security.db`) with automatic migrations support.

---

## 📂 Project Structure

```
├── config/
│   └── settings.py          # Configuration manager (config.json loader/saver)
├── core/
│   ├── database.py          # Database models (duplicate helper)
│   ├── remediation_engine.py# SSH/Web credential hardening logic
│   └── scanner.py           # Core scanning thread and host discovery engine
├── database.py              # Main SQLite operations, history, and stats
├── plugins/
│   ├── hikvision.py         # HTTP port check and device classification
│   └── telnet_weak_auth.py  # Dictionary attack check on Port 23
├── ui/
│   ├── main_window.py       # Main GUI structure & theme styling (QSS)
│   └── pages.py             # UI pages (Dashboard, Scan, Results, History, Settings)
├── utils/
│   └── notifications.py     # HTML Email, Webhook, and SMS alert manager
├── main.py                  # Entry point of the PyQt6 application
├── requirements.txt         # Project dependencies
└── .gitignore               # Excludes DBs, temp logs, virtual environments, and caches
```

---

## 🚀 Getting Started

### 📋 Prerequisites

1. **Python 3.10+**
2. **Nmap** (Optional but highly recommended for fast scanning):
   - **Windows**: Download and install from [nmap.org](https://nmap.org/download.html). Ensure it's in your system PATH or installed in default directories (`C:\Program Files\Nmap`).
   - **Linux**: `sudo apt install nmap`
   - **macOS**: `brew install nmap`

### 💻 Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/Arsiione/iot_security.git
   cd iot_security
   ```

2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Launch the application:
   ```bash
   python main.py
   ```

---

## 🛠️ Configuration

The application generates a `config.json` file on first run. You can configure:
- **Scan Settings**: Default IP ranges, thread count, custom port lists.
- **Alerts**: SMTP email settings (Gmail, Custom Mail) and Webhook URLs.
- **Remediation**: Toggle auto-fix and configuration backup before applying fixes.

---

## 📜 License

This project is open-source and available under the MIT License.
