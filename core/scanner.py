from PyQt6.QtCore import QThread, pyqtSignal

import importlib.util
import ipaddress
import os
import platform
import re
import shutil
import socket
import subprocess
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Set


COMMON_VENDOR_OUIS = {
    "00:03:93": "Apple",
    "00:05:02": "Apple",
    "00:0a:27": "Apple",
    "00:0a:95": "Apple",
    "00:0d:93": "Apple",
    "00:16:cb": "Apple",
    "00:17:f2": "Apple",
    "00:19:e3": "Apple",
    "00:1b:63": "Apple",
    "00:1e:c2": "Apple",
    "00:1f:f3": "Apple",
    "00:21:e9": "Apple",
    "00:22:41": "Apple",
    "00:23:12": "Apple",
    "00:23:32": "Apple",
    "00:23:df": "Apple",
    "00:24:36": "Apple",
    "00:25:00": "Apple",
    "00:25:4b": "Apple",
    "00:25:bc": "Apple",
    "00:26:08": "Apple",
    "00:26:4a": "Apple",
    "00:26:b0": "Apple",
    "04:0c:ce": "Apple",
    "04:15:52": "Apple",
    "04:1e:64": "Apple",
    "04:26:65": "Apple",
    "04:4b:ed": "Apple",
    "04:52:f3": "Apple",
    "04:54:53": "Apple",
    "04:69:f8": "Apple",
    "04:db:56": "Apple",
    "08:00:07": "Apple",
    "08:70:45": "Apple",
    "0c:30:21": "Apple",
    "10:40:f3": "Apple",
    "14:10:9f": "Apple",
    "18:34:51": "Apple",
    "1c:1a:c0": "Apple",
    "20:a2:e4": "Apple",
    "24:a0:74": "Apple",
    "28:cf:e9": "Apple",
    "2c:1f:23": "Apple",
    "30:10:e4": "Apple",
    "34:12:98": "Apple",
    "38:48:4c": "Apple",
    "3c:15:c2": "Apple",
    "40:30:04": "Apple",
    "44:d8:84": "Apple",
    "48:43:7c": "Apple",
    "4c:57:ca": "Apple",
    "50:ed:3c": "Apple",
    "54:26:96": "Apple",
    "58:55:ca": "Apple",
    "5c:95:ae": "Apple",
    "60:33:4b": "Apple",
    "64:20:0c": "Apple",
    "68:96:7b": "Apple",
    "6c:40:08": "Apple",
    "70:56:81": "Apple",
    "78:31:c1": "Apple",
    "7c:50:49": "Apple",
    "80:49:71": "Apple",
    "88:53:95": "Apple",
    "8c:85:90": "Apple",
    "90:27:e4": "Apple",
    "94:94:26": "Apple",
    "98:01:a7": "Apple",
    "9c:fc:01": "Apple",
    "a0:99:9b": "Apple",
    "a4:5e:60": "Apple",
    "a8:20:66": "Apple",
    "ac:bc:32": "Apple",
    "b0:34:95": "Apple",
    "b4:f0:ab": "Apple",
    "b8:09:8a": "Apple",
    "bc:52:b7": "Apple",
    "c0:84:7a": "Apple",
    "c8:33:4b": "Apple",
    "cc:08:e0": "Apple",
    "d0:23:db": "Apple",
    "d4:9a:20": "Apple",
    "d8:30:62": "Apple",
    "dc:2b:2a": "Apple",
    "e0:b5:2d": "Apple",
    "e4:ce:8f": "Apple",
    "e8:06:88": "Apple",
    "ec:35:86": "Apple",
    "f0:18:98": "Apple",
    "f4:31:c3": "Apple",
    "f8:1e:df": "Apple",
    "fc:25:3f": "Apple",
    "00:12:fb": "Samsung",
    "00:15:99": "Samsung",
    "00:16:32": "Samsung",
    "00:17:c9": "Samsung",
    "00:1a:8a": "Samsung",
    "00:1d:25": "Samsung",
    "00:1e:7d": "Samsung",
    "00:21:19": "Samsung",
    "00:23:39": "Samsung",
    "00:23:d6": "Samsung",
    "00:26:37": "Samsung",
    "04:fe:31": "Samsung",
    "08:08:c2": "Samsung",
    "0c:14:20": "Samsung",
    "10:30:47": "Samsung",
    "14:89:fd": "Samsung",
    "18:3a:2d": "Samsung",
    "1c:5a:3e": "Samsung",
    "20:02:af": "Samsung",
    "24:4b:03": "Samsung",
    "28:27:bf": "Samsung",
    "30:07:4d": "Samsung",
    "34:23:ba": "Samsung",
    "38:aa:3c": "Samsung",
    "40:0e:85": "Samsung",
    "44:4e:1a": "Samsung",
    "48:5a:3f": "Samsung",
    "50:32:75": "Samsung",
    "58:1f:aa": "Samsung",
    "5c:0a:5b": "Samsung",
    "60:6b:bd": "Samsung",
    "68:eb:c5": "Samsung",
    "70:f9:27": "Samsung",
    "78:1f:db": "Samsung",
    "80:18:a7": "Samsung",
    "84:51:81": "Samsung",
    "88:32:9b": "Samsung",
    "8c:77:12": "Samsung",
    "90:18:7c": "Samsung",
    "94:51:bf": "Samsung",
    "98:52:b1": "Samsung",
    "9c:02:98": "Samsung",
    "a0:21:95": "Samsung",
    "a8:9f:ba": "Samsung",
    "ac:5f:3e": "Samsung",
    "b0:ec:71": "Samsung",
    "b8:5e:7b": "Samsung",
    "bc:72:b1": "Samsung",
    "c0:bd:d1": "Samsung",
    "c8:19:f7": "Samsung",
    "cc:07:ab": "Samsung",
    "d0:17:6a": "Samsung",
    "d8:90:e8": "Samsung",
    "dc:71:44": "Samsung",
    "e4:7c:f9": "Samsung",
    "e8:50:8b": "Samsung",
    "ec:1f:72": "Samsung",
    "f0:25:b7": "Samsung",
    "f4:42:8f": "Samsung",
    "f8:04:2e": "Samsung",
    "fc:19:10": "Samsung",
    "00:9a:cd": "Huawei",
    "04:bd:70": "Huawei",
    "08:19:a6": "Huawei",
    "10:1b:54": "Huawei",
    "14:b9:68": "Huawei",
    "18:c5:8a": "Huawei",
    "20:08:ed": "Huawei",
    "28:3c:e4": "Huawei",
    "30:87:30": "Huawei",
    "34:29:12": "Huawei",
    "38:37:8b": "Huawei",
    "40:4d:8e": "Huawei",
    "44:6a:2e": "Huawei",
    "48:46:fb": "Huawei",
    "54:25:ea": "Huawei",
    "58:2a:f7": "Huawei",
    "60:de:44": "Huawei",
    "68:a0:f6": "Huawei",
    "70:72:3c": "Huawei",
    "78:d7:52": "Huawei",
    "80:71:7a": "Huawei",
    "84:a8:e4": "Huawei",
    "88:cf:98": "Huawei",
    "90:67:1c": "Huawei",
    "9c:28:40": "Huawei",
    "a0:8d:16": "Huawei",
    "a4:dc:be": "Huawei",
    "ac:e2:15": "Huawei",
    "b0:e5:ed": "Huawei",
    "b4:30:52": "Huawei",
    "bc:76:70": "Huawei",
    "c0:70:09": "Huawei",
    "c8:d1:5e": "Huawei",
    "cc:53:b5": "Huawei",
    "d0:7a:b5": "Huawei",
    "d8:49:0b": "Huawei",
    "dc:d2:fc": "Huawei",
    "e0:19:1d": "Huawei",
    "e4:a7:c5": "Huawei",
    "e8:cd:2d": "Huawei",
    "f0:43:47": "Huawei",
    "f4:dc:4d": "Huawei",
    "fc:48:ef": "Huawei",
    "18:59:36": "Xiaomi",
    "28:6c:07": "Xiaomi",
    "34:80:b3": "Xiaomi",
    "38:a4:ed": "Xiaomi",
    "50:8f:4c": "Xiaomi",
    "64:09:80": "Xiaomi",
    "68:df:dd": "Xiaomi",
    "74:51:ba": "Xiaomi",
    "78:02:f8": "Xiaomi",
    "8c:be:be": "Xiaomi",
    "98:fa:e3": "Xiaomi",
    "a0:86:c6": "Xiaomi",
    "ac:c1:ee": "Xiaomi",
    "b0:e2:35": "Xiaomi",
    "c4:0b:cb": "Xiaomi",
    "d4:97:0b": "Xiaomi",
    "e4:46:da": "Xiaomi",
    "f8:a4:5f": "Xiaomi",
    "70:3a:51": "Intel",
    "1c:1b:b5": "Intel",
    "e8:6a:64": "Intel",
    "18:0f:76": "Router",
}


def run_hidden_command(args: List[str], timeout: int = 30) -> subprocess.CompletedProcess:
    kwargs = {
        "capture_output": True,
        "text": True,
        "timeout": timeout,
        "stdin": subprocess.DEVNULL,
        "encoding": "oem" if platform.system() == "Windows" else "utf-8",
        "errors": "ignore",
    }

    if platform.system() == "Windows":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        kwargs["startupinfo"] = startupinfo
        kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)

    return subprocess.run(args, **kwargs)


def mask_to_prefix(mask: str) -> Optional[int]:
    try:
        return ipaddress.IPv4Network(f"0.0.0.0/{mask}").prefixlen
    except Exception:
        return None


def normalize_mac(mac: str) -> str:
    if not mac or mac == "Unknown":
        return "Unknown"
    mac = mac.strip().replace("-", ":").lower()
    if re.fullmatch(r"[0-9a-f]{2}(?::[0-9a-f]{2}){5}", mac):
        return mac
    return "Unknown"


def vendor_from_mac(mac: str) -> str:
    mac = normalize_mac(mac)
    if mac == "Unknown":
        return "Näbelli"

    first_octet = int(mac.split(":")[0], 16)
    if first_octet & 0x02:
        return "Private MAC"

    return COMMON_VENDOR_OUIS.get(":".join(mac.split(":")[:3]), "Näbelli")


def parse_ipv4(value: str) -> Optional[str]:
    match = re.search(r"(\d{1,3}(?:\.\d{1,3}){3})", value or "")
    if not match:
        return None
    try:
        return str(ipaddress.ip_address(match.group(1)))
    except ValueError:
        return None


def parse_route_gateways() -> Dict[str, str]:
    try:
        result = run_hidden_command(["route", "print", "-4"], timeout=8)
    except Exception:
        return {}

    gateways = {}
    for line in result.stdout.splitlines():
        parts = line.split()
        if len(parts) >= 5 and parts[0] == "0.0.0.0" and parts[1] == "0.0.0.0":
            gateway = parse_ipv4(parts[2])
            interface_ip = parse_ipv4(parts[3])
            if gateway and interface_ip:
                gateways.setdefault(interface_ip, gateway)
    return gateways


def parse_ipconfig_adapters() -> List[Dict[str, Any]]:
    try:
        result = run_hidden_command(["ipconfig", "/all"], timeout=10)
    except Exception:
        return []

    route_gateways = parse_route_gateways()
    adapters = []
    current = None

    def finish_current():
        if not current:
            return
        adapter = build_adapter(current, route_gateways)
        if adapter:
            adapters.append(adapter)

    for raw_line in result.stdout.splitlines():
        line = raw_line.rstrip()
        header_match = re.match(r"^.*(?:adapter|адаптер)\s+(.+):$", line, flags=re.IGNORECASE)
        if header_match:
            finish_current()
            current = {
                "name": header_match.group(1).strip(),
                "raw_header": line.strip(),
                "default_gateway_candidates": [],
            }
            continue

        if current is None or ":" not in line:
            continue

        key, value = line.split(":", 1)
        key = re.sub(r"\s*\.+\s*", " ", key).strip().lower()
        value = value.strip()

        if "media state" in key or "состояние среды" in key:
            current["media_state"] = value
        elif "description" in key or "описание" in key:
            current["description"] = value
        elif "physical address" in key or "физический адрес" in key:
            current["mac"] = normalize_mac(value)
        elif "ipv4 address" in key or ("ipv4" in key and "адрес" in key):
            current["ip"] = parse_ipv4(value)
        elif "subnet mask" in key or "маска подсети" in key:
            current["mask"] = parse_ipv4(value)
        elif "default gateway" in key or "основной шлюз" in key:
            gateway = parse_ipv4(value)
            if gateway:
                current["default_gateway_candidates"].append(gateway)
        elif "dhcp server" in key or "dhcp-сервер" in key:
            current["dhcp_server"] = parse_ipv4(value)
        elif current.get("default_gateway_candidates") is not None:
            # ipconfig sometimes prints additional gateway lines without a clear key.
            gateway = parse_ipv4(value)
            if gateway and "default_gateway_candidates" in current:
                current["default_gateway_candidates"].append(gateway)

    finish_current()
    return sorted(adapters, key=lambda item: item["score"], reverse=True)


def build_adapter(raw: Dict[str, Any], route_gateways: Dict[str, str]) -> Optional[Dict[str, Any]]:
    name = raw.get("name", "")
    description = raw.get("description", "")
    text = f"{name} {description} {raw.get('raw_header', '')}".lower()
    ip = raw.get("ip")
    mask = raw.get("mask")

    media_state = raw.get("media_state", "").lower()
    if media_state.startswith("media disconnected") or "недоступ" in media_state:
        return None
    if not ip or not mask:
        return None

    try:
        address = ipaddress.ip_address(ip)
    except ValueError:
        return None
    if address.is_loopback or address.is_link_local:
        return None
    if "loopback" in text or "npcap" in text:
        return None

    prefix = mask_to_prefix(mask)
    if prefix is None:
        return None

    network = ipaddress.ip_network(f"{ip}/{prefix}", strict=False)
    gateway = route_gateways.get(ip)
    if not gateway:
        for candidate in raw.get("default_gateway_candidates", []):
            try:
                if ipaddress.ip_address(candidate) in network:
                    gateway = candidate
                    break
            except ValueError:
                pass
    if not gateway and raw.get("dhcp_server"):
        try:
            if ipaddress.ip_address(raw["dhcp_server"]) in network:
                gateway = raw["dhcp_server"]
        except ValueError:
            pass

    is_wifi = any(token in text for token in ["wireless", "wi-fi", "wifi", "беспровод"])
    is_vpn = any(token in text for token in ["vpn", "tailscale", "tunnel", "wireguard", "openvpn", "tap-windows"])
    is_ethernet = "ethernet" in text and not is_vpn

    score = 0
    if address.is_private:
        score += 30
    if gateway:
        score += 30
    if is_wifi:
        score += 80
    elif is_ethernet:
        score += 70
    if is_vpn:
        score -= 120
    if prefix == 24:
        score += 10

    adapter_type = "Wi-Fi" if is_wifi else "VPN" if is_vpn else "Ethernet" if is_ethernet else "Adapter"
    label = f"{adapter_type} {ip}/{prefix}"
    if description:
        label = f"{label} - {description[:45]}"

    return {
        "name": name,
        "description": description,
        "label": label,
        "adapter_type": adapter_type,
        "ip": ip,
        "mask": mask,
        "prefix": prefix,
        "network": str(network),
        "gateway": gateway or "",
        "mac": raw.get("mac", "Unknown"),
        "is_wifi": is_wifi,
        "is_ethernet": is_ethernet,
        "is_vpn": is_vpn,
        "score": score,
    }


def get_network_adapters() -> List[Dict[str, Any]]:
    return parse_ipconfig_adapters()


def get_default_network_adapter() -> Optional[Dict[str, Any]]:
    adapters = get_network_adapters()
    return adapters[0] if adapters else None


def get_default_ip_range() -> str:
    adapter = get_default_network_adapter()
    if adapter:
        return adapter["network"]
    return "192.168.0.0/24"


class ScanThread(QThread):
    progress = pyqtSignal(int, str)
    device_found = pyqtSignal(dict)
    scan_finished = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(
        self,
        ip_range: str = "192.168.0.0/24",
        max_workers: int = 10,
        adapter: Optional[Dict[str, Any]] = None,
    ):
        super().__init__()
        self.ip_range = ip_range
        self.max_workers = max_workers
        self.adapter = adapter or self.adapter_for_range(ip_range) or get_default_network_adapter()
        self.devices = []
        self._stop_requested = False
        self._arp_cache = {}
        self.discovery_methods: Dict[str, Set[str]] = {}
        self.discovery_counts: Dict[str, int] = {}
        self.plugins = self.load_plugins()

    def run(self):
        try:
            adapter_label = self.adapter.get("label", "Näbelli adapter") if self.adapter else "Näbelli adapter"
            self.progress.emit(0, f"Tor skanirlemesi başlandy: {adapter_label} | {self.ip_range}")
            hosts = self.discover_hosts()
            if not hosts:
                self.progress.emit(100, "Işjeň hostlar tapylmady.")
                self.scan_finished.emit([])
                return

            executor = ThreadPoolExecutor(max_workers=self.max_workers)
            futures = [executor.submit(self.scan_device, host) for host in hosts]
            completed = 0

            try:
                for future in as_completed(futures):
                    if self._stop_requested:
                        self.progress.emit(100, "Skanirleme saklandy.")
                        break

                    try:
                        device = future.result()
                        if device:
                            self.devices.append(device)
                            self.device_found.emit(device)
                    except Exception as exc:
                        self.error.emit(f"Skanirleme ýalňyşlygy: {exc}")

                    completed += 1
                    progress = 10 + int(completed / len(futures) * 90)
                    self.progress.emit(progress, f"Skanirlenýär... {min(progress, 100)}%")
            finally:
                if self._stop_requested:
                    for future in futures:
                        future.cancel()
                executor.shutdown(wait=True, cancel_futures=True)

            if len(self.devices) <= 1:
                self.progress.emit(
                    100,
                    "Diňe bir gurluş tapyldy. Wi-Fi aralygyny, Guest Wi-Fi/AP isolation sazlamasyny we telefonyň ekrany açykdygyny barlaň.",
                )
            self.scan_finished.emit(self.devices)

        except Exception as exc:
            self.error.emit(f"Kritiki ýalňyşlyk: {exc}")

    def stop(self):
        self._stop_requested = True

    def discover_hosts(self) -> List[str]:
        self.discovery_methods = {}
        self.discovery_counts = {"Nmap": 0, "ARP": 0, "Ping": 0, "Gateway": 0, "Local PC": 0}

        nmap_path = self.find_nmap()
        if nmap_path:
            try:
                self.add_hosts(self.discover_hosts_with_nmap(nmap_path, ["-sn", "-n"]), "Nmap")
                self.progress.emit(4, f"Nmap ping scan: {self.discovery_counts['Nmap']} host")
            except Exception as exc:
                self.progress.emit(4, f"Nmap ping scan başartmady: {exc}")

            if self.is_direct_private_network():
                try:
                    self.add_hosts(self.discover_hosts_with_nmap(nmap_path, ["-PR", "-sn", "-n"]), "ARP")
                    self.progress.emit(6, f"Nmap ARP scan: {self.discovery_counts['ARP']} host")
                except Exception as exc:
                    self.progress.emit(6, f"Nmap ARP scan başartmady: {exc}")
        else:
            self.progress.emit(4, "Nmap tapylmady, ARP/Ping gözlegi ulanylýar.")

        probe_hosts = self.hosts_from_range(self.ip_range)
        if self.should_probe_network(probe_hosts):
            self.progress.emit(7, "ARP keşini täzelemek üçin ýerli tor ping arkaly barlanýar...")
            ping_hosts = self.ping_sweep(probe_hosts)
            self.add_hosts(ping_hosts, "Ping")

        self._arp_cache = self.load_arp_cache()
        self.add_hosts(self.arp_hosts_in_range(self._arp_cache), "ARP")
        self.add_local_and_gateway()

        summary = ", ".join(f"{key}: {value}" for key, value in self.discovery_counts.items() if value)
        self.progress.emit(10, f"Işjeň hostlar: {len(self.discovery_methods)} ({summary or 'usul ýok'})")

        return self.sort_hosts(self.discovery_methods.keys())

    def add_hosts(self, hosts: Iterable[str], method: str):
        before = sum(1 for methods in self.discovery_methods.values() if method in methods)
        for host in hosts:
            if self.host_in_range(host):
                self.discovery_methods.setdefault(host, set()).add(method)
        after = sum(1 for methods in self.discovery_methods.values() if method in methods)
        self.discovery_counts[method] = after

    def add_local_and_gateway(self):
        if self.adapter and self.host_in_range(self.adapter.get("ip", "")):
            self.discovery_methods.setdefault(self.adapter["ip"], set()).add("Local PC")
            self.discovery_counts["Local PC"] = 1
        if self.adapter and self.host_in_range(self.adapter.get("gateway", "")):
            self.discovery_methods.setdefault(self.adapter["gateway"], set()).add("Gateway")
            self.discovery_counts["Gateway"] = 1

    def find_nmap(self) -> Optional[str]:
        paths = [
            Path("C:/Program Files (x86)/Nmap/nmap.exe"),
            Path("C:/Program Files/Nmap/nmap.exe"),
        ]
        for path in paths:
            if path.exists():
                return str(path)
        return shutil.which("nmap")

    def discover_hosts_with_nmap(self, nmap_path: str, arguments: List[str]) -> List[str]:
        result = run_hidden_command([nmap_path, *arguments, self.ip_range], timeout=90)
        if result.returncode not in (0, 1):
            raise RuntimeError((result.stderr or result.stdout or "nmap ýalňyşlygy").strip())

        hosts = []
        seen = set()
        for line in result.stdout.splitlines():
            if "Nmap scan report for" not in line:
                continue
            match = re.search(r"(\d{1,3}(?:\.\d{1,3}){3})", line)
            if match and match.group(1) not in seen:
                ip = match.group(1)
                seen.add(ip)
                hosts.append(ip)
        return hosts

    def hosts_from_range(self, ip_range: str) -> List[str]:
        try:
            network = ipaddress.ip_network(ip_range, strict=False)
            if network.num_addresses == 1:
                return [str(network.network_address)]
            return [str(ip) for ip in network.hosts()]
        except ValueError:
            return [ip_range.strip()] if ip_range.strip() else []

    def network_from_range(self) -> Optional[ipaddress.IPv4Network]:
        try:
            return ipaddress.ip_network(self.ip_range, strict=False)
        except ValueError:
            return None

    def host_in_range(self, host: str) -> bool:
        if not host:
            return False
        network = self.network_from_range()
        if not network:
            return False
        try:
            return ipaddress.ip_address(host) in network
        except ValueError:
            return False

    def is_direct_private_network(self) -> bool:
        network = self.network_from_range()
        return bool(network and network.is_private and (not self.adapter or not self.adapter.get("is_vpn")))

    def should_probe_network(self, hosts: List[str]) -> bool:
        network = self.network_from_range()
        if not network or not network.is_private:
            return False
        if self.adapter and self.adapter.get("is_vpn"):
            return False
        return 1 < len(hosts) <= 512

    def ping_sweep(self, hosts: List[str]) -> List[str]:
        if platform.system() == "Windows":
            command_for = lambda ip: ["ping", "-n", "1", "-w", "450", ip]
        else:
            command_for = lambda ip: ["ping", "-c", "1", "-W", "1", ip]

        responsive = []
        workers = min(64, max(8, self.max_workers * 2))
        with ThreadPoolExecutor(max_workers=workers) as executor:
            future_map = {executor.submit(run_hidden_command, command_for(host), 2): host for host in hosts}
            for future in as_completed(future_map):
                if self._stop_requested:
                    for item in future_map:
                        item.cancel()
                    break
                host = future_map[future]
                try:
                    result = future.result()
                    if result.returncode == 0:
                        responsive.append(host)
                except Exception:
                    pass
        return responsive

    def arp_hosts_in_range(self, arp_cache: Dict[str, str]) -> List[str]:
        network = self.network_from_range()
        if not network:
            return []

        hosts = []
        for ip in arp_cache:
            try:
                address = ipaddress.ip_address(ip)
            except ValueError:
                continue
            if address in network and not address.is_multicast and not str(address).endswith(".255"):
                hosts.append(ip)
        return self.sort_hosts(hosts)

    def sort_hosts(self, hosts) -> List[str]:
        return sorted(set(hosts), key=lambda value: ipaddress.ip_address(value))

    def scan_device(self, host: str) -> Optional[Dict[str, Any]]:
        if self._stop_requested:
            return None

        hostname = self.resolve_hostname(host)
        mac = self.get_mac_address(host)
        device = {
            "ip": host,
            "mac": mac,
            "vendor": vendor_from_mac(mac),
            "hostname": hostname,
            "name": hostname if hostname != "Unknown" else host,
            "ports": [],
            "services": {},
            "os_guess": "",
            "vulnerable": False,
            "vulnerabilities": [],
            "recommendation": "",
            "recommendations": [],
            "risk_level": "low",
            "discovery_method": ", ".join(sorted(self.discovery_methods.get(host, {"Unknown"}))),
            "device_type": "Unknown",
            "adapter": self.adapter.get("label", "") if self.adapter else "",
            "network": self.ip_range,
        }

        device["ports"] = self.scan_ports(host)
        if self._stop_requested:
            return None

        device["services"] = self.identify_services(device["ports"])
        device["device_type"] = self.classify_device(device)
        device = self.run_vulnerability_checks(device)
        return self.normalize_device(device)

    def scan_ports(self, host: str) -> List[int]:
        common_iot_ports = [
            21, 22, 23, 80, 443, 554, 8000, 8080, 8888,
            1883, 8883, 5683, 1900, 49152, 5353,
        ]

        open_ports = []
        for port in common_iot_ports:
            if self._stop_requested:
                break
            if self.is_port_open(host, port):
                open_ports.append(port)
        return open_ports

    def is_port_open(self, host: str, port: int, timeout: float = 0.8) -> bool:
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
                sock.settimeout(timeout)
                return sock.connect_ex((host, port)) == 0
        except OSError:
            return False

    def identify_services(self, ports: List[int]) -> Dict[int, str]:
        port_service_map = {
            21: "ftp",
            22: "ssh",
            23: "telnet",
            80: "http",
            443: "https",
            554: "rtsp",
            1883: "mqtt",
            8883: "mqtt/ssl",
            5683: "coap",
            1900: "ssdp",
            5353: "mdns",
        }
        return {port: port_service_map.get(port, "unknown") for port in ports}

    def classify_device(self, device: Dict[str, Any]) -> str:
        ip = device.get("ip")
        ports = set(device.get("ports", []))
        vendor = device.get("vendor", "")
        methods = device.get("discovery_method", "")
        hostname = (device.get("hostname") or "").lower()

        if self.adapter and ip == self.adapter.get("gateway"):
            return "Router"
        if self.adapter and ip == self.adapter.get("ip"):
            return "PC"
        if any(token in hostname for token in ["desktop", "laptop", "pc", "windows"]):
            return "PC"
        if ports & {23, 80, 443, 554, 8000, 8080, 8888, 1883, 8883, 5683, 1900}:
            return "IoT kandidat"
        if vendor in {"Apple", "Samsung", "Huawei", "Xiaomi"}:
            return "Telefon/Unknown"
        if "ARP" in methods or "Ping" in methods:
            return "Telefon/Unknown"
        return "Unknown"

    def run_vulnerability_checks(self, device: Dict[str, Any]) -> Dict[str, Any]:
        for plugin in self.plugins:
            try:
                device = plugin.check(device)
            except Exception as exc:
                self.error.emit(f"Plugin ýalňyşlygy {plugin.__name__}: {exc}")
        return device

    def normalize_device(self, device: Dict[str, Any]) -> Dict[str, Any]:
        vulnerabilities = device.get("vulnerabilities", [])
        recommendations = device.get("recommendations", [])

        device["vulnerable"] = bool(vulnerabilities) or bool(device.get("vulnerable", False))
        device["name"] = device.get("name") or device.get("hostname") or device.get("ip", "Unknown")

        if not device.get("recommendation"):
            if recommendations:
                device["recommendation"] = "; ".join(
                    rec.get("details") or rec.get("action") or str(rec)
                    for rec in recommendations
                )
            elif device["vulnerable"]:
                device["recommendation"] = "Gurluşyň sazlamalaryny barlaň we firmware-i täzeläň."
            elif not device.get("ports"):
                device["recommendation"] = "Portlar açyk däl, gurluş diňe torda ýüze çykaryldy."
            else:
                device["recommendation"] = "Kritiki maslahat ýok."

        return device

    def load_plugins(self) -> List:
        plugins = []
        plugin_dir = Path(__file__).resolve().parent.parent / "plugins"

        if not plugin_dir.exists():
            plugin_dir.mkdir(parents=True, exist_ok=True)
            return plugins

        for file_name in os.listdir(plugin_dir):
            if not file_name.endswith(".py") or file_name == "__init__.py":
                continue
            try:
                spec = importlib.util.spec_from_file_location(
                    f"plugins.{file_name[:-3]}",
                    plugin_dir / file_name,
                )
                if spec and spec.loader:
                    module = importlib.util.module_from_spec(spec)
                    spec.loader.exec_module(module)
                    if hasattr(module, "check") and callable(module.check):
                        plugins.append(module)
            except Exception as exc:
                self.error.emit(f"Plugin ýüklemek başartmady {file_name}: {exc}")

        return plugins

    def load_arp_cache(self) -> Dict[str, str]:
        command = ["arp", "-a"] if platform.system() == "Windows" else ["arp", "-n"]
        try:
            result = run_hidden_command(command, timeout=5)
        except Exception:
            return {}

        cache = {}
        ip_pattern = r"(\d{1,3}(?:\.\d{1,3}){3})"
        mac_pattern = r"([0-9a-fA-F]{2}(?:[:-][0-9a-fA-F]{2}){5})"

        for line in result.stdout.splitlines():
            ip_match = re.search(ip_pattern, line)
            mac_match = re.search(mac_pattern, line)
            if ip_match and mac_match:
                cache[ip_match.group(1)] = normalize_mac(mac_match.group(1))

        return cache

    def get_mac_address(self, ip: str) -> str:
        if self.adapter and ip == self.adapter.get("ip"):
            return self.adapter.get("mac", "Unknown")
        return self._arp_cache.get(ip, "Unknown")

    def resolve_hostname(self, ip: str) -> str:
        try:
            return socket.gethostbyaddr(ip)[0]
        except OSError:
            return "Unknown"

    def adapter_for_range(self, ip_range: str) -> Optional[Dict[str, Any]]:
        try:
            target = ipaddress.ip_network(ip_range, strict=False)
        except ValueError:
            return None
        for adapter in get_network_adapters():
            try:
                if ipaddress.ip_address(adapter["ip"]) in target:
                    return adapter
            except ValueError:
                pass
        return None

    def run_hidden_command(self, args: List[str], timeout: int) -> subprocess.CompletedProcess:
        return run_hidden_command(args, timeout)
