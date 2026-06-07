# plugins/telnet_weak_auth.py
import logging
import telnetlib
from typing import Any, Dict

logger = logging.getLogger(__name__)


def check(device: Dict[str, Any]) -> Dict[str, Any]:
    """Telnet hyzmatynda gowşak login/parol jübütlerini barlaýar."""
    if 23 not in device.get("ports", []):
        return device

    weak_credentials = [
        ("admin", "admin"),
        ("root", "root"),
        ("user", "user"),
        ("admin", "password"),
        ("root", "123456"),
        ("admin", "1234"),
    ]

    for username, password in weak_credentials:
        if test_telnet_login(device["ip"], username, password):
            vulnerability = {
                "type": "telnet_weak_auth",
                "severity": "high",
                "description": f"Telnet gowşak autentifikasiýasy: {username}/{password}",
                "fix_available": True,
                "fix_method": "change_password",
            }

            device.setdefault("vulnerabilities", []).append(vulnerability)
            device["vulnerable"] = True
            device["risk_level"] = update_risk_level(device.get("risk_level", "low"), "high")

            recommendation = {
                "action": "Telnet parolyny üýtgetmek",
                "priority": "high",
                "details": f"Gowşak parol tapyldy: {username}/{password}",
            }
            device.setdefault("recommendations", []).append(recommendation)
            break

    return device


def test_telnet_login(ip: str, username: str, password: str, timeout: int = 2) -> bool:
    """Login/parol kabul edilende hakyky shell nyşanynyň çykandygyny barlaýar."""
    tn = None
    try:
        tn = telnetlib.Telnet(ip, 23, timeout=timeout)

        index, _, _ = tn.expect(
            [b"login:", b"username:", b"Login:", b"Username:"],
            timeout=timeout,
        )
        if index < 0:
            return False
        tn.write(username.encode("ascii") + b"\n")

        index, _, _ = tn.expect(
            [b"password:", b"Password:"],
            timeout=timeout,
        )
        if index < 0:
            return False
        tn.write(password.encode("ascii") + b"\n")

        index, _, text = tn.expect(
            [b"incorrect", b"failed", b"denied", b"Login incorrect", b"Welcome", b"#", b"\\$", b">"],
            timeout=timeout,
        )
        response = text.lower()
        if index in (0, 1, 2, 3) or b"incorrect" in response or b"failed" in response:
            return False

        return index >= 4

    except Exception as exc:
        logger.debug("Telnet login failed for %s: %s", ip, exc)
        return False
    finally:
        if tn is not None:
            try:
                tn.close()
            except Exception:
                pass


def update_risk_level(current_level: str, new_level: str) -> str:
    """Töwekgelçilik derejesini ýokarlandyrýar."""
    levels = {"low": 0, "medium": 1, "high": 2, "critical": 3}
    return max([current_level, new_level], key=lambda value: levels.get(value, 0))
