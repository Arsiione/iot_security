def check(device):
    if 80 in device.get("ports", []):
        vulnerability = {
            "type": "possible_hikvision_http",
            "severity": "medium",
            "description": "HTTP interfeýsi açyk. Kameranyň firmware-i we hasaplary barlaň.",
            "fix_available": False,
            "fix_method": "manual_review"
        }
        recommendation = {
            "action": "Gurluşyň web interfeýsini barlamak",
            "priority": "medium",
            "details": "Firmware-i täzeläň, standart hasaplary öçüriň we HTTP elýeterliligini daşarky tordan ýapyň."
        }

        device.setdefault("vulnerabilities", []).append(vulnerability)
        device.setdefault("recommendations", []).append(recommendation)
        device["risk_level"] = "medium"
    return device
