from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from matplotlib.figure import Figure
from pathlib import Path
import tempfile

def get_report_font():
    font_paths = [
        Path("C:/Windows/Fonts/arial.ttf"),
        Path("C:/Windows/Fonts/segoeui.ttf"),
    ]

    for font_path in font_paths:
        if font_path.exists():
            pdfmetrics.registerFont(TTFont("IoTReportFont", str(font_path)))
            return "IoTReportFont"

    return "Helvetica"

def generate_pdf(devices, filename="report.pdf"):
    c = canvas.Canvas(filename, pagesize=A4)
    font_name = get_report_font()
    c.setFont(font_name, 12)
    y = 800
    c.drawString(50, y, "IoT Howpsuzlyk Hasabaty")
    y -= 30

    for device in devices:
        if y < 90:
            c.showPage()
            c.setFont(font_name, 12)
            y = 800

        name = device.get("name") or device.get("hostname") or "Näbelli"
        ip = device.get("ip", "Näbelli")
        if name == "Unknown":
            name = "Näbelli"
        if ip == "Unknown":
            ip = "Näbelli"
        vulnerable = bool(device.get("vulnerable") or device.get("vulnerabilities"))
        recommendation = device.get("recommendation") or "Kritiki maslahat ýok."
        ports = ", ".join(map(str, device.get("ports", []))) or "ýok"
        vendor = device.get("vendor") or "Näbelli"
        device_type = device.get("device_type") or "Unknown"
        discovery_method = device.get("discovery_method") or "Unknown"

        c.drawString(50, y, f"Gurluş: {name} ({ip})")
        y -= 20
        c.drawString(50, y, f"Görnüşi: {device_type} | Öndüriji: {vendor} | Tapylan usul: {discovery_method}")
        y -= 20
        c.drawString(50, y, f"Portlar: {ports}")
        y -= 20
        c.drawString(50, y, f"Gowşaklyk: {'Hawa' if vulnerable else 'Ýok'}")
        y -= 20
        c.drawString(50, y, f"Maslahat: {recommendation[:110]}")
        y -= 30

    if devices:
        fig = Figure(figsize=(3, 2))
        ax = fig.add_subplot(111)
        safe = sum(1 for d in devices if not (d.get("vulnerable") or d.get("vulnerabilities")))
        vuln = len(devices) - safe
        ax.pie([safe, vuln], labels=["Safe", "Vulnerable"], colors=["#66FF66", "#FF6666"], autopct="%1.1f%%")

        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            chart_path = tmp.name

        try:
            fig.savefig(chart_path, bbox_inches="tight")
            if y < 260:
                c.showPage()
                c.setFont(font_name, 12)
                y = 800
            c.drawImage(chart_path, 50, y - 200, width=200, height=150)
        finally:
            Path(chart_path).unlink(missing_ok=True)

    c.save()
