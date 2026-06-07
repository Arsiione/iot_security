from pathlib import Path
import math

from PIL import Image, ImageDraw, ImageFont


OUT_DIR = Path("diploma_assets/program_visuals")
W, H = 1800, 1100

FONT = "C:/Windows/Fonts/segoeui.ttf"
FONT_BOLD = "C:/Windows/Fonts/seguisb.ttf"


def font(size, bold=False):
    path = FONT_BOLD if bold else FONT
    return ImageFont.truetype(path, size)


COLORS = {
    "bg": "#f7fafc",
    "ink": "#0f172a",
    "muted": "#64748b",
    "line": "#94a3b8",
    "teal": "#14b8a6",
    "blue": "#2563eb",
    "cyan": "#06b6d4",
    "green": "#22c55e",
    "orange": "#f59e0b",
    "red": "#ef4444",
    "purple": "#8b5cf6",
    "card": "#ffffff",
    "soft": "#e2e8f0",
}


def new_canvas(title, subtitle=None):
    img = Image.new("RGB", (W, H), COLORS["bg"])
    d = ImageDraw.Draw(img)
    # subtle background grid
    for x in range(0, W, 80):
        d.line((x, 0, x, H), fill="#eef2f7", width=1)
    for y in range(0, H, 80):
        d.line((0, y, W, y), fill="#eef2f7", width=1)

    d.text((90, 64), title, font=font(44, True), fill=COLORS["ink"])
    if subtitle:
        d.text((92, 122), subtitle, font=font(22), fill=COLORS["muted"])
    return img, d


def rounded(d, box, fill, outline="#cbd5e1", width=2, radius=28):
    d.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def centered_text(d, box, text, size=28, fill=None, bold=False, line_gap=8):
    fill = fill or COLORS["ink"]
    x1, y1, x2, y2 = box
    lines = text.split("\n")
    f = font(size, bold)
    heights = []
    widths = []
    for line in lines:
        b = d.textbbox((0, 0), line, font=f)
        widths.append(b[2] - b[0])
        heights.append(b[3] - b[1])
    total_h = sum(heights) + line_gap * (len(lines) - 1)
    y = y1 + ((y2 - y1) - total_h) / 2
    for line, w, h in zip(lines, widths, heights):
        x = x1 + ((x2 - x1) - w) / 2
        d.text((x, y), line, font=f, fill=fill)
        y += h + line_gap


def draw_arrow(d, start, end, color="#334155", width=5):
    x1, y1 = start
    x2, y2 = end
    d.line((x1, y1, x2, y2), fill=color, width=width)
    angle = math.atan2(y2 - y1, x2 - x1)
    length = 20
    wing = math.pi / 7
    p1 = (x2 - length * math.cos(angle - wing), y2 - length * math.sin(angle - wing))
    p2 = (x2 - length * math.cos(angle + wing), y2 - length * math.sin(angle + wing))
    d.polygon([(x2, y2), p1, p2], fill=color)


def icon_monitor(d, cx, cy, scale=1, color="#2563eb"):
    w, h = 110 * scale, 70 * scale
    d.rounded_rectangle((cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2), 10, outline=color, width=int(5 * scale), fill="#eff6ff")
    d.line((cx - 22 * scale, cy + h / 2 + 12 * scale, cx + 22 * scale, cy + h / 2 + 12 * scale), fill=color, width=int(5 * scale))
    d.line((cx, cy + h / 2, cx, cy + h / 2 + 13 * scale), fill=color, width=int(5 * scale))


def icon_database(d, cx, cy, scale=1, color="#8b5cf6"):
    w, h = 100 * scale, 115 * scale
    d.ellipse((cx - w / 2, cy - h / 2, cx + w / 2, cy - h / 2 + 28 * scale), outline=color, width=int(4 * scale), fill="#f5f3ff")
    d.rectangle((cx - w / 2, cy - h / 2 + 14 * scale, cx + w / 2, cy + h / 2 - 14 * scale), outline=color, width=int(4 * scale), fill="#f5f3ff")
    d.ellipse((cx - w / 2, cy + h / 2 - 28 * scale, cx + w / 2, cy + h / 2), outline=color, width=int(4 * scale), fill="#f5f3ff")
    d.arc((cx - w / 2, cy - h / 2 + 14 * scale, cx + w / 2, cy - h / 2 + 42 * scale), 0, 180, fill=color, width=int(4 * scale))


def icon_shield(d, cx, cy, scale=1, color="#22c55e"):
    pts = [
        (cx, cy - 62 * scale),
        (cx + 58 * scale, cy - 35 * scale),
        (cx + 44 * scale, cy + 42 * scale),
        (cx, cy + 72 * scale),
        (cx - 44 * scale, cy + 42 * scale),
        (cx - 58 * scale, cy - 35 * scale),
    ]
    d.polygon(pts, fill="#f0fdf4", outline=color)
    d.line(pts + [pts[0]], fill=color, width=int(5 * scale), joint="curve")
    d.line((cx - 24 * scale, cy, cx - 5 * scale, cy + 22 * scale, cx + 30 * scale, cy - 22 * scale), fill=color, width=int(7 * scale))


def icon_router(d, cx, cy, scale=1, color="#14b8a6"):
    d.rounded_rectangle((cx - 70 * scale, cy - 28 * scale, cx + 70 * scale, cy + 28 * scale), 16, fill="#ecfeff", outline=color, width=int(5 * scale))
    for dx in [-35, 0, 35]:
        d.ellipse((cx + dx * scale - 6 * scale, cy - 5 * scale, cx + dx * scale + 6 * scale, cy + 7 * scale), fill=color)
    d.arc((cx - 78 * scale, cy - 98 * scale, cx + 78 * scale, cy + 58 * scale), 220, 320, fill=color, width=int(4 * scale))
    d.arc((cx - 116 * scale, cy - 136 * scale, cx + 116 * scale, cy + 96 * scale), 220, 320, fill=color, width=int(4 * scale))


def icon_phone(d, cx, cy, scale=1, color="#0ea5e9"):
    d.rounded_rectangle((cx - 36 * scale, cy - 68 * scale, cx + 36 * scale, cy + 68 * scale), 14, fill="#eff6ff", outline=color, width=int(5 * scale))
    d.ellipse((cx - 5 * scale, cy + 48 * scale, cx + 5 * scale, cy + 58 * scale), fill=color)


def icon_camera(d, cx, cy, scale=1, color="#f59e0b"):
    d.rounded_rectangle((cx - 62 * scale, cy - 36 * scale, cx + 34 * scale, cy + 36 * scale), 14, fill="#fff7ed", outline=color, width=int(5 * scale))
    d.polygon([(cx + 34 * scale, cy - 22 * scale), (cx + 78 * scale, cy - 44 * scale), (cx + 78 * scale, cy + 44 * scale), (cx + 34 * scale, cy + 22 * scale)], fill="#fff7ed", outline=color)
    d.ellipse((cx - 28 * scale, cy - 22 * scale, cx + 16 * scale, cy + 22 * scale), outline=color, width=int(5 * scale))


def icon_gear(d, cx, cy, scale=1, color="#64748b"):
    for i in range(8):
        angle = math.pi * 2 * i / 8
        x = cx + math.cos(angle) * 50 * scale
        y = cy + math.sin(angle) * 50 * scale
        d.ellipse((x - 10 * scale, y - 10 * scale, x + 10 * scale, y + 10 * scale), fill=color)
    d.ellipse((cx - 45 * scale, cy - 45 * scale, cx + 45 * scale, cy + 45 * scale), fill="#f8fafc", outline=color, width=int(5 * scale))
    d.ellipse((cx - 18 * scale, cy - 18 * scale, cx + 18 * scale, cy + 18 * scale), fill="#e2e8f0", outline=color, width=int(4 * scale))


def save(img, name):
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    path = OUT_DIR / name
    img.save(path, quality=95)
    return path


def diagram_architecture():
    img, d = new_canvas("Programma arhitekturasy", "UI, skanirleme ýadrosy, pluginler, maglumat bazasy we hasabat")
    # nodes
    nodes = {
        "UI\nPyQt6": (120, 390, 430, 610, "#dbeafe", COLORS["blue"], icon_monitor),
        "Scanner\nCore": (745, 360, 1055, 640, "#ecfeff", COLORS["teal"], icon_gear),
        "Nmap\nARP / Ping": (740, 170, 1060, 300, "#f0f9ff", COLORS["cyan"], None),
        "Plugins": (1220, 260, 1530, 430, "#fff7ed", COLORS["orange"], None),
        "SQLite": (1220, 520, 1530, 700, "#f5f3ff", COLORS["purple"], icon_database),
        "PDF\nReportLab": (740, 760, 1060, 920, "#f0fdf4", COLORS["green"], None),
    }
    for label, (x1, y1, x2, y2, fill, outline, ic) in nodes.items():
        rounded(d, (x1, y1, x2, y2), fill, outline, 4)
        if ic:
            ic(d, (x1 + x2) / 2, y1 + 70, 0.75, outline)
            centered_text(d, (x1, y1 + 120, x2, y2), label, 31, bold=True)
        else:
            centered_text(d, (x1, y1, x2, y2), label, 34, bold=True)
    draw_arrow(d, (430, 500), (745, 500), COLORS["blue"])
    draw_arrow(d, (900, 360), (900, 300), COLORS["cyan"])
    draw_arrow(d, (1055, 445), (1220, 350), COLORS["orange"])
    draw_arrow(d, (1055, 565), (1220, 610), COLORS["purple"])
    draw_arrow(d, (900, 640), (900, 760), COLORS["green"])
    d.text((120, 960), "Surat 1. Programma bölekleriniň özara baglanyşygy", font=font(26, True), fill=COLORS["muted"])
    return save(img, "01_programma_arhitekturasy.png")


def diagram_scan_flow():
    img, d = new_canvas("Skanirleme algoritmi", "Adapterden hasabata çenli esasy iş akymy")
    steps = [
        ("Adapter", COLORS["blue"]),
        ("IP aralyk", COLORS["cyan"]),
        ("Host tapmak", COLORS["teal"]),
        ("Port barlagy", COLORS["orange"]),
        ("Risk", COLORS["red"]),
        ("Hasabat", COLORS["green"]),
    ]
    y = 500
    w_box, h_box, gap = 220, 150, 55
    x = 120
    for i, (label, color) in enumerate(steps):
        box = (x, y - h_box / 2, x + w_box, y + h_box / 2)
        rounded(d, box, "#ffffff", color, 4, 24)
        centered_text(d, box, label, 31, bold=True)
        if i < len(steps) - 1:
            draw_arrow(d, (x + w_box, y), (x + w_box + gap - 10, y), "#334155", 4)
        x += w_box + gap
    # discovery detail bubbles
    bubbles = [("Nmap", 570, 730), ("ARP", 740, 785), ("Ping", 910, 730), ("Gateway", 1080, 785)]
    for text, bx, by in bubbles:
        d.ellipse((bx - 72, by - 42, bx + 72, by + 42), fill="#f8fafc", outline="#94a3b8", width=3)
        centered_text(d, (bx - 72, by - 42, bx + 72, by + 42), text, 23, bold=True)
        draw_arrow(d, (bx, by - 42), (780, 575), "#94a3b8", 3)
    d.text((120, 960), "Surat 2. Programma skanirleme tapgyrlarynyň yzygiderliligi", font=font(26, True), fill=COLORS["muted"])
    return save(img, "02_skanirleme_algoritmi.png")


def diagram_network():
    img, d = new_canvas("Lokal toruň skanirlenişi", "Programma Wi‑Fi torundaky gurluşlary birnäçe usul bilen tapýar")
    # central router and scanner
    rounded(d, (110, 430, 450, 690), "#ffffff", COLORS["blue"], 4)
    icon_monitor(d, 280, 505, 1.05, COLORS["blue"])
    centered_text(d, (110, 570, 450, 670), "IoT Security\nScanner", 30, bold=True)
    icon_router(d, 900, 540, 1.15, COLORS["teal"])
    centered_text(d, (760, 620, 1040, 700), "Wi‑Fi Router", 28, bold=True)
    draw_arrow(d, (450, 550), (760, 550), COLORS["blue"], 5)

    devices = [
        ("Telefon", 1310, 260, icon_phone, COLORS["cyan"]),
        ("IP kamera", 1470, 520, icon_camera, COLORS["orange"]),
        ("IoT datçik", 1285, 805, icon_shield, COLORS["green"]),
        ("PC", 720, 835, icon_monitor, COLORS["purple"]),
    ]
    for label, cx, cy, ic, color in devices:
        d.ellipse((cx - 105, cy - 105, cx + 105, cy + 105), fill="#ffffff", outline=color, width=4)
        ic(d, cx, cy - 10, 0.72, color)
        centered_text(d, (cx - 105, cy + 52, cx + 105, cy + 105), label, 23, bold=True)
        draw_arrow(d, (980, 560), (cx - 105 if cx > 980 else cx + 105, cy), "#64748b", 4)

    # method chips
    for i, text in enumerate(["Nmap", "ARP", "Ping", "MAC"]):
        x1 = 520 + i * 150
        rounded(d, (x1, 300, x1 + 120, 360), "#eff6ff", "#60a5fa", 2, 18)
        centered_text(d, (x1, 300, x1 + 120, 360), text, 22, bold=True)
    d.text((120, 960), "Surat 3. Lokal Wi‑Fi torunda gurluşlaryň ýüze çykarylyşy", font=font(26, True), fill=COLORS["muted"])
    return save(img, "03_lokal_tor_skanirlemesi.png")


def diagram_technologies():
    img, d = new_canvas("Ulanylan tehnologiýalar", "Programma üpjünçiligini döretmekde ulanylan esasy gurallar")
    # center Python
    d.ellipse((720, 330, 1080, 690), fill="#eff6ff", outline=COLORS["blue"], width=6)
    centered_text(d, (720, 330, 1080, 690), "Python", 48, bold=True, fill=COLORS["blue"])
    items = [
        ("PyQt6\nUI", 350, 285, COLORS["teal"]),
        ("Nmap\nScan", 1350, 285, COLORS["orange"]),
        ("SQLite\nDB", 350, 755, COLORS["purple"]),
        ("ReportLab\nPDF", 1350, 755, COLORS["green"]),
        ("Paramiko\nSSH", 900, 840, COLORS["red"]),
        ("PyInstaller\nEXE", 900, 180, COLORS["cyan"]),
    ]
    for label, cx, cy, color in items:
        rounded(d, (cx - 160, cy - 80, cx + 160, cy + 80), "#ffffff", color, 4, 28)
        centered_text(d, (cx - 150, cy - 72, cx + 150, cy + 72), label, 29, bold=True)
        draw_arrow(d, (900, 510), (cx, cy), "#94a3b8", 3)
    d.text((120, 960), "Surat 4. Programma toplumynda ulanylan tehnologiýalar toplumy", font=font(26, True), fill=COLORS["muted"])
    return save(img, "04_ulanylan_tehnologiyalar.png")


def diagram_data_flow():
    img, d = new_canvas("Maglumat akymy", "Skanirleme netijeleriniň saklanyşy we görkezilişi")
    rounded(d, (120, 410, 450, 650), "#ffffff", COLORS["teal"], 4)
    centered_text(d, (120, 410, 450, 650), "Skanirleme\nnetijeleri", 33, bold=True)
    icon_database(d, 770, 530, 1.1, COLORS["purple"])
    centered_text(d, (610, 650, 930, 720), "SQLite", 32, bold=True, fill=COLORS["purple"])
    draw_arrow(d, (450, 530), (650, 530), COLORS["teal"], 5)
    outputs = [
        ("Panel", 1250, 285, COLORS["blue"]),
        ("Netijeler", 1450, 530, COLORS["orange"]),
        ("Taryh", 1250, 780, COLORS["purple"]),
        ("PDF", 1020, 780, COLORS["green"]),
    ]
    for label, cx, cy, color in outputs:
        rounded(d, (cx - 140, cy - 75, cx + 140, cy + 75), "#ffffff", color, 4, 24)
        centered_text(d, (cx - 135, cy - 70, cx + 135, cy + 70), label, 31, bold=True)
        draw_arrow(d, (880, 530), (cx - 140 if cx > 880 else cx, cy), "#64748b", 4)
    d.text((120, 960), "Surat 5. Skanirleme netijeleriniň maglumat bazasyndan sahypalara geçişi", font=font(26, True), fill=COLORS["muted"])
    return save(img, "05_maglumat_akymy.png")


def diagram_risk():
    img, d = new_canvas("Töwekgelçilik we maslahat", "Tapylan gurluşlar risk derejesine görä seljerilýär")
    cards = [
        ("Pes", "Port ýok", COLORS["green"], icon_phone),
        ("Orta", "Web / RTSP", COLORS["orange"], icon_camera),
        ("Ýokary", "Telnet", COLORS["red"], icon_router),
    ]
    xs = [250, 700, 1150]
    for (title, detail, color, ic), x in zip(cards, xs):
        rounded(d, (x, 330, x + 330, 650), "#ffffff", color, 5, 30)
        ic(d, x + 165, 425, 0.72, color)
        centered_text(d, (x, 500, x + 330, 575), title, 40, bold=True, fill=color)
        centered_text(d, (x, 575, x + 330, 635), detail, 25, bold=True)
        draw_arrow(d, (x + 165, 650), (x + 165, 770), color, 4)
    rounded(d, (390, 770, 1410, 900), "#f8fafc", "#64748b", 4, 28)
    icon_shield(d, 510, 835, 0.55, COLORS["teal"])
    centered_text(d, (600, 780, 1370, 890), "Maslahat: parol, port, firmware, howpsuz protokol", 30, bold=True)
    d.text((120, 960), "Surat 6. Gurluşlaryň töwekgelçilik derejesi we howpsuzlyk maslahaty", font=font(26, True), fill=COLORS["muted"])
    return save(img, "06_towekgelcilik_maslahat.png")


def main():
    paths = [
        diagram_architecture(),
        diagram_scan_flow(),
        diagram_network(),
        diagram_technologies(),
        diagram_data_flow(),
        diagram_risk(),
    ]
    captions = [
        "Surat 1 - Programma arhitekturasy: UI, Scanner Core, Plugins, SQLite we PDF hasabatlarynyň baglanyşygy.",
        "Surat 2 - Skanirleme algoritmi: adapter saýlamakdan hasabat taýýarlamaga çenli iş akymy.",
        "Surat 3 - Lokal toruň skanirlenişi: programma router, telefon, PC we IoT gurluşlaryny tapýar.",
        "Surat 4 - Ulanylan tehnologiýalar: Python, PyQt6, Nmap, SQLite, ReportLab, Paramiko we PyInstaller.",
        "Surat 5 - Maglumat akymy: skanirleme netijeleri SQLite bazasyndan Panel, Netijeler, Taryh we PDF hasabata geçýär.",
        "Surat 6 - Töwekgelçilik we maslahat: pes, orta we ýokary risk derejeleri boýunça howpsuzlyk maslahatlary.",
    ]
    (OUT_DIR / "captions_turkmen.txt").write_text("\n".join(captions), encoding="utf-8")
    print("Created images:")
    for path in paths:
        print(path.resolve())
    print((OUT_DIR / "captions_turkmen.txt").resolve())


if __name__ == "__main__":
    main()
