from pathlib import Path
import math

from PIL import Image, ImageDraw, ImageFont


OUT_DIR = Path("diploma_assets/program_visuals_v2")
W, H = 1800, 1100
FONT = "C:/Windows/Fonts/segoeui.ttf"
FONT_BOLD = "C:/Windows/Fonts/seguisb.ttf"


def font(size, bold=False):
    return ImageFont.truetype(FONT_BOLD if bold else FONT, size)


C = {
    "bg": "#f8fafc",
    "ink": "#0f172a",
    "muted": "#64748b",
    "line": "#94a3b8",
    "blue": "#2563eb",
    "sky": "#0ea5e9",
    "cyan": "#06b6d4",
    "teal": "#14b8a6",
    "green": "#22c55e",
    "orange": "#f59e0b",
    "red": "#ef4444",
    "purple": "#8b5cf6",
    "card": "#ffffff",
    "soft": "#e2e8f0",
    "dark": "#1e293b",
}


def canvas(title, subtitle):
    img = Image.new("RGB", (W, H), C["bg"])
    d = ImageDraw.Draw(img)
    for x in range(0, W, 60):
        d.line((x, 0, x, H), fill="#eef2f7", width=1)
    for y in range(0, H, 60):
        d.line((0, y, W, y), fill="#eef2f7", width=1)
    d.rounded_rectangle((55, 40, W - 55, H - 40), 30, outline="#e2e8f0", width=2)
    d.text((90, 70), title, font=font(46, True), fill=C["ink"])
    d.text((92, 130), subtitle, font=font(23), fill=C["muted"])
    return img, d


def rounded(d, box, fill, outline, width=3, radius=28):
    d.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def text_center(d, box, text, size=28, color=None, bold=False, gap=7):
    color = color or C["ink"]
    f = font(size, bold)
    lines = text.split("\n")
    dims = [d.textbbox((0, 0), line, font=f) for line in lines]
    widths = [b[2] - b[0] for b in dims]
    heights = [b[3] - b[1] for b in dims]
    total = sum(heights) + gap * (len(lines) - 1)
    x1, y1, x2, y2 = box
    y = y1 + (y2 - y1 - total) / 2
    for line, w, h in zip(lines, widths, heights):
        d.text((x1 + (x2 - x1 - w) / 2, y), line, font=f, fill=color)
        y += h + gap


def arrow(d, a, b, color="#334155", width=5, dashed=False):
    if dashed:
        x1, y1 = a
        x2, y2 = b
        length = math.hypot(x2 - x1, y2 - y1)
        steps = max(1, int(length / 22))
        for i in range(steps):
            if i % 2 == 0:
                s = i / steps
                e = min((i + 1) / steps, 1)
                d.line((x1 + (x2 - x1) * s, y1 + (y2 - y1) * s, x1 + (x2 - x1) * e, y1 + (y2 - y1) * e), fill=color, width=width)
    else:
        d.line((*a, *b), fill=color, width=width)
    ang = math.atan2(b[1] - a[1], b[0] - a[0])
    ln, wing = 21, math.pi / 7
    p1 = (b[0] - ln * math.cos(ang - wing), b[1] - ln * math.sin(ang - wing))
    p2 = (b[0] - ln * math.cos(ang + wing), b[1] - ln * math.sin(ang + wing))
    d.polygon([b, p1, p2], fill=color)


def monitor(d, cx, cy, s=1, color=C["blue"]):
    d.rounded_rectangle((cx - 70*s, cy - 45*s, cx + 70*s, cy + 45*s), 12, fill="#eff6ff", outline=color, width=int(5*s))
    d.rectangle((cx - 12*s, cy + 45*s, cx + 12*s, cy + 70*s), fill=color)
    d.rounded_rectangle((cx - 48*s, cy + 68*s, cx + 48*s, cy + 78*s), 4, fill=color)


def db(d, cx, cy, s=1, color=C["purple"]):
    d.ellipse((cx - 62*s, cy - 75*s, cx + 62*s, cy - 35*s), fill="#f5f3ff", outline=color, width=int(5*s))
    d.rectangle((cx - 62*s, cy - 55*s, cx + 62*s, cy + 55*s), fill="#f5f3ff", outline=color, width=int(5*s))
    d.ellipse((cx - 62*s, cy + 35*s, cx + 62*s, cy + 75*s), fill="#f5f3ff", outline=color, width=int(5*s))
    for yy in [-18, 20]:
        d.arc((cx - 62*s, cy + yy*s, cx + 62*s, cy + (yy+40)*s), 0, 180, fill=color, width=int(4*s))


def router(d, cx, cy, s=1, color=C["teal"]):
    d.rounded_rectangle((cx - 82*s, cy - 35*s, cx + 82*s, cy + 35*s), 18, fill="#ecfeff", outline=color, width=int(5*s))
    for dx in [-42, 0, 42]:
        d.ellipse((cx+dx*s-7*s, cy-6*s, cx+dx*s+7*s, cy+8*s), fill=color)
    d.arc((cx - 90*s, cy - 115*s, cx + 90*s, cy + 65*s), 218, 322, fill=color, width=int(5*s))
    d.arc((cx - 130*s, cy - 155*s, cx + 130*s, cy + 105*s), 218, 322, fill=color, width=int(4*s))


def phone(d, cx, cy, s=1, color=C["sky"]):
    d.rounded_rectangle((cx - 42*s, cy - 78*s, cx + 42*s, cy + 78*s), 18, fill="#eff6ff", outline=color, width=int(5*s))
    d.rounded_rectangle((cx - 16*s, cy - 66*s, cx + 16*s, cy - 58*s), 4, fill=color)
    d.ellipse((cx - 6*s, cy + 55*s, cx + 6*s, cy + 67*s), fill=color)


def camera(d, cx, cy, s=1, color=C["orange"]):
    d.rounded_rectangle((cx - 78*s, cy - 42*s, cx + 38*s, cy + 42*s), 15, fill="#fff7ed", outline=color, width=int(5*s))
    d.polygon([(cx+38*s, cy-25*s), (cx+92*s, cy-55*s), (cx+92*s, cy+55*s), (cx+38*s, cy+25*s)], fill="#fff7ed", outline=color)
    d.ellipse((cx - 40*s, cy - 28*s, cx + 16*s, cy + 28*s), outline=color, width=int(5*s))


def shield(d, cx, cy, s=1, color=C["green"]):
    pts = [(cx, cy-70*s), (cx+62*s, cy-38*s), (cx+48*s, cy+45*s), (cx, cy+78*s), (cx-48*s, cy+45*s), (cx-62*s, cy-38*s)]
    d.polygon(pts, fill="#f0fdf4", outline=color)
    d.line(pts + [pts[0]], fill=color, width=int(5*s))
    d.line((cx-28*s, cy, cx-7*s, cy+25*s, cx+34*s, cy-28*s), fill=color, width=int(7*s))


def gear(d, cx, cy, s=1, color="#64748b"):
    for i in range(10):
        a = math.tau * i / 10
        x, y = cx + math.cos(a)*58*s, cy + math.sin(a)*58*s
        d.rounded_rectangle((x-12*s, y-12*s, x+12*s, y+12*s), 5, fill=color)
    d.ellipse((cx-55*s, cy-55*s, cx+55*s, cy+55*s), fill="#f8fafc", outline=color, width=int(6*s))
    d.ellipse((cx-21*s, cy-21*s, cx+21*s, cy+21*s), fill="#e2e8f0", outline=color, width=int(4*s))


def chip(d, x, y, text, color):
    rounded(d, (x, y, x+170, y+58), "#ffffff", color, 3, 18)
    text_center(d, (x, y, x+170, y+58), text, 22, color, True)


def caption(d, text):
    d.text((95, 1000), text, font=font(27, True), fill=C["muted"])


def draw_mini_ui(d, box):
    x1, y1, x2, y2 = box
    rounded(d, box, "#0f172a", "#334155", 4, 26)
    d.rectangle((x1, y1, x1+150, y2), fill="#1e293b")
    for i, label in enumerate(["Baş", "Scan", "Panel", "Netije", "Taryh"]):
        yy = y1 + 55 + i*54
        fill = "#06b6d4" if label == "Scan" else "#334155"
        rounded(d, (x1+18, yy, x1+132, yy+36), fill, fill, 1, 10)
    rounded(d, (x1+180, y1+45, x2-25, y1+120), "#1e293b", "#475569", 2, 14)
    for i in range(5):
        yy = y1 + 160 + i*43
        d.line((x1+195, yy, x2-45, yy), fill="#64748b", width=3)
    d.rectangle((x1+185, y2-90, x2-35, y2-42), fill="#111827", outline="#475569")


def save(img, filename):
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    path = OUT_DIR / filename
    img.save(path, quality=95)
    return path


def fig1_architecture():
    img, d = canvas("Programma arhitekturasy", "Interfeýs, skanirleme ýadrosy, pluginler, baza we hasabat bir ulgamda işleýär")
    # layer backgrounds
    rounded(d, (95, 245, 1695, 425), "#eef6ff", "#bfdbfe", 2, 30)
    rounded(d, (95, 465, 1695, 645), "#ecfdf5", "#99f6e4", 2, 30)
    rounded(d, (95, 685, 1695, 865), "#fff7ed", "#fed7aa", 2, 30)
    d.text((125, 260), "Ulanyjy gatlagy", font=font(24, True), fill=C["blue"])
    d.text((125, 480), "Logika gatlagy", font=font(24, True), fill=C["teal"])
    d.text((125, 700), "Netije gatlagy", font=font(24, True), fill=C["orange"])
    boxes = [
        ("PyQt6\nUI", (260, 295, 520, 395), C["blue"], monitor),
        ("ScanThread\nCore", (700, 500, 1000, 610), C["teal"], gear),
        ("Nmap\nARP/Ping", (1140, 500, 1400, 610), C["cyan"], None),
        ("Plugins", (1140, 295, 1400, 395), C["orange"], None),
        ("SQLite\nTaryh", (520, 735, 780, 835), C["purple"], db),
        ("PDF\nReportLab", (1020, 735, 1280, 835), C["green"], None),
    ]
    for label, b, color, icon in boxes:
        rounded(d, b, "#ffffff", color, 4, 24)
        if icon:
            icon(d, b[0]+55, (b[1]+b[3])//2, 0.43, color)
            text_center(d, (b[0]+105, b[1], b[2], b[3]), label, 24, bold=True)
        else:
            text_center(d, b, label, 26, bold=True)
    arrow(d, (520, 345), (700, 555), C["blue"], 4)
    arrow(d, (1000, 555), (1140, 555), C["teal"], 4)
    arrow(d, (1000, 520), (1140, 355), C["orange"], 4)
    arrow(d, (850, 610), (650, 735), C["purple"], 4)
    arrow(d, (900, 610), (1150, 735), C["green"], 4)
    caption(d, "Surat 1 - Programma arhitekturasynyň gatlaklaýyn görnüşi")
    return save(img, "01_programma_arhitekturasy_v2.png")


def fig2_scan_flow():
    img, d = canvas("Skanirleme algoritmi", "Host tapmak, port barlamak we maslahat döretmek yzygiderli ýerine ýetirilýär")
    steps = [("Adapter", C["blue"]), ("IP", C["sky"]), ("Nmap", C["cyan"]), ("ARP/Ping", C["teal"]), ("Port", C["orange"]), ("Risk", C["red"]), ("PDF", C["green"])]
    x0, y = 115, 440
    for i, (txt, col) in enumerate(steps):
        x = x0 + i*235
        d.ellipse((x, y, x+130, y+130), fill="#ffffff", outline=col, width=5)
        text_center(d, (x, y, x+130, y+130), txt, 25, col, True)
        if i < len(steps)-1:
            arrow(d, (x+130, y+65), (x+225, y+65), C["dark"], 4)
    # mini timeline
    rounded(d, (230, 675, 1560, 835), "#ffffff", "#cbd5e1", 3, 28)
    for i, label in enumerate(["Wi‑Fi saýlanýar", "192.168.1.0/24", "Gurluşlar", "Portlar", "Maslahat"]):
        xx = 290 + i*250
        chip(d, xx, 725, label, [C["blue"], C["sky"], C["teal"], C["orange"], C["green"]][i])
    caption(d, "Surat 2 - Skanirleme algoritminiň giňeldilen wizual akymy")
    return save(img, "02_skanirleme_algoritmi_v2.png")


def fig3_network():
    img, d = canvas("Lokal toruň kartasy", "Programma lokal Wi‑Fi segmentinde PC, router, telefon we IoT kandidatlary görkezýär")
    draw_mini_ui(d, (90, 300, 470, 735))
    router(d, 900, 535, 1.25, C["teal"])
    text_center(d, (760, 635, 1040, 710), "Wi‑Fi Router\n192.168.1.1", 25, bold=True)
    arrow(d, (470, 520), (730, 520), C["blue"], 5)
    devices = [
        ("PC", "192.168.1.3", 720, 245, monitor, C["purple"]),
        ("Telefon", "ARP/Ping", 1320, 245, phone, C["sky"]),
        ("IP kamera", "80 / 554", 1440, 560, camera, C["orange"]),
        ("IoT datçik", "MQTT", 1180, 825, shield, C["green"]),
    ]
    for name, sub, cx, cy, ic, col in devices:
        d.ellipse((cx-115, cy-115, cx+115, cy+115), fill="#ffffff", outline=col, width=5)
        ic(d, cx, cy-18, 0.62, col)
        text_center(d, (cx-110, cy+45, cx+110, cy+98), name + "\n" + sub, 21, bold=True)
        arrow(d, (980, 540), (cx, cy), "#64748b", 4)
    for i, m in enumerate(["Nmap", "ARP", "Ping", "Gateway"]):
        chip(d, 555 + i*170, 790, m, [C["cyan"], C["teal"], C["sky"], C["purple"]][i])
    caption(d, "Surat 3 - Lokal torda gurluşlaryň wizual kartalaşdyrylyşy")
    return save(img, "03_lokal_tor_kartasy_v2.png")


def fig4_tech():
    img, d = canvas("Ulanylan tehnologiýalar", "Python programma ýadrosy bolup, beýleki tehnologiýalar onuň töwereginde birleşýär")
    d.ellipse((725, 355, 1075, 705), fill="#eff6ff", outline=C["blue"], width=7)
    text_center(d, (725, 355, 1075, 705), "Python\nprogrammirleme\ndili", 36, C["blue"], True)
    tech = [
        ("PyQt6", "interfeýs", 350, 250, C["teal"], monitor),
        ("Nmap", "skanirleme", 1450, 250, C["orange"], gear),
        ("SQLite", "baza", 290, 700, C["purple"], db),
        ("ReportLab", "PDF", 1510, 700, C["green"], None),
        ("Paramiko", "SSH", 900, 850, C["red"], shield),
        ("PyInstaller", "EXE", 900, 190, C["cyan"], None),
    ]
    for name, sub, cx, cy, col, ic in tech:
        rounded(d, (cx-170, cy-86, cx+170, cy+86), "#ffffff", col, 4, 30)
        if ic:
            ic(d, cx-105, cy, 0.42, col)
            text_center(d, (cx-50, cy-70, cx+155, cy+70), name+"\n"+sub, 25, col, True)
        else:
            text_center(d, (cx-160, cy-75, cx+160, cy+75), name+"\n"+sub, 27, col, True)
        arrow(d, (900, 530), (cx, cy), "#94a3b8", 3, dashed=True)
    caption(d, "Surat 4 - Taslamada ulanylan tehnologiýalaryň wizual toplumy")
    return save(img, "04_tehnologiyalar_v2.png")


def fig5_data_flow():
    img, d = canvas("Maglumatlaryň hereketi", "Tapylan gurluşlar bazada saklanyp, birnäçe sahypada gaýtadan ulanylýar")
    rounded(d, (100, 360, 420, 705), "#ffffff", C["teal"], 5, 30)
    text_center(d, (100, 380, 420, 490), "ScanThread", 31, C["teal"], True)
    for i, row in enumerate(["IP", "MAC", "Port", "Risk"]):
        chip(d, 170, 520+i*48, row, C["teal"])
    db(d, 735, 535, 1.25, C["purple"])
    text_center(d, (610, 660, 860, 720), "iot_security.db", 28, C["purple"], True)
    arrow(d, (420, 535), (620, 535), C["teal"], 5)
    pages = [("Panel", 1220, 300, C["blue"]), ("Netijeler", 1410, 535, C["orange"]), ("Taryh", 1220, 770, C["purple"]), ("PDF", 980, 770, C["green"])]
    for name, cx, cy, col in pages:
        rounded(d, (cx-145, cy-75, cx+145, cy+75), "#ffffff", col, 4, 25)
        text_center(d, (cx-140, cy-70, cx+140, cy+70), name, 30, col, True)
        arrow(d, (845, 535), (cx-145 if cx > 845 else cx, cy), "#64748b", 4)
    caption(d, "Surat 5 - Netijeleriň maglumat bazasyndan sahypalara paýlanyşy")
    return save(img, "05_maglumat_akymy_v2.png")


def fig6_risk():
    img, d = canvas("Töwekgelçilik modeli", "Portlar we pluginler esasynda programma maslahat döredýär")
    # left device cards
    devs = [("Port ýok", C["green"], phone), ("HTTP/RTSP", C["orange"], camera), ("Telnet", C["red"], router)]
    for i, (label, col, ic) in enumerate(devs):
        y = 300 + i*205
        rounded(d, (130, y, 450, y+145), "#ffffff", col, 4, 24)
        ic(d, 210, y+72, 0.45, col)
        text_center(d, (270, y+15, 430, y+130), label, 26, col, True)
        arrow(d, (450, y+72), (680, 535), col, 4)
    # risk engine
    d.ellipse((680, 365, 1080, 705), fill="#ffffff", outline=C["purple"], width=6)
    text_center(d, (700, 390, 1060, 680), "Risk\nbahalandyrmak\nPlugins", 34, C["purple"], True)
    # outputs
    outs = [("Pes", C["green"], 1300, 310), ("Orta", C["orange"], 1440, 535), ("Ýokary", C["red"], 1300, 760)]
    for name, col, cx, cy in outs:
        d.ellipse((cx-90, cy-90, cx+90, cy+90), fill="#ffffff", outline=col, width=5)
        text_center(d, (cx-85, cy-85, cx+85, cy+85), name, 29, col, True)
        arrow(d, (1080, 535), (cx-90, cy), col, 4)
    rounded(d, (595, 825, 1235, 920), "#f8fafc", "#64748b", 3, 22)
    text_center(d, (610, 825, 1220, 920), "Maslahat: parol, firmware, port, howpsuz protokol", 26, bold=True)
    caption(d, "Surat 6 - Töwekgelçiligiň pluginler arkaly bahalandyrylyşy")
    return save(img, "06_risk_modeli_v2.png")


def main():
    paths = [fig1_architecture(), fig2_scan_flow(), fig3_network(), fig4_tech(), fig5_data_flow(), fig6_risk()]
    captions = [
        "Surat 1 - Programma arhitekturasynyň gatlaklaýyn görnüşi.",
        "Surat 2 - Skanirleme algoritminiň giňeldilen wizual akymy.",
        "Surat 3 - Lokal torda gurluşlaryň wizual kartalaşdyrylyşy.",
        "Surat 4 - Taslamada ulanylan tehnologiýalaryň wizual toplumy.",
        "Surat 5 - Netijeleriň maglumat bazasyndan sahypalara paýlanyşy.",
        "Surat 6 - Töwekgelçiligiň pluginler arkaly bahalandyrylyşy.",
    ]
    (OUT_DIR / "captions_turkmen.txt").write_text("\n".join(captions), encoding="utf-8")
    for p in paths:
        print(p.resolve())
    print((OUT_DIR / "captions_turkmen.txt").resolve())


if __name__ == "__main__":
    main()
