import json
import qrcode
from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm

# Configuración
BASE_URL = "https://wsuarez03.github.io/instrumento/instrumento.html?id="
JSON_FILE = "instrumentos.json"
PDF_FILE = "etiquetas_qr.pdf"

# Tamaño de hoja y márgenes
PAGE_WIDTH, PAGE_HEIGHT = A4
QR_SIZE = 35 * mm
MARGIN_X = 15 * mm
MARGIN_Y = 15 * mm
COLS = 3
ROWS = 8
PADDING = 5 * mm

# Crear canvas
c = canvas.Canvas(PDF_FILE, pagesize=A4)

# Cargar JSON
with open(JSON_FILE, encoding="utf-8") as f:
    data = json.load(f)

x, y = 0, 0
label_count = 0

for section, instrumentos in data.items():
    for inst in instrumentos:
        id_ = inst.get("IDENTIFICACIÓN")
        if not id_:
            continue

        url = f"{BASE_URL}{id_}"
        qr_img = qrcode.make(url)
        qr_path = f"temp_qr.png"
        qr_img.save(qr_path)

        col = label_count % COLS
        row = label_count // COLS % ROWS

        pos_x = MARGIN_X + col * ((PAGE_WIDTH - 2 * MARGIN_X) / COLS)
        pos_y = PAGE_HEIGHT - MARGIN_Y - row * ((PAGE_HEIGHT - 2 * MARGIN_Y) / ROWS)

        # Dibujar QR
        c.drawImage(qr_path, pos_x, pos_y - QR_SIZE, QR_SIZE, QR_SIZE)

        # Texto debajo del QR
        c.setFont("Helvetica", 6)
        text_y = pos_y - QR_SIZE - 2
        c.drawString(pos_x, text_y - 8, f"ID: {id_}")
        c.drawString(pos_x, text_y - 16, f"Marca: {inst.get('FABRICANTE', '')}")
        c.drawString(pos_x, text_y - 24, f"Modelo: {inst.get('MODELO', '')}")
        c.drawString(pos_x, text_y - 32, f"Calibrado: {inst.get('FECHA DE CALIBRACION', '')[:10]}")

        label_count += 1
        if label_count % (COLS * ROWS) == 0:
            c.showPage()

# Guardar PDF
c.save()
print("✅ PDF de etiquetas generado:", PDF_FILE)
