import json
import qrcode
import os

# === CONFIGURACIÓN ===
BASE_URL = "https://wsuarez03.github.io/instrumento/instrumento.html?id="
JSON_FILE = "instrumentos.json"  # Debes haberlo generado previamente
OUTPUT_DIR = "qrs"

# Crear carpeta si no existe
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Cargar los datos del JSON
with open(JSON_FILE, encoding="utf-8") as f:
    data = json.load(f)

# Generar QR por cada instrumento
for section, instrumentos in data.items():
    for instrumento in instrumentos:
        id_inst = instrumento.get("IDENTIFICACIÓN")
        if not id_inst:
            continue
        url = f"{BASE_URL}{id_inst}"
        qr = qrcode.make(url)
        qr.save(os.path.join(OUTPUT_DIR, f"{id_inst}.png"))

print("✅ QRs generados en la carpeta 'qrs'")
