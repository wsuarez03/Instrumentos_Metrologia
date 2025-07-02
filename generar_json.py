import pandas as pd
import requests
import json

# 1. Descargar el Excel
url = (
    "https://valserindustriales-my.sharepoint.com"
    "/personal/sst_valserindustriales_com/_layouts/15/download.aspx"
    "?share=EX92mI4ZUiRKgyLGkriSWP4BFF5E4yCIuMbIQif16dm9Ug"
)

print("🔄 Descargando Excel...")
resp = requests.get(url)

print("📏 Tamaño descargado:", len(resp.content))
start = resp.content[:200].decode(errors="ignore")
print("🔍 Inicio del contenido:", start[:100].replace("\n", ""))

if resp.status_code != 200 or len(resp.content) < 10_000 or start.lstrip().startswith("<!DOCTYPE html"):
    raise Exception("❌ No se descargó un Excel válido. Revisa el enlace o permisos")

with open("temp.xlsx", "wb") as f:
    f.write(resp.content)
print("✅ Archivo guardado: temp.xlsx")

# 2. Leer todo el Excel sin encabezados
print("🔄 Leyendo Excel...")
df = pd.read_excel(
    "temp.xlsx",
    dtype=str,
    header=None,
    engine="openpyxl"
)

# 3. Identificar dónde empiezan los bloques
bloques = []
for idx, row in df.iterrows():
    if str(row[0]).strip() in ["PLANTA", "VST1", "VST2"]:
        nombre = str(row[0]).strip()
        encabezado_idx = idx + 1
        bloques.append((nombre, encabezado_idx))

if not bloques:
    raise Exception("❌ No se encontraron bloques con etiquetas PLANTA, VST1 o VST2.")

print("🔍 Bloques detectados:", [b[0] for b in bloques])

# 4. Extraer cada bloque
tablas = []
for i, (nombre, encabezado_idx) in enumerate(bloques):
    if i + 1 < len(bloques):
        fin_idx = bloques[i + 1][1] - 2
    else:
        fin_idx = len(df)

    data = df.iloc[encabezado_idx + 1 : fin_idx].copy()
    data.columns = df.iloc[encabezado_idx]
    data = data.reset_index(drop=True)
    data = data.dropna(how="all")
    data["IDENTIFICACIÓN"] = data["IDENTIFICACIÓN"].fillna(method="ffill")
    data["ORIGEN"] = nombre

    # 🚀 Aquí formateamos las fechas
    for col in ["FECHA DE CALIBRACION", "FECHA PROXIMA CALIBRACIÓN", "NaN"]:
        if col in data.columns:
            data[col] = data[col].str[:10]

    tablas.append(data)

# 5. Combinar todas las tablas
df_final = pd.concat(tablas, ignore_index=True)

# 6. Mostrar resumen
print("📋 Columnas finales:", df_final.columns.tolist())
print("📈 Filas totales:", len(df_final))

# 7. Guardar JSON
data_json = df_final.where(pd.notnull(df_final), None).to_dict(orient="records")
with open("instrumentos.json", "w", encoding="utf-8") as f:
    json.dump(data_json, f, ensure_ascii=False, indent=2)

print("✅ JSON creado con", len(data_json), "registros")