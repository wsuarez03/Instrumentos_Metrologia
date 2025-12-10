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

# 2. Leer la hoja CONTROL CALIBRACIONES
print("🔄 Leyendo hoja CONTROL CALIBRACIONES...")
df = pd.read_excel(
    "temp.xlsx",
    sheet_name="CONTROL CALIBRACIONES",
    dtype=str,
    header=None,
    engine="openpyxl"
)

# 3. Identificar dónde empiezan las secciones (PLANTA, VST2, VST3)
secciones = []
for idx, row in df.iterrows():
    valor = str(row[0]).strip() if pd.notna(row[0]) else ""
    if valor in ["PLANTA", "VST2", "VST3"]:
        encabezado_idx = idx + 1
        secciones.append((valor, encabezado_idx, idx))

if not secciones:
    raise Exception("❌ No se encontraron secciones PLANTA, VST2 o VST3.")

print(f"🔍 Secciones detectadas: {[s[0] for s in secciones]}")

# 4. Extraer cada sección
tablas = []
for i, (nombre, encabezado_idx, inicio_idx) in enumerate(secciones):
    # Determinar el final de la sección
    if i + 1 < len(secciones):
        fin_idx = secciones[i + 1][2]
    else:
        fin_idx = len(df)

    print(f"\n📂 Procesando sección: {nombre} (filas {inicio_idx} a {fin_idx})")
    
    # Obtener encabezados
    encabezados = df.iloc[encabezado_idx]
    
    # Extraer datos
    data = df.iloc[encabezado_idx + 1 : fin_idx].copy()
    data.columns = encabezados
    data = data.reset_index(drop=True)
    data = data.dropna(how="all")
    
    # Filtrar filas vacías en IDENTIFICACIÓN
    if "IDENTIFICACIÓN" in data.columns:
        data = data[data["IDENTIFICACIÓN"].notna() & (data["IDENTIFICACIÓN"].str.strip() != "")]
    
    if len(data) == 0:
        print(f"⚠️  No hay datos en sección {nombre}, saltando...")
        continue
    
    # Agregar origen
    data["ORIGEN"] = nombre

    # 🚀 Aquí formateamos las fechas
    columnas_fecha = [col for col in data.columns if col and ("FECHA" in str(col).upper())]
    for col in columnas_fecha:
        if col in data.columns:
            try:
                data[col] = data[col].astype(str).str[:10]
            except:
                pass
    
    # Limpiar espacios en blanco de los encabezados
    data.columns = [str(col).strip() if col else "NaN" for col in data.columns]
    
    print(f"✅ {len(data)} registros procesados")
    tablas.append(data)

# 5. Combinar todas las tablas
if not tablas:
    raise Exception("❌ No se encontraron datos en ninguna sección.")

df_final = pd.concat(tablas, ignore_index=True)

# 🧹 Limpiar y consolidar columnas de estado
# Si ESTADO\nCALIBRACIÓN está vacío, usar CALIBRACIÓN
if "CALIBRACIÓN" in df_final.columns:
    mascara = df_final["ESTADO\nCALIBRACIÓN"].isna() | (df_final["ESTADO\nCALIBRACIÓN"].str.strip() == "")
    df_final.loc[mascara, "ESTADO\nCALIBRACIÓN"] = df_final.loc[mascara, "CALIBRACIÓN"]
    df_final = df_final.drop(columns=["CALIBRACIÓN"])

# Limpiar valores "0" de fechas
for col in ["FECHA DE CALIBRACION", "FECHA PROXIMA CALIBRACIÓN"]:
    if col in df_final.columns:
        df_final.loc[df_final[col] == "0", col] = None

# 6. Mostrar resumen
print("\n" + "="*70)
print("📊 RESUMEN FINAL")
print("="*70)
print("📋 Columnas finales:", df_final.columns.tolist())
print("📈 Filas totales:", len(df_final))
print("🔍 Orígenes encontrados:", df_final["ORIGEN"].unique().tolist() if "ORIGEN" in df_final.columns else "N/A")
print("✅ Validación: Todos los registros tienen IDENTIFICACIÓN" if df_final["IDENTIFICACIÓN"].notna().all() else "⚠️  Algunos registros sin IDENTIFICACIÓN")

# 7. Guardar JSON
data_json = df_final.where(pd.notnull(df_final), None).to_dict(orient="records")
with open("instrumentos.json", "w", encoding="utf-8") as f:
    json.dump(data_json, f, ensure_ascii=False, indent=2)

print("✅ JSON creado con", len(data_json), "registros")
