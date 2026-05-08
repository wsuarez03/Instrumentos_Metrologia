import pandas as pd
import requests
import json

# =========================================================
# 1. Descargar el Excel
# =========================================================

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

# Validar que realmente sea un Excel
if (
    resp.status_code != 200
    or len(resp.content) < 10000
    or start.lstrip().startswith("<!DOCTYPE html")
):
    raise Exception("❌ No se descargó un Excel válido. Revisa el enlace o permisos")

# Guardar archivo temporal
with open("temp.xlsx", "wb") as f:
    f.write(resp.content)

print("✅ Archivo guardado: temp.xlsx")

# =========================================================
# 2. Leer hoja CONTROL CALIBRACIONES
# =========================================================

print("🔄 Leyendo hoja CONTROL CALIBRACIONES...")

df = pd.read_excel(
    "temp.xlsx",
    sheet_name="CONTROL CALIBRACIONES",
    dtype=str,
    header=None,
    engine="openpyxl"
)

print(f"✅ Hoja cargada: {len(df)} filas")

# =========================================================
# 3. Detectar secciones dinámicamente
# =========================================================

secciones = []

for idx, row in df.iterrows():

    valor = str(row[0]).strip() if pd.notna(row[0]) else ""

    # Detectar nombres de secciones
    if valor in ["PLANTA", "VST2", "VST3"]:

        print(f"\n🔍 Sección encontrada: {valor} en fila {idx}")

        encabezado_idx = None

        # Buscar encabezado real automáticamente
        for j in range(idx + 1, min(idx + 10, len(df))):

            fila = df.iloc[j].astype(str)

            if fila.str.contains(
                "IDENTIFICACIÓN",
                case=False,
                na=False
            ).any():

                encabezado_idx = j

                print(f"✅ Encabezado encontrado en fila {j}")

                break

        # Si no encuentra encabezado, saltar sección
        if encabezado_idx is None:
            print(f"⚠️ No se encontró encabezado para {valor}")
            continue

        # Guardar sección
        secciones.append((valor, encabezado_idx, idx))

# Validar secciones encontradas
if not secciones:
    raise Exception("❌ No se encontraron secciones válidas.")

print("\n📂 Secciones detectadas:")
for s in secciones:
    print(f" - {s[0]}")

# =========================================================
# 4. Extraer datos de cada sección
# =========================================================

tablas = []

for i, (nombre, encabezado_idx, inicio_idx) in enumerate(secciones):

    # Determinar final de sección
    if i + 1 < len(secciones):
        fin_idx = secciones[i + 1][2]
    else:
        fin_idx = len(df)

    print("\n" + "=" * 60)
    print(f"📂 Procesando sección: {nombre}")
    print(f"📍 Filas: {inicio_idx} -> {fin_idx}")

    # Obtener encabezados
    encabezados = df.iloc[encabezado_idx]

    # Extraer datos
    data = df.iloc[encabezado_idx + 1 : fin_idx].copy()

    # Asignar encabezados
    data.columns = encabezados

    # Resetear índice
    data = data.reset_index(drop=True)

    # Eliminar filas completamente vacías
    data = data.dropna(how="all")

    # Limpiar nombres de columnas
    data.columns = [
        str(col).strip() if pd.notna(col) else "SIN_NOMBRE"
        for col in data.columns
    ]

    # Filtrar registros vacíos
    if "IDENTIFICACIÓN" in data.columns:

        data = data[
            data["IDENTIFICACIÓN"].notna()
            & (data["IDENTIFICACIÓN"].astype(str).str.strip() != "")
        ]

    else:
        print(f"⚠️ La sección {nombre} no tiene IDENTIFICACIÓN")
        continue

    # Si no hay datos válidos
    if len(data) == 0:
        print(f"⚠️ No hay registros válidos en {nombre}")
        continue

    # Agregar origen
    data["ORIGEN"] = nombre

    # =====================================================
    # Formatear columnas fecha
    # =====================================================

    columnas_fecha = [
        col
        for col in data.columns
        if "FECHA" in str(col).upper()
    ]

    for col in columnas_fecha:

        try:
            data[col] = data[col].astype(str).str[:10]

            # Limpiar valores inválidos
            data.loc[data[col] == "0", col] = None
            data.loc[data[col] == "nan", col] = None

        except Exception as e:
            print(f"⚠️ Error procesando fecha {col}: {e}")

    print(f"✅ Registros procesados: {len(data)}")

    tablas.append(data)

# =========================================================
# 5. Combinar tablas
# =========================================================

if not tablas:
    raise Exception("❌ No se encontraron datos válidos.")

df_final = pd.concat(tablas, ignore_index=True)

# =========================================================
# 6. Consolidar columnas de calibración
# =========================================================

if (
    "ESTADO\nCALIBRACIÓN" in df_final.columns
    and "CALIBRACIÓN" in df_final.columns
):

    mascara = (
        df_final["ESTADO\nCALIBRACIÓN"].isna()
        | (
            df_final["ESTADO\nCALIBRACIÓN"]
            .astype(str)
            .str.strip()
            == ""
        )
    )

    df_final.loc[
        mascara,
        "ESTADO\nCALIBRACIÓN"
    ] = df_final.loc[mascara, "CALIBRACIÓN"]

    df_final = df_final.drop(columns=["CALIBRACIÓN"])

# =========================================================
# 7. Mostrar resumen
# =========================================================

print("\n" + "=" * 70)
print("📊 RESUMEN FINAL")
print("=" * 70)

print("📋 Columnas finales:")
for c in df_final.columns:
    print(" -", c)

print("\n📈 Total registros:", len(df_final))

if "ORIGEN" in df_final.columns:
    print("📂 Orígenes:", df_final["ORIGEN"].unique().tolist())

if "IDENTIFICACIÓN" in df_final.columns:

    completos = df_final["IDENTIFICACIÓN"].notna().sum()

    print(f"✅ Registros con IDENTIFICACIÓN: {completos}")

# =========================================================
# 8. Guardar JSON
# =========================================================

data_json = df_final.where(pd.notnull(df_final), None).to_dict(
    orient="records"
)

with open("instrumentos.json", "w", encoding="utf-8") as f:
    json.dump(data_json, f, ensure_ascii=False, indent=2)

print("\n✅ JSON creado correctamente")
print(f"📦 Total registros exportados: {len(data_json)}")
print("💾 Archivo generado: instrumentos.json")
