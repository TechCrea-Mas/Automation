import time
from datetime import datetime
from pathlib import Path
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- Limpieza inicial de columnas ---
def limpiar_columnas(df):
    df.columns = df.columns.str.strip()  # Quita espacios al inicio/fin
    df.columns = df.columns.str.replace('\n', ' ', regex=True)  # Quita saltos de línea
    return df

# --- Asegurar tipo texto en columna clave ---
def forzar_texto(df, columna):
    if columna in df.columns:
        df[columna] = df[columna].astype(str).str.strip()
    return df

# 📂 Crear carpeta de salida si no existe
Path("TEST_salida").mkdir(exist_ok=True)

# 📂 Cargar archivo generado por Script 1
fecha_hoy = datetime.today().strftime('%Y-%m-%d')
archivo_salida = Path("TEST_salida") / f"resultado_observaciones_{fecha_hoy}.xlsx"

if not archivo_salida.exists():
    raise FileNotFoundError(f"❌ No se encontró el archivo: {archivo_salida}")

# Cargar y limpiar el DataFrame
df = pd.read_excel(archivo_salida)
df = limpiar_columnas(df)

# 📌 Columna que contiene los DNIs (ya limpia)
COLUMNA_DNIS = "Documento de identidad (DNI/Pasaporte/Cédula):\n"

def limpiar_columna_dnis(df, COLUMNA_DNIS):
    df[COLUMNA_DNIS] = df[COLUMNA_DNIS].astype(str).apply(lambda x: x.strip() if isinstance(x, str) else x)
    return df

# Convertir columna a lista
dnis = df[COLUMNA_DNIS].tolist()
print(f"📄 Archivo cargado con {len(dnis)} DNIs.")

# 🔹 Configuración para Chrome en modo headless
def crear_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    return webdriver.Chrome(options=chrome_options)

# 🔍 Función para buscar nombre por DNI en SUNAT
def buscar_nombre(dni):
    driver = crear_driver()
    resultado = {"dni": dni, "nombre": None, "OBS_DNI": "❌ ERROR"}
    try:
        driver.get("https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/FrameCriterioBusquedaWeb.jsp")

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "btnPorDocumento"))
        ).click()

        input_dni = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "txtNumeroDocumento"))
        )
        input_dni.clear()
        input_dni.send_keys(dni)

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "btnAceptar"))
        ).click()

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "list-group-item-heading"))
        )

        time.sleep(1)
        nombre_element = driver.find_element(By.XPATH, "//h4[2]")
        nombre = nombre_element.text.strip()

        resultado["nombre"] = nombre
        resultado["OBS_DNI"] = "✅ OK"
        print(f"✅ DNI {dni}: {nombre}")

    except Exception as e:
        print(f"⚠️ Error con DNI {dni}: {e}")
    finally:
        driver.quit()
        return resultado

# ▶️ Procesar todos los DNIs
resultados = [buscar_nombre(dni) for dni in dnis]
df_resultados = pd.DataFrame(resultados)

# 📌 Unir OBS_DNI y nombre al DataFrame original
df = df.merge(
    df_resultados[["dni", "nombre", "OBS_DNI"]],
    left_on=COLUMNA_DNIS,
    right_on="dni",
    how="left"
).drop(columns=["dni"])

# 💾 Guardar archivo actualizado
df.to_excel(archivo_salida, index=False)
print(f"📁 Archivo final actualizado con OBS_DNI: {archivo_salida}")
