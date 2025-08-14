import time
import random
from datetime import datetime
from pathlib import Path
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import tempfile

# 📂 Carpeta de depuración
DEBUG_FOLDER = Path("DEBUG_FOLDER")
DEBUG_FOLDER.mkdir(exist_ok=True)

# 📂 Carpeta donde está el archivo de entrada
Path("TEST_salida").mkdir(exist_ok=True)

# 📂 Ruta del archivo generado por Script 1
fecha_hoy = datetime.today().strftime('%Y-%m-%d')
archivo_salida = Path("TEST_salida") / f"resultado_observaciones_{fecha_hoy}.xlsx"

# ⚠️ Verificar existencia antes de leer
if not archivo_salida.exists():
    raise FileNotFoundError(f"❌ No se encontró el archivo: {archivo_salida}")

# 📄 Leer archivo Excel
df_dnis = pd.read_excel(archivo_salida)

COLUMNA_DNIS = "DNI"
if COLUMNA_DNIS not in df_dnis.columns:
    raise ValueError(f"❌ No se encontró la columna '{COLUMNA_DNIS}' en el archivo.")

# Convertir DNIs a texto y limpiar espacios
df_dnis[COLUMNA_DNIS] = df_dnis[COLUMNA_DNIS].astype(str).str.strip()
dnis = df_dnis[COLUMNA_DNIS].tolist()

# 🔹 Configuración para Chrome en modo headless
chrome_options = Options()
chrome_options.add_argument("--headless")  # Modo sin interfaz
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

# Crear un directorio temporal para el perfil de Chrome
user_data_dir = tempfile.mkdtemp()
chrome_options.add_argument(f"--user-data-dir={user_data_dir}")

driver = webdriver.Chrome(options=chrome_options)

# 🔍 Función para buscar nombre por DNI en SUNAT con guardado de depuración
def buscar_nombre(driver, dni):
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

        nombre_element = driver.find_element(By.XPATH, "//h4[2]")
        nombre = nombre_element.text.strip()

        resultado["nombre"] = nombre
        resultado["OBS_DNI"] = "✅ OK"
        print(f"✅ DNI {dni}: {nombre}")

        # Guardar HTML y screenshot exitosos
        with open(DEBUG_FOLDER / f"{dni}_ok.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        driver.save_screenshot(DEBUG_FOLDER / f"{dni}_ok.png")

    except Exception as e:
        print(f"⚠️ Error con DNI {dni}: {e}")
        driver.save_screenshot(DEBUG_FOLDER / f"{dni}_error.png")
        with open(DEBUG_FOLDER / f"{dni}_error.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)

    return resultado

# ▶️ Procesar todos los DNIs con pausas aleatorias
resultados = []
for dni in dnis:
    resultados.append(buscar_nombre(driver, dni))
    time.sleep(random.uniform(2, 5))  # Pausa aleatoria para evitar bloqueos

# Cerrar driver al final
driver.quit()

# 📌 Crear DataFrame de resultados y forzar tipo string en 'dni'
df_resultados = pd.DataFrame(resultados)
df_resultados["dni"] = df_resultados["dni"].astype(str)

# 📌 Unir resultados al DataFrame original
df_final = df_dnis.merge(
    df_resultados,
    left_on=COLUMNA_DNIS,
    right_on="dni",
    how="left"
).drop(columns=["dni"])

# 💾 Guardar archivo actualizado
df_final.to_excel(archivo_salida, index=False)
print(f"📁 Archivo final actualizado con OBS_DNI: {archivo_salida}")
print(f"📂 Archivos de depuración guardados en: {DEBUG_FOLDER.resolve()}")
