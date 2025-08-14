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

# 📂 Crear carpeta de salida si no existe
Path("TEST_salida").mkdir(exist_ok=True)

# 📂 Cargar archivo generado por Script 1
fecha_hoy = datetime.today().strftime('%Y-%m-%d')
archivo_salida = Path("TEST_salida") / f"resultado_observaciones_{fecha_hoy}.xlsx"

if not archivo_salida.exists():
    raise FileNotFoundError(f"❌ No se encontró el archivo: {archivo_salida}")

df = pd.read_excel(archivo_salida)
COLUMNA_DNIS = "DNI"
dnis = df[COLUMNA_DNIS].astype(str).tolist()

# 🔹 Configuración para Chrome en modo headless
def crear_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    return webdriver.Chrome(options=chrome_options)

# 🔍 Función para buscar nombre por DNI en SUNAT usando driver ya abierto
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

    except Exception as e:
        print(f"⚠️ Error con DNI {dni}: {e}")
    finally:
        return resultado

# ▶️ Procesar todos los DNIs con un solo driver
driver = crear_driver()
resultados = []
for dni in dnis:
    resultados.append(buscar_nombre(driver, dni))
    time.sleep(random.uniform(2, 5))  # pausa aleatoria entre 2 y 5 segundos

driver.quit()

# 📌 Unir OBS_DNI y nombre al DataFrame original
df_resultados = pd.DataFrame(resultados)
df = df.merge(
    df_resultados[["dni", "nombre", "OBS_DNI"]],
    left_on=COLUMNA_DNIS,
    right_on="dni",
    how="left"
).drop(columns=["dni"])

# 💾 Guardar archivo actualizado
df.to_excel(archivo_salida, index=False)
print(f"📁 Archivo final actualizado con OBS_DNI: {archivo_salida}")

