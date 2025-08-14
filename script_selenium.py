import pandas as pd
import time
import random
import tempfile
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# === CONFIGURAR SELENIUM PARA GITHUB ACTIONS ===
chrome_options = Options()
chrome_options.add_argument("--headless")  # Sin interfaz
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

# Crear un perfil temporal único para esta ejecución
user_data_dir = tempfile.mkdtemp()
chrome_options.add_argument(f"--user-data-dir={user_data_dir}")

driver = webdriver.Chrome(options=chrome_options)

# === LEER LISTA DE DNIs ===
df = pd.read_excel("dnis.xlsx")  # tu archivo con columna DOCUMENTO
resultados = []

# === ABRIR UNA SOLA SESIÓN ===
driver.get("https://www2.sunat.gob.pe/")

for dni in df["DOCUMENTO"]:
    try:
        # Aquí iría el flujo para buscar el DNI en SUNAT
        # (Esto depende de la estructura de la web, ajusta selectores según corresponda)

        # Ejemplo: escribir DNI
        campo = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "txtDni"))
        )
        campo.clear()
        campo.send_keys(str(dni))

        # Clic en botón buscar
        boton = driver.find_element(By.ID, "btnBuscar")
        boton.click()

        # Esperar resultado
        nombre = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".resultadoNombre"))
        ).text

        resultados.append({"DNI": dni, "Nombre": nombre})

        # Pausa aleatoria para evitar bloqueo
        time.sleep(random.uniform(3, 6))

    except Exception as e:
        resultados.append({"DNI": dni, "Nombre": f"Error: {e}"})

# === GUARDAR RESULTADOS ===
df_resultados = pd.DataFrame(resultados)
df_resultados.to_excel("resultado_dnis.xlsx", index=False)

driver.quit()
print("✅ Proceso completado y resultados guardados en resultado_dnis.xlsx")
