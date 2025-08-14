import time
import random
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==============================
# 1. CONFIGURAR NAVEGADOR
# ==============================
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36")

driver = webdriver.Chrome(options=chrome_options)

# ==============================
# 2. LEER EXCEL DE DNIs
# ==============================
df_dnis = pd.read_excel("entrada.xlsx")  # Columna: DNI
df_dnis["DNI"] = df_dnis["DNI"].astype(str).str.strip()

# Lista para guardar resultados
resultados = []

# ==============================
# 3. HACER CONSULTAS CON UN SOLO DRIVER
# ==============================
driver.get("https://e-consulta.sunat.gob.pe/cl-ti-itmrconsruc/jcrS00Alias")

for dni in df_dnis["DNI"]:
    try:
        # Esperar campo de ingreso
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "txtRuc"))).clear()
        
        # Escribir DNI
        driver.find_element(By.ID, "txtRuc").send_keys(dni)
        driver.find_element(By.ID, "txtRuc").send_keys(Keys.RETURN)

        # Esperar que aparezca el resultado (ajusta selector si cambia)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "someElementResult")))
        
        # Extraer nombre (AJUSTAR SELECTOR AL REAL)
        nombre = driver.find_element(By.ID, "nombreContribuyente").text.strip()

        resultados.append({"DNI": dni, "NOMBRE": nombre})
        print(f"✅ {dni} → {nombre}")

    except Exception as e:
        print(f"⚠️ Error con DNI {dni}: {e}")
        resultados.append({"DNI": dni, "NOMBRE": None})

    # Pausa aleatoria para no ser bloqueados
    time.sleep(random.uniform(2.5, 6.0))

# ==============================
# 4. CERRAR NAVEGADOR
# ==============================
driver.quit()

# ==============================
# 5. UNIFICAR Y GUARDAR RESULTADOS
# ==============================
df_resultados = pd.DataFrame(resultados)
df_final = pd.merge(df_dnis, df_resultados, on="DNI", how="left")

df_final.to_excel("resultado_final.xlsx", index=False)
print("📁 Resultados guardados en resultado_final.xlsx")
