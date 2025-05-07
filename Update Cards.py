import xlwings as xw
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager

# === Function ===
def get_price_from_cardmarket(driver, url):
    try:
        driver.get(url)
        wait = WebDriverWait(driver, 15)
        prices = wait.until(EC.presence_of_all_elements_located(
            (By.CSS_SELECTOR, "div.price-container span.text-nowrap")
        ))
        if prices:
            return prices[0].text.strip()
        else:
            return "Price non trouvé"
    except Exception as e:
        return f"Erreur : {e}"

# === Open the excel file ===
# Insert the path of your excel file
wb = xw.Book(r"C:\Users\...")
sheet = wb.sheets[0]

# === Reed excel data ===
df = sheet.range("A1").options(pd.DataFrame, header=1, index=False, expand='table').value

# === Open Edge ===
options = Options()
# options.add_argument("--headless") # You can disable the showing of the Edge window by deleting the first #
driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()), options=options)

# === Search for prices ===
price_list = []
for name, url in zip(df["Card"], df["URL"]):
    print(f"{name} → ", end="")
    price = get_price_from_cardmarket(driver, url)
    print(price)
    price_list.append(price)
    time.sleep(5)

driver.quit()

# Adding the price of cards in the excel and disable the writing of numbers in the excel
df["Price"] = price_list
sheet.range("A1").value = df.columns.tolist()  # écrit l'en-tête
sheet.range("A2").value = df.values.tolist()   # écrit les valeurs (sans index)

# Save the excel file
wb.save()
