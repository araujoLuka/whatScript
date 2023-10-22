from selenium import webdriver
from selenium.webdriver.common.by import By
import time

options = webdriver.ChromeOptions()
options.add_argument("user-data-dir=Perfil")

nav = webdriver.Chrome(options=options)
nav.get("https://web.whatsapp.com")

while len(nav.find_elements(By.ID, "side")) < 1:
    time.sleep(1)
time.sleep(2)

import pandas as pd
tabela = pd.read_excel("enviar.xlsx")
# print(tabela)

import urllib.parse as parse

mensagem = "Ola, tudo bem? "
for linha in tabela.index:
    if tabela.loc[linha, "enviado"] == "Ok":
        continue

    numero = tabela.loc[linha, "telefone"]
    extra = tabela.loc[linha, "extra"]

    texto = parse.quote(f"{mensagem} {extra}")

    link = f"https://web.whatsapp.com/send?phone={numero}&text={texto}"

    nav.get(link)

    while len(nav.find_elements(By.ID, "side")) < 1:
        time.sleep(1)
    time.sleep(2)

    if len(nav.find_elements(By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[1]')) < 1:
        nav.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span').click()
        tabela.loc[linha, "enviado"] = "Ok"
        print(f"Mensagem enviada para {numero}")
    else:
        tabela.loc[linha, "enviado"] = "Falhou"
        print(f"Problema ao enviar mensagem para {numero}")

    if linha == len(tabela.index) - 1:
        print("Fim da lista")
        break

    time.sleep(5)

# print(tabela)

# Save the file
df = pd.DataFrame(tabela)
df.to_excel("enviar.xlsx", index=False)
