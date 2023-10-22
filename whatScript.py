from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pandas as pd
import urllib.parse as parse

def pause(text=""):
    if text == "":
        text = "Pressione Enter para continuar..."
    print(text)
    input()

contadorMensagensEnviadas = 0

print("Abrindo planilha do Excel...")
fileName = "C:\\Users\\Gerencia\\Desktop\\Bot Whatsapp\\whatScript\\enviar.xlsx"
try:
    tabela = pd.read_excel(io=fileName, dtype={"ENVIADOS": str})
except Exception as e:
    print(e)
    pause()
    exit
    
# print(tabela)

if (tabela is None):
    print("Tabela vazia")
    exit
    
print("Tabela encontrada - Número de registros:", len(tabela))

# pause()


print("Abrindo Chrome com uso de cookies...")
options = webdriver.ChromeOptions()
options.add_argument("user-data-dir=\.\\PerfilW\\")

nav = webdriver.Chrome(options=options)
nav.get("https://web.whatsapp.com")


print("Aguardando carregamento do Whatsapp Web...")
while len(nav.find_elements(By.ID, "side")) < 1:
    time.sleep(1)
time.sleep(2)

print("Enviando mensagens...")

mensagem = "Olá, tudo bem?\n\
Sentiu saudades? ❤️\n\
Nós sentimos! 🥰\n\
E por a saudade ser grande, temos um presente para você!\n\
Desconto de 50% nos purificadores EUROPA.\n\
E de 15% na manutenção.\n\
\n\
Vamos conversar?"

for linha in tabela.index:
    # print(linha)
    # print(tabela.loc[linha])
    # pause()
    numero = tabela.loc[linha, "PHONE"]
    
    if tabela.loc[linha, "ENVIADOS"] == "Ok":
        print("Numero", numero, "ja recebeu uma mensagem!")
        continue
    
    print("Enviado mensagem para numero:", numero)
    
    nome = tabela.loc[linha, "NAME"]

    texto = parse.quote(f"{mensagem}")

    link = f"https://web.whatsapp.com/send?phone={numero}&text={texto}"

    nav.get(link)

    while len(nav.find_elements(By.ID, "side")) < 1:
        time.sleep(1)
    time.sleep(2)

    if len(nav.find_elements(By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[1]')) < 1:
        nav.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span').click()
        tabela.loc[linha, "ENVIADOS"] = str("Ok")
        print(f"Mensagem enviada para {numero}")
        contadorMensagensEnviadas += 1
    else:
        tabela.loc[linha, "ENVIADOS"] = str("Falhou")
        print(f"Problema ao enviar mensagem para {numero}")

    if linha == len(tabela.index)-1:
        print("Fim da lista")
        break

    time.sleep(5)

if (contadorMensagensEnviadas == 0):
    print("Sem mensagens para enviar! Atualize a planilha com novos numeros")
    pause("Pressione Enter para finalizar...")
    exit()

print("Todas as mensagens foram enviadas!")
print("- Total de", contadorMensagensEnviadas, "mensagens enviadas")

# Save the file
print("Atualizando planilha...")
df = pd.DataFrame(tabela)
df.to_excel(fileName, index=False)

pause("Pressione Enter para finalizar...")