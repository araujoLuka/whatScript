from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pandas as pd
import urllib.parse as parse
import tkinter
from tkinter import filedialog
import os
from openpyxl import load_workbook
from datetime import date, datetime

# Const values
dataPath = r'C:\\Program Files (x86)\\whatScript\\data\\'
limiteMensagens = 200
introducao = "Programa WhatScript \n" + \
"- Autor: Lucas Araujo \n" + \
"- Versao: 1.3 \n" + \
"\n" + \
"> Envio de mensagens automáticas via WhatsApp \n" + \
"- Limite: " + str(limiteMensagens) + "\n"
mensagem = "Olá, tudo bem?\n\
Sentiu saudades? ❤️\n\
Nós sentimos! 🥰\n\
E por a saudade ser grande, temos um presente para você!\n\
Desconto de 50% nos purificadores EUROPA.\n\
E de 15% na manutenção.\n\
\n\
Vamos conversar?"
contPath=dataPath + "cont.txt"
logPath=dataPath + "wscript.log"

def pause(text: str = "") -> None:
    if text == "":
        text = "Pressione Enter para continuar..."
    print(text)
    input()

def defineContador(fPath: str = contPath) -> int:
    if not os.path.exists(fPath):
        print("ATENÇÃO: Sem informação sobre o número de " + \
                "mensagens já enviadas hoje!")
        print("ATENÇÃO: Iniciando contador em zero(0)")
        return 0
    
    with open(fPath, 'r') as f:
        log = f.read()
    
    logSplited: list[str] = log.split()
    if len(logSplited) <= 1:
        print("ATENÇÃO: Informações de contagem foram salvas " + \
                "incorretamente no ultimo uso!")
        print("ATENÇÃO: Iniciando contador em zero(0)")
        return 0
        
    logDate: str = logSplited[0]
    logCont: str = logSplited[1]
    
    if not logCont.isnumeric():
        print("ATENÇÃO: Contagem salva não está no formato correto!")
        print("ATENÇÃO: Iniciando contador em zero(0)")
        return 0
    
    today: str = str(date.today())
    if today != logDate:
        print("> Novo dia detectado!")
        print("- Ultima data de uso foi " + logDate + \
                " com " + logCont + " mensagens enviadas!")
        print("> Iniciando contador em zero(0) para um novo dia de envio")
        return 0
    
    return int(logCont)

def salvaContador(contador: int, fPath:str = contPath):
    today: str = str(date.today())
    log: str = today + ' ' + str(contador)
    with open(fPath, 'w') as f:
        f.write(log)

def salvarPlanilha(updated: tuple):        
    workbook = load_workbook(fileName)
    sheet = workbook.active

    linha: int = updated[0]
    valor: str = updated[1]
    sheet['C' + str(linha + 2)] = valor

    workbook.save(fileName)

def geraLog(element):
    now = datetime.now()
    dt = now.strftime("[ %Y-%m-%d %H:%M:%S ]")
    name = str(element[0].split('.')[0])
    fone = str(element[1])
    status = str(element[2])
    
    message = "Falha ao enviar mensagem"
    if status == "ok":
        message = "Mensagem enviada com sucesso"
    message += " (t:{} n:{})".format(fone, name)
    
    log = dt + " "
    log += "Message Log: "
    log += " " + message
    print(log)
    
    with open(logPath, "r+") as f:
        fData = f.read()
        if len(fData) == 0:
            f.write("WhatScript Log File\n")
            f.write("- File created by program")
        f.write('\n')
        f.write(log)

def geraLogErro(element: Exception):
    now = datetime.now()
    dt = now.strftime("[ %Y-%m-%d %H:%M:%S ]")
    
    message = "Ocorreu um erro inesperado ({})".format(str(element))
    
    log = dt + " "
    log += "Error Log: "
    log += " " + message
    print(log)
    
    with open(logPath, "r+") as f:
        fData = f.read()
        if len(fData) == 0:
            f.write("WhatScript Log File\n")
            f.write("- File created by program")
        f.write('\n')
        f.write(log)
        
def enviarMensagem():
    pass

print(introducao)

contadorDiario: int = defineContador()
if contadorDiario >= limiteMensagens:
    print("Você já atingiu o limite de mensagens de hoje!")
    pause("Pressione Enter para finalizar...")
    exit()

print("\nContador de mensagens do dia: " + str(contadorDiario))

salvaContador(contadorDiario)

# Printing the current working directory
# print("The Current working directory is: {0}".format(os.getcwd()))

tkinter.Tk().withdraw()

print("\nIndique o local onde está a planilha de dados:")
filePath = filedialog.askopenfilename(
    title="Selecione a planilha com os dados",
    filetypes={("Planilha do Excel (.xlsx)", ".xlsx")},
    )
print(filePath)

# Changing the current working directory
if not os.path.exists(dataPath):
    os.makedirs(dataPath)
os.chdir(dataPath)

fileName: str = filePath

print("\nAbrindo planilha do Excel...")
try:
    df = pd.read_excel(io=fileName, dtype={"ENVIADOS": str})
except Exception as e:
    print(e)
    pause("Pressione Enter para finalizar...")
    exit()
    
# print(df)

if (df is None):
    print("\nERRO: Tabela está vazia")
    pause("Pressione Enter para finalizar...")
    exit()

print("Tabela encontrada - Número de registros:", len(df))

print("\nAbrindo Chrome com uso de cookies...")
options = webdriver.ChromeOptions()
options.add_argument("user-data-dir={}\\cache-chrome\\".format(dataPath))

nav = webdriver.Chrome(options=options)
nav.get("https://web.whatsapp.com")


print("Aguardando carregamento do Whatsapp Web...")
while len(nav.find_elements(By.ID, "side")) < 1:
    time.sleep(1)
time.sleep(2)

print("\nIniciando envio de mensagens...")
contadorMensagensEnviadas: int = 0

updated: tuple = ()

try:
    for linha in df.index:
        numero = df.loc[linha, "PHONE"]
        
        if df.loc[linha, "ENVIADOS"] == "ok":
            #print("Numero", numero, "ja recebeu uma mensagem!")
            continue
        
        print("Enviado mensagem para numero:", numero)
        
        nome = df.loc[linha, "NAME"]

        texto = parse.quote(f"{mensagem}")

        link = f"https://web.whatsapp.com/send?phone={numero}&text={texto}"

        nav.get(link)

        while len(nav.find_elements(By.ID, "side")) < 1:
            time.sleep(1)
        time.sleep(2)
        
        
        if not len(nav.find_elements(By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[1]')) < 1:
            df.loc[linha, "ENVIADOS"] = str("falhou")
            updated = (linha, "falhou")
            
            geraLog(df.loc[linha, :].values.flatten().tolist())
            
            if linha == len(df.index)-1:
                print("Fim da lista")
                break
                
            continue
            
        nav.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span').click()
        df.loc[linha, "ENVIADOS"] = str("ok")
        updated = (linha, "ok")
        
        geraLog(df.loc[linha, :].values.flatten().tolist())
        
        contadorMensagensEnviadas += 1
        contadorDiario += 1
        salvaContador(contadorDiario)

        if contadorDiario >= limiteMensagens:
            print("Limite de mensagens atingido")
            break

        if linha == len(df.index)-1:
            print("Fim da lista")
            break

        # Save the file
        print("Atualizando planilha...")
        salvarPlanilha(updated)
        print("Planilha atualizada!")

        time.sleep(5)
except Exception as e:
    print(e)
    geraLogErro(e)
    pause("Pressione Enter para finalizar...")

print()
if (contadorMensagensEnviadas == 0):
    print("Sem mensagens para enviar! Atualize a planilha com novos numeros")
    pause("Pressione Enter para finalizar...")
    exit()

if linha == len(df.index)-1:
    print("Atualize a planilha com novos numeros!")
else:
    print("Re-execute amanhã para concluir o envio de mensagens")
    
print("- Foram enviadas", contadorMensagensEnviadas, "mensagens nessa execução\n")
print("- Hoje ja foram enviadas", contadorDiario, "mensagens no total!\n")

print("Programa encerrado!")
pause("Pressione Enter para finalizar...")