#Bibliotecas padrão
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from datetime import datetime, timedelta
import time
import pandas as pd
import os
from pathlib import Path
import fitz

#Bibliotecas para aguardar carregamento do site
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

#encontra o dia util imediatamente anterior a hoje
def ultimo_dia_util():
    hoje = datetime.today()
    variacao = timedelta(max(1,(hoje.weekday() + 6) % 7 - 3))
    #variacao = datetime.timedelta(max(2,(hoje.weekday() + 6) % 7 - 3))
    D_1 = hoje - variacao
    return D_1

#Pasta para salvar arquivo
options = webdriver.ChromeOptions()
prefs = {"download.default_directory" : "Z:\Operacional\Automacao\BoletimDiarioB3",
         "plugins.plugins_disabled" : "Chrome PDF Viewer",
         "plugins.always_open_pdf_externally": True,
        }
options.add_experimental_option("prefs",prefs) #Allows users to add these preferences to their Selenium webdriver object.
#options.add_argument('--headless')

#Cria uma instância do Chrome
driver = webdriver.Chrome(executable_path='./chromedriver.exe', options=options)

#Entra no site
driver.get('https://www.b3.com.br/pt_br/market-data-e-indices/servicos-de-dados/market-data/consultas/boletim-diario/boletim-diario-do-mercado/')

#Aguarda site carregar
driver.implicitly_wait(10)

#Rejeita todos os cookies
cookies = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, "//button[text()='REJEITAR TODOS OS COOKIES']")))
cookies.click()

#Muda para o iframe
driver.switch_to.frame(driver.find_element_by_xpath("//iframe[@id='bvmf_iframe']"))

#Clica para abrir a opção do Boletim Diário
WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, "//a[@class='plus-btn']"))).click()

#data_analise = date.format()
data = datetime.now()
ano = data.year
mes = data.month
dia = data.day -1
data_analise = ultimo_dia_util().strftime("%d/%m/%Y")
data_arquivo = ultimo_dia_util().strftime("%Y%m%d")

#Nomes dos arquivos
pasta_principal = Path("Z:/Operacional/Automacao/BoletimDiarioB3")
nome_pdf = "BDI_00_" + data_arquivo + '.pdf'
arquivo_excel = "CupomLimpo.csv"
caminho_excel = pasta_principal / arquivo_excel
caminho_pdf = pasta_principal / nome_pdf
planilha_excel = "CupomLimpo"

#Checa se arquivo a ser baixado já existe
if os.path.exists(caminho_pdf):
    os.remove(caminho_pdf)

#Baixa o arquivo
xpath = "//a[text()='Boletim Diário de:']"
link = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, xpath)))
link.click()
time.sleep(10)

#Abre pdf e pega o valor do cupom limpo
with fitz.open(caminho_pdf) as doc:
    text = ""
    for page in doc:
        text += page.get_text()
    cupom_limpo = text.find("DOL-CL")
    cupom_limpo=text[cupom_limpo+7:cupom_limpo+13]
    cupom_limpo = float(cupom_limpo.replace(",","."))

#Cria o data frame
data = {'Data':[data_analise], 'Valor': [cupom_limpo]}
df = pd.DataFrame(data)

#Coloca dado no fim do arquivo csv
df.to_csv(caminho_excel, mode='a', index=False, header=False)

driver.quit()