from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl

#acessar o site
#https://www.kabum.com.br/computadores/pc/pc-gamer

driver = webdriver.Chrome()
driver.get('https://www.kabum.com.br/computadores/pc/pc-gamer')

#extrair todos os títulos
titulos = driver.find_elements(By.XPATH,"//span[@class='sc-d79c9c3f-0 nlmfp sc-cdc9b13f-16 eHyEuD nameCard']")

#extrair todos os preços
precos = driver.find_elements(By.XPATH,"//span[@class='sc-620f2d27-2 bMHwXA priceCard']")

#criando a planilha
workbook = openpyxl.Workbook()
#criando a página produtos
workbook.create_sheet('produtos')
#selecionando a página produtos
sheet_produtos = workbook['produtos']
#criando títulos de colunas
sheet_produtos['A1'].value = 'Produto'
sheet_produtos['B1'].value = 'Preços'


#inserir os esultados na planilha
#inserindo somente os produtos que contem as duas informações
for titulo, preco in zip(titulos, precos):
    sheet_produtos.append([titulo.text,preco.text])

#salvando na planilha
workbook.save('produtos.xlsx')