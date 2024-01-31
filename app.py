from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl

driver = webdriver.Chrome()
driver.get('https://www.novaliderinformatica.com.br/computadores-gamers')

#extracao dos titulos
titulos = driver.find_elements(By.XPATH,"//a[@class='nome-produto']")
#extracao dos precos
precos = driver.find_elements(By.XPATH, "//strong[@class='preco-promocional']")

workbook = openpyxl.Workbook()
workbook.create_sheet('produtos')
sheet_produtos = workbook['produtos']
sheet_produtos['A1'].value = 'Produto'
sheet_produtos['B1'].value = 'Preco'


#inserir os titulos na planilha
for titulo, preco in zip(titulos, precos):
    sheet_produtos.append([titulo.text, preco.text])

workbook.save('produtos.xlsx')