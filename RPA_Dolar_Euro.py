from selenium import webdriver as opcoes_selenium
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pyautogui as tempoPausaComputador
import xlsxwriter
import os

meuNavegador = opcoes_selenium.Chrome()
meuNavegador.get("https://www.google.com.br")

tempoPausaComputador.sleep(8)

meuNavegador.find_element(By.NAME, "q").send_keys("Dolar hoje")

tempoPausaComputador.sleep(4)

meuNavegador.find_element(By.NAME, "q").send_keys(Keys.RETURN)

tempoPausaComputador.sleep(4)

valorDolarPeloGoogle = meuNavegador.find_elements(By.XPATH,'//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]')[0].text

#----------------------------
tempoPausaComputador.sleep(2)

meuNavegador.find_element(By.NAME, "q").send_keys("")
tempoPausaComputador.sleep(4)

tempoPausaComputador.press('tab')

tempoPausaComputador.sleep(4)

tempoPausaComputador.press('enter')

tempoPausaComputador.sleep(4)

meuNavegador.find_element(By.NAME, "q").send_keys("Euro hoje")

tempoPausaComputador.sleep(4)

meuNavegador.find_element(By.NAME, "q").send_keys(Keys.RETURN)

tempoPausaComputador.sleep(4)

valorEuroPeloGoogle = meuNavegador.find_elements(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]')[0].text

#print("Dolar: " + valorDolarPeloGoogle)
#print("Euro: " + valorEuroPeloGoogle)

nomecaminhoArquivo = 'C:\\RPA\\Dolar_Euro_Google.xlsx'
planilhaCriada = xlsxwriter.Workbook(nomecaminhoArquivo)
sheet1 = planilhaCriada.add_worksheet()

tempoPausaComputador.sleep(6)

sheet1.write("A1", "Dolar")
sheet1.write("B1", "Euro")
sheet1.write("A2", valorDolarPeloGoogle)
sheet1.write("B2", valorEuroPeloGoogle)

planilhaCriada.close()

os.startfile(nomecaminhoArquivo)

print("Dolar e Euro extraídos com sucesso!")