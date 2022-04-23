from openpyxl import load_workbook
import os
from selenium.webdriver.common.by import By
from selenium import webdriver as opcoesSelenium
from selenium.webdriver.common.keys import Keys
import pyautogui as tempoEspera

nome_arquivo_cep = "C:\RPA\Endereco.xlsx"
planilhaEndereco = load_workbook(nome_arquivo_cep)

sheet_selecionada = planilhaEndereco['CEP']

navegador = opcoesSelenium.Chrome()
navegador.get("https://buscacepinter.correios.com.br/app/endereco/index.php")

navegador.find_element(By.NAME, "endereco").send_keys("30190-921")

tempoEspera.sleep(2)

navegador.find_element(By.NAME, "btn_pesquisar").click()

tempoEspera.sleep(3)

for i in range(2, len(sheet_selecionada['A']) + 1):
    tempoEspera.sleep(3)

    navegador.find_element(By.NAME, "btn_voltar").click()

    cepPesquisa = sheet_selecionada['A%s' % i].value

    tempoEspera.sleep(2)

    navegador.find_element(By.NAME, "endereco").send_keys(cepPesquisa)

    tempoEspera.sleep(2)

    navegador.find_element(By.NAME, "btn_pesquisar").click()

    tempoEspera.sleep(5)

    rua = navegador.find_element(By.XPATH, '//*[@id="resultado-DNEC"]/tbody/tr/td[1]' ).text
    print("Rua: ", rua)

    bairro = navegador.find_element(By.XPATH, '//*[@id="resultado-DNEC"]/tbody/tr/td[2]').text
    print("Bairro: ", bairro)

    cidade = navegador.find_element(By.XPATH, '//*[@id="resultado-DNEC"]/tbody/tr/td[3]').text
    print("Cidade: ", cidade)

    cep = navegador.find_element(By.XPATH, '//*[@id="resultado-DNEC"]/tbody/tr/td[4]').text
    print("CEP: ", cep)

    sheet_Endereco = planilhaEndereco['Endereco']

    linha = len(sheet_Endereco['A']) + 1
    colunaA = "A" + str(linha)
    colunaB = "B" + str(linha)
    colunaC = "C" + str(linha)
    colunaD = "D" + str(linha)

    sheet_Endereco[colunaA] = rua
    sheet_Endereco[colunaB] = bairro
    sheet_Endereco[colunaC] = cidade
    sheet_Endereco[colunaD] = cep

planilhaEndereco.save(filename= nome_arquivo_cep)

os.startfile(nome_arquivo_cep)