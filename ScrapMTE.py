# -*- encoding: utf-8 -*-
from time import sleep
from selenium import webdriver
from selenium.webdriver.firefox.options import Options

def scrap(CNPJ):
    print("Iniciando Request1 CONVENCAO - "+ str(CNPJ))
    ops = Options()
    ops.headless = True
    navegador = webdriver.Firefox(options=ops, executable_path=r'./driver/geckodriver.exe')
    navegador.get("http://www3.mte.gov.br/sistemas/mediador/ConsultarInstColetivo")
    label1 = navegador.find_element_by_xpath("//select[@id='cboTPRequerimento']//option[@value='convencao']").click()
    label2 = navegador.find_element_by_xpath("//select[@id='cboSTVigencia']//option[@value='1']").click()
    checkbox = navegador.find_element_by_xpath("//input[@id='chkNRCNPJ']")
    textCNPJ = navegador.find_element_by_id("txtNRCNPJ")

    checkbox.click()
    textCNPJ.send_keys(CNPJ)
    aux = []
    try:
        Botao = navegador.find_element_by_xpath("//button[@id='btnPesquisar']").click()
        sleep(5)
        Dado = navegador.find_elements_by_xpath("//table[@class='TbForm']//tbody//tr/td[@ALIGN='left']")
        for x in range(len(Dado)):
            aux.append(Dado[x].text)
    except Exception as erro:
        print(erro)
        navegador.quit()
    navegador.quit()
    return aux

def scrap2(CNPJ):
    print("Iniciando Request2 TERMO - "+ str(CNPJ))
    ops = Options()
    ops.headless = True
    navegador = webdriver.Firefox(options=ops, executable_path=r'./driver/geckodriver.exe')
    navegador.get("http://www3.mte.gov.br/sistemas/mediador/ConsultarInstColetivo")
    label1 = navegador.find_element_by_xpath("//select[@id='cboTPRequerimento']//option[@value='termoAditivoConvecao']").click()
    label2 = navegador.find_element_by_xpath("//select[@id='cboSTVigencia']//option[@value='1']").click()
    checkbox = navegador.find_element_by_xpath("//input[@id='chkNRCNPJ']")
    textCNPJ = navegador.find_element_by_id("txtNRCNPJ")

    checkbox.click()
    textCNPJ.send_keys(CNPJ)
    try:
        Botao = navegador.find_element_by_xpath("//button[@id='btnPesquisar']").click()
        sleep(5)
        Dado = navegador.find_elements_by_xpath("//table[@class='TbForm']//tbody//tr/td[@ALIGN='left']")
        aux2 = []
        for x in range(len(Dado)):
            aux2.append(Dado[x].text)
    except Exception as erro:
        print(erro)
        navegador.quit()
    navegador.quit()
    return aux2
