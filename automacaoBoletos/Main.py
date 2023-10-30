import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from Login import Login

browser = webdriver.Chrome()
browser.get("https://www.safe2pay.com.br/")
browser.maximize_window()
protocolo = "1004846398"

login = Login()

browser.find_element(By.XPATH, '//*[@id="navbarSupportedContent"]/ul[2]/li[1]').click()
time.sleep(1)
browser.find_element(By.ID, 'inputLoginMail').send_keys(str(login.get_login()))
browser.find_element(By.ID, 'inputLoginPassword').send_keys(str(login.get_senha()))
browser.find_element(By.XPATH, '//*[@id="SignInModal"]/div/div/div[3]/button').click()

time.sleep(10)

browser.find_element(By.XPATH, '/html/body/div[4]/div/div[3]/div/div/div/div/div').click()

time.sleep(5)

def liberarBoleto():

    browser.find_element(By.XPATH, '//*[@id="floattable"]/tbody/tr[1]/td[12]/small[1]').click()
    browser.find_element(By.XPATH, '//*[@id="SidebarTransacao"]/div/div/div[2]/div[3]/div[1]/button').click()
    browser.find_element(By.XPATH, '/html/body/div[13]/div/div[4]/div[2]/button').click()


def buscarBoleto():
    browser.find_element(By.XPATH, '/html/body/div[4]/div[2]/div[1]/div[2]/div/div[2]/div/button[3]').click()
    time.sleep(1)
    try:
        browser.find_element(By.ID, 'floattable').click()
        liberarBoleto()
    except:
        time.sleep(1)
        browser.find_element(By.ID, 'btnLimparFiltros').click()
        time.sleep(1)
        browser.find_element(By.XPATH, '/html/body/div[4]/div[2]/div[1]/div[1]/form/div/div[11]/div/p/span/button').click()
        time.sleep(1)
        browser.find_element(By.XPATH, '/html/body/div[4]/div[2]/div[1]/div[1]/form/div/div[11]/div/p/div/ul/li[2]/div/div[1]/button').click()
        time.sleep(1)
        if(len(protocolo) == 10):
            browser.find_element(By.ID, 'txtFiltroCompradorReferencia').send_keys(protocolo)
        else:
            browser.find_element(By.ID, 'txtFiltroCodigo').send_keys(protocolo)
            browser.find_element(By.XPATH, '/html/body/div[4]/div[2]/div[1]/div[2]/div/div[2]/div/button[3]').click()
    
        



if(len(protocolo) == 10):
    try:
        browser.find_element(By.ID, 'txtFiltroCompradorReferencia').send_keys(protocolo)
        buscarBoleto()
    except:
        browser.find_element(By.XPATH, '/html/body/div[4]/div[2]/div[1]/div[2]/div/div[1]/div/button').click()
        time.sleep(1)
        browser.find_element(By.XPATH, '/html/body/div[4]/div[2]/div[1]/div[2]/div/div[1]/div/ul/li[4]/label/input').click()
        time.sleep(1)
        browser.find_element(By.ID, 'txtFiltroCompradorReferencia').send_keys(protocolo)
        buscarBoleto()
else:
    browser.find_element(By.ID, 'txtFiltroCodigo').send_keys(protocolo)
    buscarBoleto()




    

        
        
        
