import sys
import time
from playwright.sync_api import sync_playwright
import time
from Login import Login

login = Login()

email = str(login.get_email)
senha = str(login.get_senha())

protocoloBoleto = '71907965'

with sync_playwright() as p:
    browser = p.chromium.launch()
    context = browser.new_context()
    page = context.new_page()

    page.goto('https://safe2pay.com.br/login?redirect=%2F')
    page.fill('input[id=inputLoginMail]', str(login.get_email()))
    page.fill('input[id=inputLoginPassword]', str(login.get_senha()))
    page.locator('//*[@id="SignInModal"]/div/div/div[3]/button').click()
    time.sleep(15)
    page.goto('https://admin.safe2pay.com.br/transacoes')


    if len(protocoloBoleto) == 10:

        page.locator('button.btn.btn-default[ng-click="CreatedDateInitial_Change();CreatedDateInitialOpen()"]').click()
        page.locator('button.btn.btn-default.uib-clear.ng-binding[ng-click="select(null, $event)"]').click()
        page.locator('button[class="btn btn-outline dropdown-toggle"]').click()
        page.locator('input[type="checkbox"][ng-model="Columns.Reference"]').click()
        page.fill('input[id="txtFiltroCompradorReferencia"]', protocoloBoleto)
        page.locator('button.btn.btn-animate.btn-animate-side.btn-safe2pay.waves-effect.waves-light.pull-right.btn-mg-r[ng-click="Filtrar_click()"]').click()
        page.locator('//span[text()="NENHUM"]').click()
        page.locator('//button[contains(text(), "Liberar Boleto")]').click()
        time.sleep(3)
        page.locator('button[class="swal-button swal-button--confirm"]').click()
        time.sleep(3)
        print("Boleto liberado com sucesso ✅")
        validador = False
        browser.close()

    else:

        page.locator('button.btn.btn-default[ng-click="CreatedDateInitial_Change();CreatedDateInitialOpen()"]').click()
        page.locator('button.btn.btn-default.uib-clear.ng-binding[ng-click="select(null, $event)"]').click()
        page.locator('button[class="btn btn-outline dropdown-toggle"]').click()
        page.fill('input[id="txtFiltroCodigo"]', protocoloBoleto)
        page.locator('button.btn.btn-animate.btn-animate-side.btn-safe2pay.waves-effect.waves-light.pull-right.btn-mg-r[ng-click="Filtrar_click()"]').click()
        page.locator('//span[text()="NENHUM"]').click()
        page.locator('//button[contains(text(), "Liberar Boleto")]').click()
        time.sleep(3)
        page.locator('button[class="swal-button swal-button--confirm"]').click()
        time.sleep(3)
        print("Boleto liberado com sucesso ✅")
        validador = False
        browser.close()




