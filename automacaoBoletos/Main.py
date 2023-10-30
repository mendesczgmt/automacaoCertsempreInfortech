import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from Login import Login

browser = webdriver.Chrome()
browser.get("https://www.safe2pay.com.br/")
browser.maximize_window()

time.sleep(2)