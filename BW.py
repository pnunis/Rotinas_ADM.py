import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import pyautogui
from UserBW import login,senha

navegador = webdriver.Chrome()
navegador.get('http://pbwpmwvm1dc.nordestao.com.br:50400/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?TEMPLATE=DASH01&DUMMY=0')
time.sleep(3)
pyautogui.typewrite(login)
pyautogui.press('tab')
pyautogui.typewrite(senha)
pyautogui.press('enter')
time.sleep(4)
#navegador.find_element(By.XPATH,'/html/body/table[1]/tbody/tr/td[1]/table/tbody/tr[3]/td[2]/div/div/div/div/table/tbody/tr/td/div/span[4]/span[1]').click()


