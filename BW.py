from selenium import webdriver
import time
import pyautogui
from UserBW import login,senha

navegador = webdriver.Chrome()

time.sleep(3)

navegador.get("http://pbwpmwvm1dc.nordestao.com.br:50400/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?TEMPLATE=DASH01&DUMMY=0")
pyautogui.typewrite(login)
pyautogui.press('tab')
pyautogui.typewrite(senha)
pyautogui.press('enter')
