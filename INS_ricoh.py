import os,json
import easygui as eg
import pyautogui as p
import INS_ricoh as insr
import shutil as sh
import tempfile, ctypes, locale
from win32comext.shell import shell
from playwright.sync_api import sync_playwright

path_project = os.path.dirname(__file__)
temp_dir = tempfile.gettempdir()
install_file = os.path.join(temp_dir,"ins.exe")
lingua = locale.windows_locale[ctypes.windll.kernel32.GetUserDefaultUILanguage()]
delaypadr = 1

with open("param.txt", "r") as file:
    ipins = file.readline()
    nomeimp = file.readline()

def restart_spooler():
    print("-------- INICIANDO/REINICIANDO SPOOLER --------")
    os.system("net stop spooler")
    p.sleep(4)
    os.system("net start spooler")


def restart_explorer():
    print("-------- INICIANDO/REINICIANDO EXPLORER --------")
    os.system("taskkill /F /IM explorer.exe")
    p.sleep(4)
    os.system("start explorer")


def ricoh3710():

    # Baixando instalador
    with sync_playwright() as pw:
        browser = pw.firefox.launch(headless=False)
        page = browser.new_page()
        page.goto('http://support.ricoh.com/bb/html/dr_ut_e/apc/model/sp3710sf/sp3710sf.htm')

        with page.expect_download() as download_pcl:
            page.locator("//a[@class='button cnv01 icon mw200 rsp_w100p']").nth(1).click()
        download = download_pcl.value
        path = download.path()
        print (path)
        sh.copy(path, install_file)
        page.wait_for_timeout(100)

    pcl6(install_file)


def ricoh3510():

    # Baixando instalador
    with sync_playwright() as pw:
        browser = pw.firefox.launch(headless=False)
        page = browser.new_page()
        page.goto('http://support.ricoh.com/bb/html/dr_ut_e/apc/model/sp35s/sp35s.htm')

        with page.expect_download() as download_pcl:
            page.locator("//a[@class='button cnv01 icon mw200 rsp_w100p']").nth(0).click()
        download = download_pcl.value
        path = download.path()
        sh.copy(path, install_file)
        page.wait_for_timeout(100)
    
    pcl6(install_file)


def pcl6(arquivo):

    os.startfile(arquivo)
    p.sleep(2)
    x,y = p.locateCenterOnScreen("img\\INS_unzip.png", confidence=0.70)
    p.sleep(delaypadr); p.click(x,y)
    p.sleep(delaypadr); p.press('enter')
    p.sleep(delaypadr); p.press('enter')
    p.sleep(delaypadr); p.hotkey('ctrl','c')
    p.sleep(delaypadr); p.press('enter')

    p.sleep(delaypadr); os.system("control printers")

    p.sleep(2);         p.press('tab')
    p.sleep(0.2);       p.press('tab')
    p.sleep(0.2);       p.press('tab')
    p.sleep(0.2);       p.press('tab')
    p.sleep(0.2);       p.press('right')
    p.sleep(0.2);       p.press('enter')
    p.sleep(delaypadr); p.press('space')

    # IP
    # p.sleep(delaypadr); p.press('space')
    # p.sleep(delaypadr); p.press('down')
    # p.sleep(0.2);       p.press('down')
    # p.sleep(0.2);       p.press('enter')
    # p.sleep(delaypadr); p.write(ipins)
    # p.sleep(0.2);       p.press('tab')
    # p.sleep(0.2);       p.press('tab')
    # p.sleep(0.2);       p.press('space')
    # p.sleep(0.2);       p.press('enter')

    # PORT
    p.sleep(delaypadr); p.press('space')
    p.sleep(0.2);       p.press('down', 4, 0.2)
    p.sleep(0.2);       p.press('space')
    p.sleep(0.2);       p.press('enter')
    p.sleep(0.2);       p.press('enter')
    
    # input("---CABO")

    p.sleep(delaypadr); p.press('tab')
    p.sleep(0.2);       p.press('space')
    p.sleep(0.2);       p.hotkey('ctrl','v')
    p.sleep(0.2);       p.write('\\DISK1')
    p.sleep(0.2);       p.press('tab')
    p.sleep(0.2);       p.press('space')
    
    p.sleep(delaypadr)
    p.sleep(0.2);       p.write('GXE6N.INF')
    p.sleep(0.2);       p.press('enter')
    p.sleep(0.2);       p.press('up')
    p.sleep(0.2);       p.press('enter')

    p.sleep(delaypadr)
    p.sleep(0.2);       p.press('tab', 6, 0.2)
    p.sleep(0.2);       p.press('down')
    p.sleep(0.2);       p.press('down')
    p.sleep(0.2);       p.press('enter')

    p.sleep(delaypadr)

    # if p.locateCenterOnScreen('img\\INS_replace.png'):
    p.sleep(0.2);       p.press('enter')
    p.sleep(0.2);       p.press('enter')
    p.sleep(1);       p.press('enter')
    p.sleep(1);       p.press('esc',6,0.2)

    print("Impressora instalada!")


def esperaimg(path,time=60,click=False,dx=0,dy=0):
    
    for _ in range(time):
        if p.locateCenterOnScreen(path):
            return True
        time.sleep(1)
    return False