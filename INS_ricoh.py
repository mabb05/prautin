import os,json
import easygui as eg
import pyautogui as p
import INS_ricoh as insr
import shutil as sh
import tempfile, ctypes, locale
from win32comext.shell import shell
from playwright.sync_api import sync_playwright

path_project = os.path.dirname(__file__)
msgjson = json.load(open(f'{path_project}\\IMP\\doc\\MSG.json', 'r', encoding='utf8'))
msgjson['pImg'] = f'{path_project}\\IMP\\img\\'

temp_dir = tempfile.gettempdir()
install_file = os.path.join(temp_dir,"ins.exe")
lingua = locale.windows_locale[ctypes.windll.kernel32.GetUserDefaultUILanguage()]
delaypadr = 1

nomeimpressora = "RICOHNOME"
ipins = "192.168.50.40"

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

def epson3150():

    # Baixando instalador
    with sync_playwright() as pw:
        browser = pw.firefox.launch(headless=False)
        page = browser.new_page()
        page.goto('https://epson.com.br/Suporte/Impressoras/Impressoras-multifuncionais/Epson-L/Epson-L3150/s/SPT_C11CG86301?review-filter=Windows+10+64-bit')

        with page.expect_download() as download_pcl:
            page.locator("//a[@class='btn detail-button btn-primary btn-with-arrow js-tl-dl']").nth(0).click()
        download = download_pcl.value
        path = download.path()
        sh.copy(path, install_file)
        page.wait_for_timeout(100)

        epson_install(install_file)

def epson_install(arquivo):
    
    os.startfile(arquivo)
    p.sleep(5)

    esperaimg('INS_EPSON_primeira.png', tempo=30, clicar=True)
    p.sleep(delaypadr); p.press('enter')
    p.sleep(4)
    esperaimg('INS_EPSON_welcome.png', tempo=30, clicar=True)
    p.sleep(delaypadr); p.press('enter')
    p.sleep(delaypadr); p.press('enter')


    

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
    #p.sleep(delaypadr); p.press('down')
    #p.sleep(0.2);       p.press('down')
    #p.sleep(0.2);       p.press('enter')
    #p.sleep(delaypadr); p.press('up')
    #p.sleep(0.2);       p.press('up')
    #p.sleep(0.2);       p.press('tab')
    #p.sleep(0.2);       p.write(ipins)
    #p.sleep(0.2);       p.press('enter')

    p.sleep(delaypadr); p.press('space')
    p.sleep(0.2); p.press('down')
    p.sleep(0.2); p.press('down')
    p.sleep(0.2); p.press('down')
    p.sleep(0.2); p.press('down')
    p.sleep(0.2); p.press('enter')
    p.sleep(0.2); p.press('enter')

    p.sleep(delaypadr); p.press('enter')
    p.sleep(delaypadr); p.hotkey('ctrl','v')
    p.sleep(0.2); p.press('tab')
    p.sleep(0.2); p.press('enter')

    p.sleep(delaypadr); p.press('tab')
    p.sleep(0.2); p.press('tab')
    p.sleep(0.2); p.press('tab')
    p.sleep(0.2); p.press('tab')
    p.sleep(0.2); p.press('tab')
    p.sleep(0.2); p.press('tab')
    p.sleep(0.2); p.press('space')
    p.sleep(0.2); p.press('enter')
    p.sleep(0.2); p.press('down')
    p.sleep(0.2); p.press('enter')

    p.sleep(0.2); p.press('tab')
    p.sleep(0.2); p.press('enter')

    p.sleep(delaypadr); p.press('tab')
    p.sleep(0.2); p.press('tab')
    p.sleep(0.2); p.press('tab')
    p.sleep(0.2); p.press('tab')
    p.sleep(0.2); p.press('tab')

    p.sleep(delaypadr); p.press('r')
    p.sleep(delaypadr); p.press('tab')
    p.sleep(0.2); p.press('down')
    p.sleep(0.2); p.press('enter')

    p.sleep(3); p.write(nomeimpressora)

    p.sleep(delaypadr); p.press('enter')

    p.sleep(4)
    if p.locateCenterOnScreen("img\\INS_perm.png", confidence=0.70):
        p.sleep(0.2); p.press('tab')
        p.sleep(0.2); p.press('tab')
        p.sleep(0.2); p.press('enter')
    p.sleep(2)

    p.sleep(delaypadr); p.press('enter')
    p.sleep(delaypadr); p.press('enter')
    

    # WHAT TO DO WITH ADMIN BOX ???

# ricoh: 3410dn 3510dn 3710dn 377dn 377sf 310SF C3003 M/320F
# epson: 3250 6190 6490 6270 
# pantum: BM5100FDW 6550FDW
# hp: M1212NF 1102W 
# samsung: M4080FX 4070fr M4020ND SCX-483X

# how it detect and install USB on pcl?
# how it will install firmware?

# FAZER FUNCIONAR
# 
# ESCREVER AGUARDATELA 

def esperaimg(path,time=60,click=False,dx=0,dy=0):
    p.locateCenterOnScreen