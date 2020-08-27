"""
Created on Mon Jun 22 22:00:41 2020
@author: Henrique S. Kisaki
"""

"""
Automating a daily process using Pyautogui and Selenium libraries.
"""

from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import pyautogui
import time
import datetime
import os
from datetime import timedelta

options = webdriver.ChromeOptions()
options.add_argument('lang=pt-br')
driver = webdriver.Chrome(ChromeDriverManager().install())
action = ActionChains(driver)

class Telas_Monitoracao():
    #Initialize class
    def __init__ (self):
        self.wa = WindowsAction()
        self.datas = [{"instituicao":"Hospital São Luiz São Caetano do Sul", "modelo":"EN - Consumo Energia Elétrica EE CAG"}, {"instituicao":"Hospital São Luiz Itaim", "modelo":"EN - Eficiencia Energética - Medidor CAG"}, {"instituicao":"Hospital Unimed", "modelo":"EN - Consumo Energético - CAG"},{"instituicao":"Honda", "modelo":"EN - Eficiencia Energética CAG"}, {"instituicao":"Hospital São Luiz Morumbi", "modelo":"EIXAD5 - Relatório EE - CAG"}]
        

    #Set contratoManutencao window
    def contratoManutencao_Tela1 (self):
        time.sleep(2)
        driver.get('https://g5.oxyn.com.br/')
        login = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'user_email')))
        password = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'user_senha')))
        login.send_keys('') #Type your e-mail between ' '
        password.send_keys('') #Type your password between ' '
        button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//button[@type='submit']")))
        button.click()
        time.sleep(7)
        pyautogui.press('esc')
        time.sleep(2)
        resumo = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//div[@id='nav']//li[2]"))) #Wait until detect element
        resumo.click()
        time.sleep(3)
        marcador = driver.find_element_by_class_name("marker-label") 
        marcador.click() 
        time.sleep(2)
        action.reset_actions()
        contratoBox = driver.find_element_by_xpath("//ul[@id='labels']//li[contains(text(),'Contrato de Manutenção')]")
        filterBox = driver.find_element_by_id("filter")
        action.drag_and_drop(contratoBox, filterBox).perform()
        time.sleep(2)
        self.apply()

    def contratoManutencao_Tela2 (self):
        for i in range (2):
            self.callTab(i+1)
            marcador = driver.find_element_by_class_name("marker-label") 
            marcador.click() 
            time.sleep(2)
            action.reset_actions()
            contratoBox = driver.find_element_by_xpath("//ul[@id='labels']//li[contains(text(),'Contrato de Manutenção')]")
            filterBox = driver.find_element_by_id("filter")
            action.drag_and_drop(contratoBox, filterBox).perform()
            time.sleep(2)
            action.reset_actions()
            if i == 0:
                conexaoBox = driver.find_element_by_xpath("//ul[@id='labels']//li[contains(text(),'Conexão')]")
            else:
                conexaoBox = driver.find_element_by_xpath("//ul[@id='labels']//li[contains(text(),'Comunicação')]")
            filterBox1 = driver.find_element_by_xpath("//ul[@id='filter']//li[1]")
            action.drag_and_drop(conexaoBox, filterBox1).perform()
            time.sleep(2)
            closeButton = driver.find_element_by_xpath("//ul[@id='filter']//li[2]//a[@href='#']")
            closeButton.click()
            time.sleep(2)
            self.apply()

    #Apply filters
    def apply (self):
        applyButton = driver.find_element_by_xpath("//div[@class='ui-dialog-buttonset']//button[contains(text(),'Aplicar')]")
        applyButton.click()
        time.sleep(2)

    #Generate reports
    def generatorButton (self):
        generator = driver.find_element_by_xpath("//div[@class='ui-dialog-buttonset']//button[contains(text(),'Gerar')]")
        generator.click()
        time.sleep(2)

    #Calc a month before and current date
    def dateCalculator (self):
        currentDate = datetime.datetime.today()
        lastMonthDate = currentDate - timedelta(days=31)
        #Initial date - a month before
        initialDateField = driver.find_element_by_id('initial-date')
        initialDateField.click()
        initialDateField.clear()
        initialDateField.send_keys(lastMonthDate.strftime("%d/%m/%Y")) #Add last month date
        pyautogui.press('enter')
        #Final date - current date
        finalDateField = driver.find_element_by_id('final-date')
        finalDateField.click()
        finalDateField.clear()
        finalDateField.send_keys(currentDate.strftime("%d/%m/%Y"))
        pyautogui.press('enter')

    def openReport (self):
        self.callTab(4)
        for i in range (5):
            fields = self.datas[i]
            select = driver.find_element_by_xpath("//span[@id='select2-sites-container']")
            select.click()
            time.sleep(2)
            insertName = driver.find_element_by_xpath("//input[@class='select2-search__field']")
            insertName.send_keys(fields["instituicao"])
            time.sleep(1)
            pyautogui.press('enter')
            time.sleep(5)
            report = driver.find_element_by_class_name('report')
            report.click()
            time.sleep(5)
            reportModel = driver.find_element_by_id('select2-report-model-select2-container')
            reportModel.click()
            fieldModel = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.CLASS_NAME, 'select2-search__field')))
            falseModel = driver.find_element_by_xpath("//li[contains(text(), 'Nenhum resultado encontrado')]")
            while falseModel.is_displayed():
                fieldModel.click()
                time.sleep(1)
            fieldModel.send_keys(fields["modelo"])
            pyautogui.press('enter')
            self.dateCalculator()
            self.generatorButton()
            self.pageDown()
            if i == 4:
                pyautogui.hotkey('Ctrl','w') #close report tab
        time.sleep(2)


    #Create a new tab with oxyn g5 logged
    def callTab (self, i):
        driver.execute_script("window.open('');")
        driver.switch_to.window(driver.window_handles[i])
        driver.get("https://g5.oxyn.com.br/deviation")
        time.sleep(3)

    #Return to report page
    def pageDown (self):
        driver.switch_to.window(driver.window_handles[3])
        time.sleep(2)
        pyautogui.press('esc')
        time.sleep(1)

    #Recenter content if it isn't in middle screen
    def centerPositionScreen (self):
        driver.maximize_window()
        time.sleep(2)
        pyautogui.moveTo(x=-960, y=15) #Set cursor in first screen
        pyautogui.dragTo(x=960, y=10, duration=0.75) #Drag window to second
        time.sleep(2)
        pyautogui.moveTo(x=2880, y=15) #Set cursor in third screen
        pyautogui.dragTo(x=960, y=10, duration=0.75) #Drag window to second
        pyautogui.click(x=960, y=10)
        time.sleep(2)


class InstallRevolverTabExtension:
    #Access "Report Tabs" page
    def GoogleAccess (self):
        driver.get('https://chrome.google.com/webstore/detail/revolver-tabs/dlknooajieciikpedpldejhhijacnbda')
    
    #Click in "Usar no Chrome" button 
    def AddExtension (self):
        addButton = WebDriverWait(driver,20).until(EC.presence_of_element_located((By.XPATH,"//div[@class='g-c-x']//div[contains(text(),'Usar no Chrome')]")))
        addButton.click()
    
    #Access alert box and accept requisition
    def AcceptRevolverTabExtension (self):
        time.sleep(5)
        pyautogui.press('left')
        pyautogui.press('enter')

    def run(self):
        self.GoogleAccess()
        self.AddExtension()
        self.AcceptRevolverTabExtension()


class WindowsAction ():
    def __init__(self):
        pass

    def sendToFirstScreen (self):
        pyautogui.moveTo(x=53, y=15)
        #Change coordinates to send to first screen
        pyautogui.dragTo(x=-960, y=15, button='left', duration=0.75) 
        time.sleep(2)

    def sendAllToThird (self):
        pyautogui.click(x=1712, y=15) #Set cursor first report tab
        pyautogui.keyDown('win')
        for i in range (4):
            pyautogui.press('right')
            time.sleep(1)
        pyautogui.press('up')
        pyautogui.keyUp('win')
        time.sleep(2)
   
    def secondScreen(self):
        for i in range (2):
            pyautogui.moveTo(x=1985, y=15) #Set cursor first report tab on thrid window
            pyautogui.dragTo(x=767, y=15, duration=0.75) #drag to center of second window
            time.sleep(2)

    #find element and click
    def maximizePDF (self):
        pyautogui.moveTo(x=1838, y=820, duration=2.0)
        pyautogui.click()
        time.sleep(1.5)

    def reportPosition (self):
        tab = ['7', '6', '5', '4']
        clear = lambda: os.system("cls")
        wait = 0
        clear()
        while wait < 10:
            clear()
            print("Wait 10 seconds")
            wait = wait + 1
            time.sleep(0.8)
        print("Ready...")

        for i in range (4):
            pyautogui.moveTo(x=960, y=540) #center of a 1920x1080 screen
            pyautogui.hotkey('ctrl', tab[i])
            time.sleep(2)
            pyautogui.hotkey('ctrl', 'f')
            time.sleep(1)
            pyautogui.write('Grafico', interval='0.1')
            pyautogui.press('enter')
            time.sleep(1)
            pyautogui.press('esc')
            time.sleep(1)
            self.maximizePDF()
            time.sleep(1)


tm = Telas_Monitoracao()
wa = WindowsAction()
IR = InstallRevolverTabExtension()
tm.centerPositionScreen()
tm.contratoManutencao_Tela1()
tm.contratoManutencao_Tela2()
tm.openReport()
wa.reportPosition()
wa.sendToFirstScreen()
wa.sendAllToThird()
wa.secondScreen()
IR.run()
pyautogui.click(x=1799, y=64)
pyautogui.click(x=3719, y=60)
