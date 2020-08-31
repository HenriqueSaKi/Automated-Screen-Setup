# -*- coding: utf-8 -*-
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

class Telas_Monitoracao():
    #Initialize class
    def __init__ (self):
        options = webdriver.ChromeOptions()
        options.add_argument('lang=pt-br')
        self.driver = webdriver.Chrome(ChromeDriverManager().install())
        self.wa = WindowsAction()
        self.datas = [{"instituicao":"Hospital São Luiz São Caetano do Sul", "modelo":"EN - Consumo Energia Elétrica EE CAG"}, {"instituicao":"Hospital São Luiz Itaim", "modelo":"EN - Eficiencia Energética - Medidor CAG"}, {"instituicao":"Hospital Unimed", "modelo":"EN - Consumo Energético - CAG"},{"instituicao":"Honda", "modelo":"EN - Eficiencia Energética CAG"}, {"instituicao":"Hospital São Luiz Morumbi", "modelo":"EIXAD5 - Relatório EE - CAG"}]

    #Set contratoManutencao window
    def contratoManutencao_Tela1 (self):
        action = ActionChains(self.driver)
        self.InstallRevolverTabExtension()
        time.sleep(2)
        self.driver.get('https://g5.oxyn.com.br/')
        time.sleep(2)
        login = self.driver.find_element_by_id("user_email")
        time.sleep(2)
        login.click()
        login.clear()   
        login.send_keys('')#Write your e-mail between ' ', like: 'example@example.com'
        password = self.driver.find_element_by_id("user_senha")
        time.sleep(2)
        password.click()
        password.clear()       
        password.send_keys('')#Write your password between ' ', like: 'PasswordExample123'
        time.sleep(2)
        submitButton = self.driver.find_element_by_xpath("//button[@type='submit']")
        submitButton.click()
        time.sleep(7)
        pyautogui.press('esc')
        time.sleep(2)
        #Wait until detect element
        resumo = WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.XPATH, "//div[@id='nav']//li[2]")))
        resumo.click()
        time.sleep(3)
        marcador = self.driver.find_element_by_class_name("marker-label") 
        marcador.click() 
        time.sleep(2)
        action.reset_actions()
        contratoBox = self.driver.find_element_by_xpath("//ul[@id='labels']//li[contains(text(),'Contrato de Manutenção')]")
        filterBox = self.driver.find_element_by_id("filter")
        action.drag_and_drop(contratoBox, filterBox).perform()
        time.sleep(2)
        self.apply()

    def contratoManutencao_Tela2 (self):
        for i in range (2):
            self.callTab(i+1)
            action = ActionChains(self.driver)
            marcador = self.driver.find_element_by_class_name("marker-label") 
            marcador.click() 
            time.sleep(2)
            action.reset_actions()
            contratoBox = self.driver.find_element_by_xpath("//ul[@id='labels']//li[contains(text(),'Contrato de Manutenção')]")
            filterBox = self.driver.find_element_by_id("filter")
            action.drag_and_drop(contratoBox, filterBox).perform()
            time.sleep(2)
            action.reset_actions()
            if i == 0:
                conexaoBox = self.driver.find_element_by_xpath("//ul[@id='labels']//li[contains(text(),'Conexão')]")
            else:
                conexaoBox = self.driver.find_element_by_xpath("//ul[@id='labels']//li[contains(text(),'Comunicação')]")
            filterBox1 = self.driver.find_element_by_xpath("//ul[@id='filter']//li[1]")
            action.drag_and_drop(conexaoBox, filterBox1).perform()
            time.sleep(2)
            closeButton = self.driver.find_element_by_xpath("//ul[@id='filter']//li[2]//a[@href='#']")
            closeButton.click()
            time.sleep(2)
            self.apply()

    def apply (self):
        applyButton = self.driver.find_element_by_xpath("//div[@class='ui-dialog-buttonset']//button[contains(text(),'Aplicar')]")
        applyButton.click()
        time.sleep(3)

    def generatorButton (self):
        generator = self.driver.find_element_by_xpath("//div[@class='ui-dialog-buttonset']//button[contains(text(),'Gerar')]")
        generator.click()
        time.sleep(2)

    def dateCalculator (self):
        currentDate = datetime.datetime.today()
        lastMonthDate = currentDate - timedelta(days=31)
        #Initial date
        initialDateField = self.driver.find_element_by_id('initial-date')
        initialDateField.click()
        initialDateField.clear()
        time.sleep(1)
        initialDateField.send_keys(lastMonthDate.strftime("%d/%m/%Y")) #Add last month date
        pyautogui.press('enter')
        #Final date
        finalDateField = self.driver.find_element_by_id('final-date')
        finalDateField.click()
        finalDateField.clear()
        time.sleep(1)
        finalDateField.send_keys(currentDate.strftime("%d/%m/%Y"))
        pyautogui.press('enter')
        time.sleep(2)

    def openReport (self):
        self.callTab(3)
        for i in range (5):
            fields = self.datas[i]
            select = self.driver.find_element_by_xpath("//span[@id='select2-sites-container']")
            select.click()
            time.sleep(2)
            insertName = self.driver.find_element_by_xpath("//input[@class='select2-search__field']")
            insertName.send_keys(fields["instituicao"])
            time.sleep(1)
            pyautogui.press('enter')
            time.sleep(5)
            report = self.driver.find_element_by_class_name('report')
            report.click()
            time.sleep(5)
            reportModel = self.driver.find_element_by_id('select2-report-model-select2-container')
            reportModel.click()
            findModel = WebDriverWait(self.driver, 120).until(EC.presence_of_element_located((By.CLASS_NAME, 'select2-search__field')))
            time.sleep(5)
            findModel.send_keys(fields["modelo"])
            pyautogui.press('enter')
            time.sleep(1)
            self.dateCalculator()
            self.generatorButton()
            self.pageDown()
            if i == 4:
                pyautogui.hotkey('Ctrl','w') #close report tab
        time.sleep(2)


    def callTab (self, i):
        self.driver.execute_script("window.open('');")
        self.driver.switch_to.window(self.driver.window_handles[i])
        self.driver.get("https://g5.oxyn.com.br/deviation")
        time.sleep(3)

    def pageDown (self):
        self.driver.switch_to.window(self.driver.window_handles[3])
        time.sleep(2)
        pyautogui.press('esc')
        time.sleep(1)

    def centerPositionScreen (self):
        self.driver.maximize_window()
        time.sleep(2)
        pyautogui.moveTo(x=-960, y=15) #Set cursor in first screen
        pyautogui.dragTo(x=960, y=10, duration=0.75) #Drag window to second
        time.sleep(2)
        pyautogui.moveTo(x=2880, y=15) #Set cursor in third screen
        pyautogui.dragTo(x=960, y=10, duration=0.75) #Drag window to second
        pyautogui.click(x=960, y=10)
        time.sleep(2)

    def GoogleAccess (self):
        self.driver.get('https://chrome.google.com/webstore/detail/revolver-tabs/dlknooajieciikpedpldejhhijacnbda?hl=pt-BR')
    
    def AddExtension (self):
        #resumo = WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.XPATH, "//div[@id='nav']//li[2]")))
        addButton = WebDriverWait(self.driver,20).until(EC.presence_of_element_located((By.XPATH,"//div[@class='g-c-x']//div[contains(text(),'Usar no Chrome')]")))
        addButton.click()
    
    def InstallRevolverTabExtension (self):
        self.GoogleAccess()
        time.sleep(3)
        self.AddExtension()
        time.sleep(5)
        pyautogui.press('left')
        pyautogui.press('enter')
        time.sleep(5)

class WindowsAction ():
    def __init__(self):
        pass

    def sendToFirstScreen (self):
        pyautogui.moveTo(x=53, y=15)
        #Change coordinates to send to first screen
        pyautogui.dragTo(x=-960, y=15, button='left', duration=0.75) 
        time.sleep(2)

    def sendAllToThird (self):
        pyautogui.click(x=1754, y=15) #Set cursor first report tab
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

    def maximizePDF (self):
        pyautogui.moveTo(x=1853, y=870, duration=2.0)
        time.sleep(1)
        pyautogui.click()
        time.sleep(1.5)

    def reportPosition (self):
        tab = ['8', '7', '6', '5', '4']
        clear = lambda: os.system("cls")
        wait = 0
        clear()
        while wait < 30:
            clear()
            print("Wait 30 seconds")
            wait = wait + 1
            time.sleep(0.8)
        print("Ready...")

        for i in range (5):
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
tm.centerPositionScreen()
tm.contratoManutencao_Tela1()
tm.contratoManutencao_Tela2()
tm.openReport()
wa.reportPosition()
wa.sendToFirstScreen()
wa.sendAllToThird()
wa.secondScreen()
pyautogui.click(x=1822, y=53)
pyautogui.click(x=3354, y=53)
