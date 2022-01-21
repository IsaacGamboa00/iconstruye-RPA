from selenium import webdriver

from time import *
from datetime import date
from time import *
from selenium import webdriver
from selenium.webdriver.support.select import Select
from scripts.fileOperations import FileOperation
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementNotInteractableException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.support import expected_conditions as EC


class botService:

    def __init__(self, driverOptions, downloadPath, driverPath):
        self.driverOptions = driverOptions
        self.downloadPath = downloadPath
        self.pathDriver = driverPath
        self.url = "https://www.iconstruye.com/index.html"
        self.operation = FileOperation.Operation(downloadPath)
        self.i=0
        self.username = "fvalenzuela"
        self.password = "Santi2017"
        self.orgName = "mkci"
        self.today = date.today()
        self.lastDate = self.today.strftime("%d-%m-%Y")
        self.initialDate = "01-01-2019"
        self.customers = [
            #ultimos clientes agregados
            {
                "customerID": "55912",
                "nameFile": "E011087"
            },
            {
                "customerID": "59235",
                "nameFile": "M001002"
            },
            #--------------------
          
            {
                "customerID": "55020",
                "nameFile": "E011083"
            },
            {
                "customerID": "55021",
                "nameFile": "E011084"
            },
            {
                "customerID": "55700",
                "nameFile": "E011085"
            },
            {
                "customerID": "55914",
                "nameFile": "E011086"
            },
            {
                "customerID": "52016",
                "nameFile": "E024022"
            },
            {
                "customerID": "51841",
                "nameFile": "E032001"
            },

            #nuevos cleintes
            {
                "customerID": "58356",
                "nameFile": "E024024"
            },
            {
                "customerID": "58739",
                "nameFile": "E024025"
            },
            {
                "customerID": "58716",
                "nameFile": "E031003"
            },
            {
                "customerID": "58713",
                "nameFile": "E011088"
            },{
                "customerID": "60020",
                "nameFile": "E011089"
            },{
                "customerID": "59776",
                "nameFile": "E024026"
            },{
                "customerID": "60020",
                "nameFile": "E011089"
            },{
                "customerID": "59776",
                "nameFile": "E024026"
            }
        ]

    def run(self):
        
        for customer in self.customers:
            self.operation.removeFile()
            self.downloadReports(customer["customerID"])
            self.operation.convertFormatToXls("prueba")
            self.operation.convertFormatToCsv(customer["nameFile"])
            self.operation.sendEmail(customer["nameFile"],"Se adjunta el archivo asociado.")
            print("se envio: "+customer["nameFile"])
            self.operation.removeFile()
            self.operation.removeFile()

        self.operation.removeFile()
        self.downloadReportsFacturas()
        print("se envio: Facturas")

        self.operation.removeFile()
        self.downloadReportsSubContrato()
        self.operation.renameFile("SC.xlsx")
        self.operation.sendEmail("SC", "Se adjunta el archivo asociado.")
        print("se envio: Sub contrato")

        self.operation.removeFile()
        self.downloadNotasCorrecion()
        self.operation.renameFile("NdeCredito.xlsx")
        self.operation.sendEmail("NdeCredito", "Se adjunta el archivo asociado.")
        print("se envio: Notas Correccion")
        



    def login(self, driver):
        driver.get(self.url)
        sleep(5)
        driver.find_element_by_class_name("js-open-login").click()
        sleep(5)
        iframe = driver.find_element_by_class_name("login ")
        driver.switch_to.frame(iframe)
        driver.find_element_by_id("txtUsuario").send_keys(self.username)
        driver.find_element_by_id("txtOrganizacion").send_keys(self.orgName)
        driver.find_element_by_id("txtClave").send_keys(self.password)
        sleep(5)
        driver.find_element_by_xpath("//input[@name='btnIngresar']").click()
        driver.switch_to.default_content()

    def downloadReportsSubContrato(self):
        with webdriver.Chrome(executable_path=self.pathDriver, chrome_options=self.driverOptions) as driver:
            try:
                self.login(driver)
                sleep(5)
                actions = ActionChains(driver)
                target= driver.find_element_by_css_selector("div#ControlMenu1_sec table tr td:nth-of-type(7)")
                actions.move_to_element(target).perform()
                sleep(1)
                target2=target.find_element_by_class_name("SubMenu ").find_element_by_css_selector("div table tr tr:nth-of-type(6)")
                actions.move_to_element(target2).perform()
                sleep(1)
                target3=target2.find_element_by_class_name("SubMenu ").find_element_by_css_selector("div table tr tr:nth-of-type(3)")
                target3.click()
                sleep(10)
                driver.switch_to.frame("detalle")
                select_element = driver.find_element_by_id("lstMoneda")
                select_object = Select(select_element)
                select_object.select_by_value('-1')
                driver.find_element_by_id("rngFechaCreacionFECHADESDE").clear()
                driver.find_element_by_id("rngFechaCreacionFECHADESDE").send_keys("01-01-2019")
                driver.find_element_by_id("rngFechaCreacionFECHAHASTA").clear()
                driver.find_element_by_id("rngFechaCreacionFECHAHASTA").send_keys("10-11-2025")
                sleep(5)
           
                select_element = driver.find_element_by_id("lstOrgC")
                select_object = Select(select_element)
                select_object.select_by_value('-1')
                sleep(7)
                driver.find_element_by_id("btnBuscar").click()
                sleep(15)
                try:
                    driver.switch_to.alert.accept()
                except:
                    pass
                sleep(10)
                driver.find_element_by_id("lnkExcel").click()
                sleep(30)
                driver.close()
            except:
                driver.close()
                self.downloadReportsSubContrato()

    def downloadReports(self,customer):
        with webdriver.Chrome(executable_path=self.pathDriver, chrome_options=self.driverOptions) as driver:
            try:
                self.login(driver)
                sleep(5)
                actions = ActionChains(driver)
                target= driver.find_element_by_css_selector("div#ControlMenu1_sec table tr td:nth-of-type(6)")
                actions.move_to_element(target).perform()
                sleep(7)
                target2=target.find_element_by_class_name("SubMenu ").find_element_by_css_selector("div table tr tr:nth-of-type(14)")
                target2.click()
                sleep(7)
                driver.switch_to.frame("ventana")
                select_element = driver.find_element_by_id("lstCentroGestion")
                select_object = Select(select_element)
                select_object.select_by_value(customer)
                sleep(3)
                driver.find_element_by_id("ctrRangoFechasFECHADESDE").clear()
                driver.find_element_by_id("ctrRangoFechasFECHADESDE").send_keys(self.initialDate)
                select_element = driver.find_element_by_id("lstEstadosLinea")
                select_object = Select(select_element)
                select_object.select_by_value('-1')
                sleep(7)
                driver.find_element_by_id("btnBuscar").click()
                sleep(10)
                try:
                    driver.switch_to.alert.accept()
                except:
                    print("no hay alert")
                    pass
                sleep(15)
                driver.find_element_by_id("btnExcel").click()
                sleep(20)
                driver.close()
            except :
                driver.close()
                self.downloadReports(customer)

    def downloadReportsFacturas(self):
        with webdriver.Chrome(executable_path=self.pathDriver, chrome_options=self.driverOptions) as driver:
            try:
                self.login(driver)
                sleep(5)
                actions = ActionChains(driver)
                target= driver.find_element_by_css_selector("div#ControlMenu1_sec table tr td:nth-of-type(10)")
                actions.move_to_element(target).perform()
                sleep(1)
                target2=target.find_element_by_class_name("SubMenu ").find_element_by_css_selector("div table tr tr:nth-of-type(4)")
                target2.click()
                sleep(5)
                driver.switch_to.frame("ventana")
                select_element = driver.find_element_by_id("lstOrgc")
                select_object = Select(select_element)
                select_object.select_by_value('-1')
                driver.find_element_by_id("RgFechaRecepcion_desde_I").clear()
                driver.find_element_by_id("RgFechaRecepcion_desde_I").send_keys(self.initialDate)
                driver.find_element_by_id("btnVerificaBuscar").click()
                sleep(7)
                driver.find_element_by_id("btnVerificaBuscar").click()
                sleep(5)
                try:
                    driver.switch_to.alert.accept()
                except:
                    print("no hay alert")
                    pass
                sleep(5)
                driver.find_element_by_id("btn").click()
                sleep(8)
                driver.find_element_by_id("correo").clear()
                driver.find_element_by_id("correo").send_keys("fvalenzuela@mkingenieria.cl")
                driver.find_element_by_id("btnDescargarExcelAsync").click()
                sleep(30)
                driver.close()
            except :
                driver.close()
                self.downloadReportsFacturas()

    def downloadNotasCorrecion(self):
        with webdriver.Chrome(executable_path=self.pathDriver, chrome_options=self.driverOptions) as driver:
            try:
                self.login(driver)
                sleep(5)
                actions = ActionChains(driver)
                target= driver.find_element_by_css_selector("div#ControlMenu1_sec table tr td:nth-of-type(10)")
                actions.move_to_element(target).perform()
                sleep(1)
                target2=target.find_element_by_class_name("SubMenu ").find_element_by_css_selector("div table tr tr:nth-of-type(4)")
                target2.click()
                sleep(5)
                driver.switch_to.frame("ventana")
                tab=driver.find_element_by_xpath("//a[@href='control_correcciones.aspx']")
                tab.click()
                select_element = driver.find_element_by_id("lstCentroGestion")
                select_object = Select(select_element)
                select_object.select_by_value('-1')
                driver.find_element_by_id("RgFechaRecepcion_desde_I").clear()
                driver.find_element_by_id("RgFechaRecepcion_desde_I").send_keys(self.initialDate)
                driver.find_element_by_id("btnBuscar").click()
                sleep(15)
                driver.find_element_by_id("btnExcel").click()
                sleep(20)
                driver.close()
            except :
                driver.close()
                self.downloadReportsFacturas()





