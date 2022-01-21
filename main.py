import os
from os import path

from selenium.webdriver.chrome.options import Options
from datetime import date
from scripts import iconstruye

path = path.dirname(path.abspath(__file__))

pathDriver = "/usr/local/bin/chromedriver"
downloadFolder = "/files/"

if os.name == 'nt':
    pathDriver = "./drivers/windows/chromedriver"
    downloadFolder = "\\files\\"

driverOptions = Options()
#driverOptions.add_argument("--headless")
driverOptions.add_argument("--no-sandbox")
driverOptions.add_argument("start-maximized")
driverOptions.add_argument("--window-size=1920,1080")
driverOptions.add_argument("--disable-dev-shm-usage")
driverOptions.add_experimental_option("prefs", {
    "download.default_directory": path + downloadFolder,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "download.safebrowsing.enabled": True,
    'excludeSwitches': ['enable-logging']
})




rpa = iconstruye.botService(driverOptions, path + downloadFolder, pathDriver)
rpa.run()


print("finalizacion")
