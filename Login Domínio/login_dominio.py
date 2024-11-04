import time

from win32com.client import Dispatch
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from time import sleep
from pyautogui import click, locateOnScreen
from os import path, remove

s = Service('V:\Setor Robô\Scripts Python\_comum\Chrome driver\chromedriver.exe')

def find_by_class(iten, driver):
    try:
        elem = driver.find_element(by=By.CLASS_NAME, value=iten)
        return elem
    except:
        return None

def get_chrome_version():
    paths = (
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    )
    
    parser = Dispatch("Scripting.FileSystemObject")
    for file in paths:
        try:
            return parser.GetFileVersion(file)
        except Exception:
            continue
    return None

def download_chrome(version):
    import requests, wget, zipfile
    version = '.'.join(version.split('.')[:-1])
    
    try:
        remove('chromedriver.zip')
    except OSError:
        pass
    
    url = f'https://chromedriver.storage.googleapis.com/LATEST_RELEASE_{version}'
    response = requests.get(url)
    version = response.text
    
    download_url = f"https://chromedriver.storage.googleapis.com/{version}/chromedriver_win32.zip"
    latest_driver_zip = wget.download(download_url, 'V:\Setor Robô\Scripts Python\_comum\Chrome driver\chromedriver.zip')
    
    with zipfile.ZipFile(latest_driver_zip, 'r') as zip_ref:
        zip_ref.extractall('V:\Setor Robô\Scripts Python\_comum\Chrome driver')
    
    remove(latest_driver_zip)

def initialize_chrome(options=None):
    print('>>> Inicializando Chromedriver...')
    
    if not options:
        options = webdriver.ChromeOptions()
        options.add_argument("start-maximized")

    options.add_argument("--ignore-certificate-errors")
    options.add_experimental_option("prefs", {"profile.default_content_setting_values.notifications": 1})
    options.add_experimental_option('excludeSwitches', ['enable-logging'], )
    
    if "chromedriver.exe" in r'V:\Setor Robô\Scripts Python\_comum\Chrome driver\chromedriver.exe':
        return True, webdriver.Chrome(service=s, options=options)
    
    version = get_chrome_version()  # get chromeBrowser version
    if version:
        try:
            download_chrome(version)  # download chromeDriver version that matches
        except PermissionError as e:
            print(str(e))
            return False, _msgs["not_close"]
        except Exception as e:
            print(str(e))
            return False, _msgs["not_download"]
    else:
        return False, _msgs["not_found"]
    
    return True, webdriver.Chrome(service=s, options=options)

def send_input(elem_id, data, driver):
    while True:
        try:
            elem = driver.find_element(by=By.XPATH, value=elem_id)
            elem.send_keys(data)
            break
        except:
            pass

def login():
    options = webdriver.ChromeOptions()
    options.add_argument('--headless=new')
    options.add_argument('--window-size=1920,1080')
    
    status, driver = initialize_chrome()

    driver.get('https://www.dominioweb.com.br/')
    send_input('/html/body/app-root/app-login/div/div/fieldset/div/div/section/form/label[1]/span[2]/input', 'robo@veigaepostal.com.br', driver)
    send_input('/html/body/app-root/app-login/div/div/fieldset/div/div/section/form/label[2]/span[2]/input', 'Rb#0086*', driver)
    driver.find_element(by=By.ID, value='enterButton').click()

    while not locateOnScreen(path.join('imgs', 'abrir_app.png'), confidence=0.9, grayscale=True):
        if find_by_class('trta1-btn-primary', driver):
            driver.find_element(by=By.CLASS_NAME, value='trta1-btn-primary').click()
        sleep(0.5)
    while locateOnScreen(path.join('imgs', 'abrir_app.png'), confidence=0.9, grayscale=True):
        click(locateOnScreen(path.join('imgs', 'abrir_app.png'), confidence=0.9, grayscale=True), button='left')

    while not locateOnScreen(path.join('imgs', 'modulos.png'), confidence=0.9):
        sleep(1)
    driver.quit()

if __name__ == '__main__':
    login()
