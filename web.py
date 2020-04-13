



from selenium import webdriver
from selenium.webdriver.chrome.options import Options

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
#from selenium.common.exceptions import timeoutException



def open(chrome_driver, website_address): # subject to further improvement
    ''' [>] Opens specified website (WebsiteAddress) with Chrome Browser.
    [>] ChromeDriver - path to the Chrome driver that is required for the version of Chrome that is utilized.
    '''
    options = Options()
    options.add_argument("--disable-extensions")
    options.add_argument("--start-maximized")
    options.add_experimental_option("useAutomationExtension", False)
    prefs = {'download.prompt_for_download': True,}
    options.add_experimental_option('prefs', prefs)
    #ChooseDriver
    browser = webdriver.Chrome(chrome_driver ,options=options)
    #Open Website
    browser.get(website_address)