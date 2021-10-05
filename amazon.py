import pytesseract
import xlwings as xw
from time import sleep
from selenium import webdriver
from selenium.webdriver.support.select import Select
from webdriver_manager.chrome import ChromeDriverManager

ws = xw.Book(r'details.xlsx').sheets("data")
rows = ws.range("A2").expand().options(numbers=int).value
driver = webdriver.Chrome(ChromeDriverManager().install()) 
driver.implicitly_wait(60)

def login():
    driver.get ("https://www.amazon.in/")
    signin = driver.find_element_by_xpath("//span[contains(text(),'Sign in')]").click()
    email = driver.find_element_by_xpath("//label[contains(text(),'Email or mobile phone number')]//following::input").send_keys("bhupsamaiet@gmail.com")
    button = driver.find_element_by_xpath("//label[contains(text(),'Email or mobile phone number')]//following::span").click()
    password = driver.find_element_by_xpath("//label[contains(text(),'Password')]//following::input").send_keys("Lav%210811")
    signin = driver.find_element_by_xpath("//label[contains(text(),'Password')]//following::span").click()
login()        

def captha():
        password = driver.find_element_by_xpath("//label[contains(text(),'Password')]//following::input").send_keys("Lav%210811")
        capta= driver.find_element_by_id('auth-captcha-image').screenshot('captcha.png')
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'
        a = pytesseract.image_to_string(r'captcha.png')
        driver.find_element_by_xpath('//label[contains(text(),"Type characters")]//following::input').send_keys(a.lstrip())

def passError():
    password = driver.find_element_by_xpath("//label[contains(text(),'Password')]//following::input").send_keys("Lav%210811")
    signin = driver.find_element_by_xpath("//label[contains(text(),'Password')]//following::span").click()

def emailError():
    email = driver.find_element_by_xpath("//label[contains(text(),'Email or mobile phone number')]//following::input").send_keys("bhupsamaiet@gmail.com")
    button = driver.find_element_by_xpath("//label[contains(text(),'Email or mobile phone number')]//following::span").click()

try:
    sleep(1)
    capthatest = driver.find_element_by_xpath("//h4[contains(text(),'Enter the characters you see')] | //h4[contains(text(),'There was a problem')] | //h1[contains(text(),'Password assistance')]").text
    sleep(1)
    i = 0
    while((capthatest=="Enter the characters you see") or (capthatest=="There was a problem") or (capthatest=="Password assistance") or (i<16)):
        if capthatest == "Enter the characters you see":
            captha()  
        elif capthatest == "There was a problem":
            passError()
        elif capthatest == "Password assistance":
            emailError()
        else:
            pass
        i+=1
        sleep(1)
except:
    pass