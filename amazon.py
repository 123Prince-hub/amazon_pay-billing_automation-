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

num = 2

for row in rows:
    col = ws.range("H"+str(num)).value
    if (col != "Success"):
      
        def login():
            email = driver.find_element_by_xpath("//label[contains(text(),'Email or mobile phone number')]//following::input").send_keys(row[0])
            button = driver.find_element_by_xpath("//label[contains(text(),'Email or mobile phone number')]//following::span").click()
            password = driver.find_element_by_xpath("//label[contains(text(),'Password')]//following::input").send_keys(row[1])
            signin = driver.find_element_by_xpath("//label[contains(text(),'Password')]//following::span").click()

        def captha():
            password = driver.find_element_by_xpath("//label[contains(text(),'Password')]//following::input").send_keys(row[1])
            capta= driver.find_element_by_id('auth-captcha-image').screenshot('captcha.png')
            pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'
            a = pytesseract.image_to_string(r'captcha.png')
            driver.find_element_by_xpath('//label[contains(text(),"Type characters")]//following::input').send_keys(a.lstrip())
            driver.find_element_by_id('signInSubmit').click()

        def passError():
            password = driver.find_element_by_xpath("//label[contains(text(),'Password')]//following::input").send_keys(row[1])
            signin = driver.find_element_by_xpath("//label[contains(text(),'Password')]//following::span").click()

        def emailError():
            button = driver.find_element_by_xpath("//label[contains(text(),'Email or mobile phone number')]//following::span").send_keys(row[0])


        
        def func1():
            sleep(1)
            i = 0
            capthatest = driver.find_element_by_xpath("//label[contains(text(),'Email or mobile phone number')] | //h4[contains(text(),'Enter the characters you see')] | //h4[contains(text(),'There was a problem')] | //h1[contains(text(),'Password assistance')]").text
            while((capthatest=="Enter the characters you see") or (capthatest=="Email or mobile phone number") or (capthatest=="There was a problem") or (capthatest=="Password assistance") and (i<16)):
                capthatest = driver.find_element_by_xpath("//label[contains(text(),'Email or mobile phone number')] |//h4[contains(text(),'Enter the characters you see')] | //h4[contains(text(),'There was a problem')] | //h1[contains(text(),'Password assistance')]").text

                if capthatest == "Enter the characters you see":
                    captha()  
                elif capthatest == "There was a problem":
                    passError()
                elif capthatest == "Password assistance":
                    emailError()
                elif capthatest == "Email or mobile phone number":
                    login()
                else:
                    pass
                i+=1
                sleep(1)


        driver.get ("https://www.amazon.in/hfc/bill/electricity?ref_=apay_deskhome_Electricity")

        state_dropdown = driver.find_element_by_xpath('//span[text()="Select State"]').click()
        state = driver.find_element_by_xpath('//a[text()="Rajasthan"]').click()

        board_dropdown = driver.find_element_by_xpath('//span[text()="Select Electricity Board to proceed"]').click()
        board = driver.find_element_by_xpath('//a[text()="Jaipur Vidyut Vitran Nigam (JVVNL)"]').click()

        k_number = driver.find_element_by_xpath('//input[@placeholder="Please enter your K Number"]').send_keys(row[4])
        fetch_bill = driver.find_element_by_xpath('//span[text()="Fetch Bill"]').click()
        sleep(5)
        bypass_id = driver.execute_script('document.querySelector("#paymentBtnId-announce").setAttribute("type", "submit")')
        continue_pay = driver.find_element_by_xpath('//span[contains(text(), "Continue to Pay")]').click()

        try:
            func1()
        except:
            pass
        
        button = driver.find_element_by_xpath('//span[text()="Place Order and Pay"]')
        driver.execute_script("arguments[0].click();", button)
    
        sleep(120)
        BBPS_Reference_Number = driver.find_element_by_xpath('//*[contains(text(), "BBPS Reference Number")]').text
        ws.range("G"+str(num)).value = BBPS_Reference_Number
        ws.range("H"+str(num)).value = "Success"
    
    num += 1