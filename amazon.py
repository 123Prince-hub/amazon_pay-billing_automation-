import os
import pytesseract
import xlwings as xw
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC

ws = xw.Book(r'details.xlsx').sheets("data")
rows = ws.range("A2").expand().options(numbers=int).value
driver = webdriver.Chrome(ChromeDriverManager().install()) 
driver.maximize_window()
driver.implicitly_wait(10)

# -------------------------------  function for login  -----------------------------------
def login():   
    email = driver.find_element_by_xpath("//label[contains(text(),'Email or mobile phone number')]//following::input").send_keys(row[0])
    button = driver.find_element_by_xpath("//label[contains(text(),'Email or mobile phone number')]//following::span").click()
    password = driver.find_element_by_xpath("//label[contains(text(),'Password')]//following::input").send_keys(row[1])
    signin = driver.find_element_by_xpath("//label[contains(text(),'Password')]//following::span").click()


# -------------------------------  function for captcha read  -----------------------------------
def captha(): 
    password = driver.find_element_by_xpath("//label[contains(text(),'Password')]//following::input").send_keys(row[1])
    capta= driver.find_element_by_id('auth-captcha-image').screenshot('captcha.png')
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'
    a = pytesseract.image_to_string(r'captcha.png')
    sleep(1)
    a = a.replace(" ", "")
    sleep(1)
    driver.find_element_by_xpath('//label[contains(text(),"Type characters")]//following::input').send_keys(a.lstrip())
    sleep(1)
    driver.find_element_by_id('signInSubmit').click()


# -------------------------------  function for password error  -----------------------------------
def passError(): 
    password = driver.find_element_by_xpath("//label[contains(text(),'Password')]//following::input").send_keys(row[1])
    signin = driver.find_element_by_xpath("//label[contains(text(),'Password')]//following::span").click()


# -------------------------------   function for email error  -----------------------------------
def emailError():  
    button = driver.find_element_by_xpath("//label[contains(text(),'Email or mobile phone number')]//following::span").send_keys(row[0])


# ----------------------  function for all login, password error capthca, email error  ---------------------------
def authentication(): 
    try:
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
    except:
        pass


# -------------------------------   function for otp sending  -----------------------------------
def otp():   
    try:
        send_otp = driver.find_element_by_xpath('').click()
        sleep(10)
        ok_button = driver.find_element_by_xpath('').click()
    except:
        pass

# -------------------------------  function for dissmiss all popups  -----------------------------------
def popups():  
    try:
        pass
    except:
        pass


# -------------------------------  start automation from here  -----------------------------------

num = 2
for row in rows:
    col = ws.range("I"+str(num)).value
    if (col != "Success"):
        try:
            driver.get ("https://www.amazon.in/hfc/bill/electricity?ref_=apay_deskhome_Electricity")

            state_dropdown = driver.find_element_by_xpath('//span[text()="Select State"]').click()
            state = driver.find_element_by_xpath('//a[text()="Rajasthan"]').click()

            board = row[3]
            board_dropdown = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, '//span[text()="Select Electricity Board to proceed"]'))
            )
            board_dropdown.click()
            driver.find_element_by_link_text(board).click()

            k_number = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, '//label[contains(text(), "K Number")]//following::input'))
            )
            k_number.send_keys(row[4])
            fetch_bill = driver.find_element_by_xpath('//span[text()="Fetch Bill"]').click()
            
            element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, '//span[contains(text(),"Continue to Pay")]'))
            )
            
            sleep(1)
            res_amount = element.text
            res_amount = res_amount[17:]
            ws.range("G"+str(num)).value = res_amount

            bypass_id = driver.execute_script('document.querySelector("#paymentBtnId-announce").setAttribute("type", "submit")')
            element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, '//span[contains(text(), "Continue to Pay")]'))
            )
            element.click()

            authentication() # ----------------------login & captcha function calling
            
            button = driver.find_element_by_xpath('//span[text()="Place Order and Pay"]')
            driver.execute_script("arguments[0].click();", button)
            
            
            sleep(30)
            try:
                success = driver.find_element_by_xpath('//*[@id="deep-dtyp-success-alert"]/div/h4')
                pending = driver.find_element_by_xpath('//*[@id="deep-dtyp-pending-widget"]/div/div/h4')
                fail = driver.find_element_by_xpath('//*[@id="deep-dtyp-failed-widget"]/div/div/h4')
                if success.is_displayed():
                    status = success.text
                elif pending.is_displayed():
                    status = pending.text
                else:
                    status = fail.text
            except:
                pass

            try:
                BBPS_Reference_Number = driver.find_element_by_xpath('//*[contains(text(), "BBPS Reference Number")]').text
                bbps_num = BBPS_Reference_Number[23:]
            except:
                pass

            if "successful" in status:
                ws.range("H"+str(num)).value = bbps_num
                ws.range("I"+str(num)).value = "Success"

            elif "pending" in status:
                ws.range("H"+str(num)).value = "NA"
                ws.range("I"+str(num)).value = "Pending"

            else:
                ws.range("H"+str(num)).value = "NA"
                ws.range("I"+str(num)).value = "Fail"
        
        except:
            ws.range("H"+str(num)).value = "NA"
            ws.range("I"+str(num)).value = "Server not response"

    try:
        os.remove('captcha.png')        
    except:
        pass

    num += 1
driver.close()
