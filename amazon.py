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
driver.implicitly_wait(80)

num = 2

for row in rows:
    col = ws.range("I"+str(num)).value
    if (col != "Success"):
        try:
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
            brd = row[3]
            board_dropdown = driver.find_element_by_xpath('//span[text()="Select Electricity Board to proceed"]').click()
            driver.find_element_by_link_text(brd).click()

            k_number = driver.find_element_by_xpath('//label[contains(text(), "K Number")]//following::input').send_keys(row[4])
            fetch_bill = driver.find_element_by_xpath('//span[text()="Fetch Bill"]').click()
            
            element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, '//span[contains(text(),"Continue to Pay")]'))
            )
            
            sleep(2)
            res_amount = element.text
            res_amount = res_amount[17:]
            ws.range("G"+str(num)).value = res_amount

            bypass_id = driver.execute_script('document.querySelector("#paymentBtnId-announce").setAttribute("type", "submit")')
            element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, '//span[contains(text(), "Continue to Pay")]'))
            )
            element.click()

            try:
                func1()
            except:
                pass
            
            button = driver.find_element_by_xpath('//span[text()="Place Order and Pay"]')
            driver.execute_script("arguments[0].click();", button)
                    
            try:
                status = driver.find_element_by_xpath('//h4[text()="Your bill payment has failed"] | //h4[text()="Your bill payment is successful"] | //h4[text()="Your bill payment is pending"]')
                if (status.is_displayed()) or (status.is_enabled()):
                    status = status.text
                    print(status)
            except:
                print("no status")

            try:
                BBPS_Reference_Number = driver.find_element_by_xpath('//*[contains(text(), "BBPS Reference Number")]').text
                bbps_num = BBPS_Reference_Number[23:]
            except:
                pass

            if "successful" in status:
                print("successful")
                ws.range("H"+str(num)).value = bbps_num
                ws.range("I"+str(num)).value = "Success"

            elif "pending" in status:
                print("pending")
                ws.range("H"+str(num)).value = "NA"
                ws.range("I"+str(num)).value = "Pending"

            elif "failed" in status:
                print("failure")
                ws.range("H"+str(num)).value = "NA"
                ws.range("I"+str(num)).value = "Failure"
            else:
                pass
            sleep(1000)
        
        except:
            ws.range("H"+str(num)).value = "NA"
            ws.range("I"+str(num)).value = "Server not response"
        
    num += 1