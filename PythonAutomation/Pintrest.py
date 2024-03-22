from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import smtplib
import zipfile
import os
import time

chromedriver_path = r"chromedriver-win64/chromedriver.exe"
chrome_binary_path = r"chrome-win64\chrome.exe"
chrome_options = webdriver.ChromeOptions()
p ={"download.default_directory": r"C:\Users\018055\Documents\Sify\Automation_testing\Selenium\Pintrest\Flowers"}
chrome_options.add_experimental_option("prefs", p)
chrome_options.binary_location = chrome_binary_path
service = Service(executable_path=chromedriver_path)

driver=webdriver.Chrome(service=service,options=chrome_options)

def search_image(img_name):
    search = driver.find_element(By.XPATH,"//input[@placeholder='Search']")
    search.click()
    search.send_keys(img_name)
    search.send_keys(Keys.ENTER)
    time.sleep(5)
    #driver.find_element(By.XPATH,"(//div[@class='KS5 hs0 un8 C9i TB_'])[12]").click()
    first_image = driver.find_element("css selector", "div[data-test-id='pinWrapper']:first-child")
    first_image.click()
    time.sleep(5)
    print("Got your result!!!!!!!!")
    driver.find_element(By.XPATH,"//button[@aria-label='More options']//div[@class='x8f INd _O1 KS5 mQ8 OGJ YbY']").click()
    time.sleep(2)
    driver.find_element(By.XPATH,"(//span[@title='Download image'])[1]").click()
    time.sleep(3)
    print(img_name," downloaded sucessfully!!!!!!")
    time.sleep(5)
    driver.find_element(By.XPATH,"//*[@id='gradient']/div/div/div[2]/div/div/div/div/div/div/div/div/div/div[2]/div[1]/div[1]/div/div/div/div[1]/div[1]/div[2]/div/div/div/div/div/div/div/div/div/button/div").click()
    time.sleep(2)
    embed_code_option = driver.find_element(By.XPATH,"(//span[@title='Get Pin embed code'])[1]")
    embed_code_option.click()
    embed_code = driver.find_element(By.XPATH,"/html/body/div[5]/div/div/div/div[2]/div/div[2]/div[2]")
    s=embed_code.text
    driver.find_element(By.XPATH,"(//div[@class='RCK Hsu USg adn CCY gn8 L4E kVc S9z DZT I56 Zr3 C9q a_A gpV hNT BG7 hDj _O1 KS5 mQ8 Tbt L4E'])[1]").click()
    print(img_name," code copied successfully!!!!!!!")
    return s

def html_file(l):
    f=open(r"Pintrest\flowers.htm","w")
    s="<html><body><h1></h1></body></html>"
    w=""
    for i in l:
        w+=i
    print(w)
    result = s[:16] + w + s[15:]
    print(result)
    f.write(result)
    f.close

def zip_file():
    base_folder = "Pintrest/Flowers"
    base_folder_1 = "Pintrest"


    files_in_folder = [os.path.join(root, file) for root, dirs, files in os.walk(base_folder) for file in files]
    list_files = files_in_folder + [os.path.join(base_folder_1, "flowers.htm")]

    with zipfile.ZipFile('final.zip', 'w') as zipF:
        for file in list_files:
            arcname = os.path.relpath(file, base_folder_1)
            zipF.write(file, arcname=arcname, compress_type=zipfile.ZIP_DEFLATED)

    print('The files have been compressed')

def mail_sent():
    sender_mail = 'sender mail'
    receivers_mail = ['receiver's mail id']
    subject = "ZIP file"
    message = "Hi!!! I'm sending zipfile to you....."

    msg=MIMEMultipart()
    msg['Subject'] = subject
    body_msg=MIMEText(message,"plain")
    msg.attach(body_msg)

    zip_file_path = r"C:\Users\018055\Documents\Sify\Automation_testing\Selenium\final.zip"
    with open(zip_file_path,'rb') as file:
        msg.attach(MIMEApplication(file.read(), Name='final.zip'))
    
    password="your password"
    smtpObj = smtplib.SMTP('smtp.gmail.com', 587) 
    smtpObj.starttls()  
    smtpObj.login(sender_mail, password)
    smtpObj.sendmail(sender_mail, receivers_mail, msg.as_string())
    smtpObj.quit()

    print("Successfully sent email!!!!!!!")

try:
    driver.get("https://www.pinterest.com/login/")
    driver.maximize_window()
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "id")))
    for _ in range(10):
        try:
            google_button = driver.find_element(By.XPATH, "//span[text()='Continue with Google']")
            google_button.click()
            break
        except:
            print("Error ocurred!!!")
            driver.refresh()
    WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
    windows = driver.window_handles
    driver.switch_to.window(windows[1])

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='email']"))
    )

    driver.find_element(By.CSS_SELECTOR, "input[type='email']").send_keys("your mail id")
    time.sleep(3)
    driver.find_element(By.ID, "identifierNext").click()
    time.sleep(3)
    driver.find_element(By.CSS_SELECTOR, "input[type='password']").send_keys("your password")
    time.sleep(3)
    driver.find_element(By.ID, "passwordNext").click()
    time.sleep(3)
    driver.switch_to.window(windows[0])
    WebDriverWait(driver, 10).until(EC.title_contains("Pinterest"))
    time.sleep(5)
    print("Successfully login!!!!!!")
    img_name=['Roses','Tulips','Lilies','Sunflowers','Gerbera Daisies']
    l=[]
    for i in img_name:
        s = search_image(i)
        l.append(s)
    print(l)
    html_file(l)
    print("HTML code is ready!!!!!!!!")
    zip_file()
    mail_sent()
    print("Completed successfully!!!!!!")


finally:
    driver.quit()
