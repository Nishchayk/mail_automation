import os
import time
import pyautogui
from zipfile import ZipFile
import wget
import requests
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.firefox import GeckoDriverManager
from webdriver_manager.opera import OperaDriverManager
from webdriver_manager.microsoft import IEDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.firefox.service import Service

class wed_automation():
    # path ---> if you want to attachment in the mail
    # image_path ---> if you want to attachment in meassage body
    def __init__(self,path,image_path):
        self.path = path
        self.image_path = image_path
    def accessing_filepath(self):
        try:
            # validate the path of file and image in existing path
            os.path.exists(self.path)
            os.path.exists(self.image_path)
        except FileNotFoundError as err:
            # if it's not there it through error 
            print("please enter valid path and run again")
            print(err)
    def indetifing_browser(self):   
            # identifing browser in windows or linux 
            # this def indetifing_browser funtion is the process with any denpendence of wed driver
            # what it does not support ther another function with download and unzip the function 
        try:
            x = os.getenv('APPDATA').split("\\")
            self.crop_id = x[2]
            pat = f'C:\\Users\\{self.crop_id}\\AppData\\Local'
            dir1 = ['Mozilla Firefox', 'Google', 'Microsoft', 'Edge', 'Internet Explorer', 'Opera', 'Safari']
            if os.path.exists(f"{pat}\\{dir1[0]}"):
                self.driver = webdriver.Firefox(executable_path=GeckoDriverManager().install())
                self.driver.get('https://outlook.office.com/mail/inbox')
            elif os.path.exists(f"{pat}\\{dir1[1]}"):
                self.driver = webdriver.Chrome(ChromeDriverManager().install())
                self.driver.get('https://outlook.office.com/mail/inbox')
                time.sleep(10)
            elif os.path.exists(f"{pat}\\{dir1[2]}\\{dir1[3]}"):
                a = dir1[3]
                self.driver = webdriver.Edge(executable_path=EdgeChromiumDriverManager().install())
                self.driver.get('https://outlook.office.com/mail/inbox')
            elif os.path.exists(f"{pat}\\{dir1[2]}\\{dir1[4]}"):
                a = dir1[4]
                self.driver = webdriver.Ie(executable_path=IEDriverManager().install())
                self.driver.get('https://outlook.office.com/mail/inbox')
            elif os.path.exists(f"{pat}\\{dir1[5]}"):
                a = dir1[5]
                self.driver = webdriver.Edge(executable_path=OperaDriverManager().install())
                self.driver.get('https://outlook.office.com/mail/inbox')
            elif os.path.exists(f"{pat}\\{dir1[6]}"):
                a = dir1[6]
                print("this will direct run on safari with some changes")
            else:
                print("download any browser")
        except:
            # identifing linux browser 
            pat = "/bin"
            dir1 = ['google-chrome', 'firefox', 'Edge', 'Internet Explorer', 'Opera', 'Safari']
            if os.path.exists(f"{pat}/{dir1[0]}"):
                self.driver = webdriver.Chrome(ChromeDriverManager().install())
                self.driver.get('https://outlook.office.com/mail/inbox')
            elif os.path.exists(f"{pat}/{dir1[1]}"):
                cap = DesiredCapabilities().FIREFOX
                cap["marionette"] = False
                self.driver = webdriver.Firefox(capabilities=cap,service=Service(GeckoDriverManager().install()))
                self.driver.get('https://outlook.office.com/mail/inbox')
            elif os.path.exists(f"{pat}/{dir1[2]}"):
                self.driver = webdriver.Edge(executable_path=EdgeChromiumDriverManager().install())
                self.driver.get('https://outlook.office.com/mail/inbox')
            elif os.path.exists(f"{pat}/{dir1[3]}"):
                self.driver = webdriver.Ie(executable_path=IEDriverManager().install())
                self.driver.get('https://outlook.office.com/mail/inbox')
            elif os.path.exists(f"{pat}/{dir1[4]}"):
                self.driver = webdriver.Opera(executable_path=OperaDriverManager().install())
                self.driver.get('https://outlook.office.com/mail/inbox')
            elif os.path.exists(f"{pat}/{dir1[5]}"):
                print("safari will work with change directly with some change")
            else:
                print("Download browser")

    def getting_version(self):
        self.path = r'C:\Program Files (x86)\Google\Chrome\Application'
        self.files = os.listdir(self.path)
        self.frist_files = self.files[2]
        print(self.frist_files[:9])

    def getting_suitable_version(self):
        url = f'https://chromedriver.storage.googleapis.com/LATEST_RELEASE_{self.frist_files[:9]}'
        responce = requests.get(url)
        self.content = responce.content.decode()
        print(self.content)

    def identifing_the_browser(self):
        # the function identifies browser version and download requried version of it 
        self.exporting_image()
        x = os.getenv('APPDATA').split("\\")
        self.crop_id = x[2]
        pat = f'C:\\Users\\{self.crop_id}\\AppData\\Local'
        dir1 = ['Google', 'Mozilla Firefox', 'Microsoft', 'Edge', 'Internet Explorer', 'Opera', 'Safari']
        if os.path.exists(f"{pat}\\{dir1[0]}"):
            self.a = dir1[0]
            if os.path.exists(os.getcwd() + '\\chromedriver_win32.zip'):
                print("file is present")
            else:
                url1 = f"https://chromedriver.storage.googleapis.com/{self.content}/chromedriver_win32.zip"
                self.file_name = wget.download(url1)
                print(self.a)

        elif os.path.exists(f"{pat}\\{dir1[1]}"):
            self.a = dir1[1]
            url = "https://github.com/mozilla/geckodriver/releases/download/v0.31.0/geckodriver-v0.31.0-win64.zip"
            wget.download(url)

            self.file_name = os.getcwd() + "//geckodriver-v0.31.0-win64.zip"
        elif os.path.exists(f"{pat}\\{dir1[2]}"):
            print("WE have enter the main area")
            if os.path.exists(f"{pat}\\{dir1[2]}\\{dir1[3]}"):
                a = dir1[3]
                url = "https://msedgedriver.azureedge.net/100.0.1185.36/edgedriver_win64.zip"
                wget.download(url)
                self.file_name = os.getcwd() + "//edgedriver_win64.zip"
            elif os.path.exists(f"{pat}\\{dir1[2]}\\{dir1[4]}"):
                a = dir1[4]
                url = "https://github.com/SeleniumHQ/selenium/releases/download/selenium-4.0.0/IEDriverServer_x64_4.0.0.zip"
                wget.download(url)
                self.file_name = os.getcwd() + "//IEDriverServer_x64_4.0.0.zip"
        elif os.path.exists(f"{pat}\\{dir1[5]}"):
            a = dir1[5]
            url = "https://github.com/operasoftware/operachromiumdriver/releases/download/v.99.0.4844.51/operadriver_win64.zip"
            wget.download(url)
            self.file_name = os.getcwd() + "//operadriver_win64.zip"
        elif os.path.exists(f"{pat}\\{dir1[6]}"):
            a = dir1[0]
            url = "https://chromedriver.storage.googleapis.com/100.0.4896.60/Safari"
            wget.download(url)
            self.file_name = os.getcwd() + "//Safari"
        else:
            print("you don't have browser access")

    def unzipping(self):
        # after downloading it will unzip the folder
        with ZipFile(self.file_name, 'r') as zip:
            zip.printdir()
            print("ALL file extracted ")
            zip.extractall()
            print("Done!")
    def required_detail(self):

        try:
            self.driver.maximize_window()
        except:
            self.driver = webdriver.Chrome(os.getcwd() + f"\\{os.path.basename('chromedriver.exe')}")
            self.driver.maximize_window()
            self.driver.get('https://outlook.office.com/mail/inbox')
            print('Going to google')
        self.user_name = "enter your outlook mail id "
        self.user_pass = "enter your outlook password"
        self.To_user = "enter whom you want to send "
        self.Cc = "enter whom you want to send"
    def wed_automation(self):
        time.sleep(3)
        WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, "i0116"))).send_keys(
            self.user_name)
        WebDriverWait(self.driver, 10, (NoSuchElementException, StaleElementReferenceException,)).until(
            EC.presence_of_element_located((By.ID, "idSIButton9"))).click()
        time.sleep(5)
        # Password id = "i0118"
        WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, "i0118"))).send_keys(
            self.user_pass)
        time.sleep(2)
        # next = 'idSIButton9'
        WebDriverWait(self.driver, 10, (NoSuchElementException, StaleElementReferenceException,)).until(
            EC.presence_of_element_located((By.ID, "idSIButton9"))).click()

        try:
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'Pivot29-Tab0')))
            self.driver.quit()
        except:
            WebDriverWait(self.driver, 10, (NoSuchElementException, StaleElementReferenceException,)).until(
                EC.presence_of_element_located((By.ID, "idSIButton9"))).click()
        time.sleep(5)
        try:
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, "id__8"))).click()
            print("1")
        except:
            # driver.findElement(By.xpath("//span[@id='id__6']"))
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'id__6'))).click()
            print("2")

        time.sleep(3)
        WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "ms-BasePicker-input"))).send_keys(self.To_user)

        try:
            #     driver.findElement(By.xpath("//div[@aria-label='Cc']"))
            Cc = self.driver.find_element_by_xpath("//div[@aria-label='Cc']").send_keys(self.Cc)

        except:
            time.sleep(1)
            pyautogui.hotkey("enter")
            pyautogui.hotkey('shift', 'tab')
            time.sleep(1)
            pyautogui.hotkey('shift', 'tab')
            time.sleep(1)
            pyautogui.hotkey('shift', 'tab')
            time.sleep(1)
            pyautogui.hotkey('shift', 'tab')
            time.sleep(2)
            pyautogui.hotkey("enter")
            pyautogui.typewrite(self.Cc)
            time.sleep(2)
        subject_field = self.driver.find_element_by_xpath('//input[contains(@class, "ms-TextField-field")]')
        # enter you subject
        subject_field.send_keys('enter SUbject')
        time.sleep(2)
        message_body = self.driver.find_element_by_xpath("//div[@aria-label='Message body']")
        # enter what you should type in message body
        message_body.send_keys("message body")
        time.sleep(1)
        # this process is used to insert image in message body
        try:
            compus = self.driver.find_element_by_xpath("//button[@id='compose_ellipses_menu']")
            compus.click()
            time.sleep(1)
            image_line = self.driver.find_element_by_xpath("//span[normalize-space()='Insert pictures inline']")
            image_line.click()
            time.sleep(10)
            pyautogui.typewrite(self.path1 + f"\\{self.table_name}")
            time.sleep(10)
            pyautogui.hotkey('enter')
            print("First")
        except:            
            time.sleep(1)
            image = self.driver.find_element_by_xpath("//button[@name='Insert pictures inline']")
            image.click()
            time.sleep(10)
            pyautogui.typewrite(self.path1  + f"\\{self.table_name}")
            time.sleep(10)
            pyautogui.hotkey('enter')
            print("second")
        time.sleep(5)
        # this attachment of the file 
        attach = self.driver.find_element_by_xpath("//button[@name='Attach']")
        attach.click()
        time.sleep(2)
        browser = self.driver.find_element_by_xpath("//button[@name='Browse this computer']")
        browser.click()
        time.sleep(5)
        pyautogui.typewrite(self.path)
        time.sleep(5)
        pyautogui.hotkey('enter')
        time.sleep(4)
        # sending mail
        try:
            time.sleep(6)
            send_button = self.driver.find_element_by_xpath("//button[@title='Send (Ctrl+Enter)']")
            send_button.click()
        except:
            time.sleep(6)
            pyautogui.hotkey('ctrl','enter')
            time.sleep(3)
            pyautogui.hotkey('enter')
            time.sleep(5)
        self.driver.quit()
        
if __name__ == '__main__':
    M = wed_automation('enter file path ','enter image path ')
    try:
        M.getting_version()
        M.getting_suitable_version()
        M.identifing_the_browser()
        if os.path.exists(os.getcwd() + "\\chromedriver.exe"):
            print("file already exists")
            pass
        else:
            M.unzipping()
    except FileNotFoundError:
        M.indetifing_browser()
    M.required_detail()
    M.wed_automation()
    print('Done!')