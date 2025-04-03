import os
import re
import time
import logging
from urllib.parse import urljoin
from selenium import webdriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException,ElementNotInteractableException,ElementClickInterceptedException
from common import HandleException

class SharePoint:
    def __init__(
        self,
        username: str,
        password: str,
        url:str,
        timeout: int = 10,
        headless:bool=False,
        download_directory:str = os.path.dirname(os.path.abspath(__file__)),
        logger_name: str = __name__,
    ):
        os.makedirs(download_directory,exist_ok=True)
        options = webdriver.ChromeOptions()
        options.add_argument('--disable-notifications')
        options.add_experimental_option("prefs", {
            "download.default_directory": download_directory,  
            "download.prompt_for_download": False, 
            "safebrowsing.enabled": True 
        })
        # Disable log
        options.add_argument('--disable-logging')
        options.add_argument('--log-level=3')  
        options.add_argument('--silent')
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        if headless: 
            options.add_argument('--headless=new')
        # Attribute
        self.url = url
        self.logger = logging.getLogger(logger_name)
        self.browser = webdriver.Chrome(options=options)
        self.browser.maximize_window()
        self.timeout = timeout
        self.wait = WebDriverWait(self.browser, timeout)
        self.username = username
        self.download_directory = download_directory
        self.password = password
        # Trạng thái đăng nhập
        self.authenticated = self.__authentication(username, password)
       
    @HandleException
    def __readyState(self) -> bool:
        return self.browser.execute_script("return document.readyState") == "complete"
    
    @HandleException
    def __wait_download_finish(self) -> bool:
        self.browser.get("chrome://downloads/")
        download_items:list[WebElement] = self.browser.execute_script("""
            return document.
            querySelector("downloads-manager").shadowRoot
            .querySelector("#mainContainer")
            .querySelector("#downloadsList")
            .querySelector("#list")
            .querySelectorAll("downloads-item")
        """)
        for item in download_items:
            quick_show_in_folder = self.browser.execute_script(f"""
                return document
                .querySelector("downloads-manager").shadowRoot
                .querySelector("#downloadsList")
                .querySelector("#list")
                .querySelector("#{item.get_attribute("id")}").shadowRoot
                .querySelector("#quick-show-in-folder")
            """)
            if not isinstance(quick_show_in_folder,WebElement):
                time.sleep(5)
                return False
        return True
    
    @HandleException
    def __authentication(self, username: str, password: str) -> bool:
        time.sleep(0.5)
        self.browser.get("https://login.microsoftonline.com/")
        # -- Wait load page
        while not self.__readyState():
            continue
        time.sleep(2)
        if self.browser.current_url.startswith("https://m365.cloud.microsoft/?auth="):
            self.logger.info("✅ Xác thực thành công!")
            return True
        
        # -- Email, phone or Skype
        self.wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR,"input[type=\"email\"]")
            )
        ).send_keys(username)
        time.sleep(0.5)
        # -- Next
        btn = self.wait.until(
            EC.presence_of_element_located(
                (By.ID,"idSIButton9")
            )
        )
        self.wait.until(EC.element_to_be_clickable(btn)).click()
        time.sleep(1)
        # -- Check usernameError
        try:
            usernameError = self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR,"div[id=\"usernameError\"]")
                )
            )
            self.logger.error(usernameError.text)
            return False
        except TimeoutException:
            pass
        # -- Password
        self.wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR,"input[type=\"password\"]")
            )
        ).send_keys(password)
        time.sleep(0.5)
        # -- Sign in
        btn = self.wait.until(
            EC.presence_of_element_located(
                (By.ID,"idSIButton9")
            )
        )
        self.wait.until(EC.element_to_be_clickable(btn)).click()
        time.sleep(0.5)
        # -- Check stay signed in
        try:
            passwordError = self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR,"div[id=\"passwordError\"]")
                )
            )
            self.logger.error(passwordError.text)
            return False
        except TimeoutException:
            pass
        # -- Stay signed in?
        self.wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR,"div[class='row text-title']")
            )
        )
        btn = self.wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR,"input[id='idSIButton9']")
            )
        )
        self.wait.until(EC.element_to_be_clickable(btn)).click()
        time.sleep(1)
        
        while not self.__readyState():
            continue
        
        if not self.browser.current_url.startswith("https://m365.cloud.microsoft/?auth="):
            self.logger.info("❌  Xác thực thất bại!")
            return False
        
        self.logger.info("✅ Xác thực thành công!")
        return True
        
    @HandleException
    def get_files_in_folder(self,site_folder:str,folder:str) -> tuple[str,list[str]]:
        """ Lấy thông tin folder
        Returns:
            tuple[str,list[str]]: thông tin folder
            str: folder_url
            list[str]: danh sách các file trong folder
        """
        time.sleep(1)
        self.browser.get(urljoin(self.url, site_folder))
        while not self.__readyState():
            continue
        # ----- #
        self.wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR,"input[type='search']"))
        )
        input = self.wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR,"input[type='search']"))
        )
        input.click()
        time.sleep(1)
        input.send_keys(Keys.CONTROL + "a") 
        input.send_keys(Keys.BACK_SPACE)    
        input.send_keys(folder)
        input.send_keys(Keys.ENTER)
        while not self.__readyState():
            continue
        time.sleep(1)
        files = []
        try:
            ms_DetailsList_contentWrapper = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR,"div[class='ms-DetailsList-contentWrapper']"))
            )
            rows = ms_DetailsList_contentWrapper.find_elements(
                by = By.CSS_SELECTOR,
                value = "div[class^='ms-DetailsRow-fields fields-']"
            )
            for row in rows:
                name_gridcell = row.find_element(
                    By.CSS_SELECTOR, 
                    "div[role='gridcell'][data-automation-key^='displayNameColumn_']"
                )
                button = name_gridcell.find_element(By.TAG_NAME,'button')
                files.append(button.text)
        except TimeoutException:
            rows = self.browser.find_elements(
                By.CSS_SELECTOR,
                "div[id^='virtualized-list_'][id*='_page-0_']"
            )
            for row in rows:
                name_gridcell = row.find_element(
                    By.CSS_SELECTOR, 
                    "div[role='gridcell'][data-automationid='field-LinkFilename']"
                )
                span = name_gridcell.find_element(By.TAG_NAME,'span')
                files.append(span.text)
        finally:
            return self.browser.current_url,files
        
    @HandleException
    def download_file(self,site_url:str,file_pattern:str) -> tuple[bool,list]:
        """ Downloaf File From Site 
        Returns:
            tuple[bool,list]: Result, Note
        """
        download_file = []
        time.sleep(0.5)
        if not self.browser.current_url == site_url:
            self.browser.get(site_url)
        while not self.__readyState():
            continue
        # Access Denied
        try:
            ms_error_header = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div#ms-error-header h1"))
            )
            if ms_error_header.text == 'Access Denied':
                SignInWithTheAccount = self.wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div#ms-error a"))
                )  
                SignInWithTheAccount.click()
            else:
                return (False,None)
        except TimeoutException:
            pass
        # -- Folder --
        folders = file_pattern.split("/")[:-1]
        for folder in folders:
            # Lấy tất cả các dòng
            rows = []
            try:
                ms_DetailsList_contentWrapper = self.wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR,"div[class='ms-DetailsList-contentWrapper']"))
                )
                rows = ms_DetailsList_contentWrapper.find_elements(
                    by = By.CSS_SELECTOR,
                    value = "div[class^='ms-DetailsRow-fields fields-']"
                )
                for row in rows:
                    icon_gridcell = row.find_element(
                        By.CSS_SELECTOR, 
                        "div[role='gridcell'][data-automationid='DetailsRowCell']"
                    )
                    if icon_gridcell.find_elements(By.TAG_NAME, "svg"): # Folder
                        name_gridcell = row.find_element(
                            By.CSS_SELECTOR, 
                            "div[role='gridcell'][data-automation-key^='displayNameColumn_']"
                        )
                        button = name_gridcell.find_element(By.TAG_NAME,'button')
                        if button.text == folder:
                            button.click()
                            time.sleep(5) # Có thể tối ưu ở đây
                            break
            except TimeoutException:
                rows = self.browser.find_elements(
                    By.CSS_SELECTOR,
                    "div[id^='virtualized-list_'][id*='_page-0_']"
                )
                for row in rows:
                    icon_gridcell = row.find_element(
                        By.CSS_SELECTOR, 
                        "div[role='gridcell'][data-automationid='field-DocIcon']"
                    )
                    name_gridcell = row.find_element(
                        By.CSS_SELECTOR, 
                        "div[role='gridcell'][data-automationid='field-LinkFilename']"
                    )
                    if icon_gridcell.find_elements(By.TAG_NAME, "svg"): # Folder
                        span = name_gridcell.find_element(By.TAG_NAME,'span')
                        if span.text == folder:
                            span.click()
                            time.sleep(5)
                            break
        # -- File --
        pattern = file_pattern.split("/")[-1]
        pattern = re.compile(pattern)
        try:
            ms_DetailsList_contentWrapper = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR,"div[class='ms-DetailsList-contentWrapper']"))
            )       
            rows = ms_DetailsList_contentWrapper.find_elements(
                by = By.CSS_SELECTOR,
                value = "div[class^='ms-DetailsRow-fields fields-']"
            )
            # Lấy tất cả các dòng    
            if not rows:            
                self.logger.info("Không tìm thấy file phù hợp")
                return (False,"Không tìm thấy file")
            for row in rows:
                icon_gridcell = row.find_element(
                    By.CSS_SELECTOR, 
                    "div[role='gridcell'][data-automationid='DetailsRowCell']"
                )
                if icon_gridcell.find_elements(By.TAG_NAME, "svg"): # Folder
                    continue
                else:
                    name_gridcell = row.find_element(
                        By.CSS_SELECTOR, 
                        "div[role='gridcell'][data-automation-key^='displayNameColumn_']"
                    )
                    button = name_gridcell.find_element(By.TAG_NAME,'button')
                    display_name = button.text
                    # Nếu display_name match với file là được
                    if pattern.match(display_name):
                        gridcell_div = row.find_element(By.XPATH, "./preceding-sibling::div[@role='gridcell']")
                        checkbox = gridcell_div.find_element(By.CSS_SELECTOR, "div[role='checkbox']")
                        if checkbox.get_attribute('aria-checked') == "false": 
                            time.sleep(1)
                            download_file.append(os.path.join(self.download_directory,display_name))
                            checkbox.click()                  
        except TimeoutException:
            rows = self.browser.find_elements(
                By.CSS_SELECTOR,
                "div[id^='virtualized-list_'][id*='_page-0_']"
            )
            for row in rows:
                icon_gridcell = row.find_element(
                    By.CSS_SELECTOR, 
                    "div[role='gridcell'][data-automationid='field-DocIcon']"
                )
                if icon_gridcell.find_elements(By.TAG_NAME, "svg"): # Folder
                    continue
                else:
                    name_gridcell = row.find_element(
                        By.CSS_SELECTOR, 
                        "div[role='gridcell'][data-automationid='field-LinkFilename']"
                    )
                    span_name_gridcell = name_gridcell.find_element(By.TAG_NAME,'span')
                    if pattern.match(span_name_gridcell.text):
                        if row.find_elements(By.CSS_SELECTOR,"div[class^='rowSelectionCell_']"):
                            time.sleep(1)
                            download_file.append(os.path.join(self.download_directory,span_name_gridcell.text))
                            row.find_element(By.CSS_SELECTOR,"div[class^='rowSelectionCell_']").click()
        # Download
        time.sleep(2)
        try:
            self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//span[text()='Download']"))
            )
            self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//span[text()='Download']"))
            ).click()
        except TimeoutException:
            self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR,"i[data-icon-name='More']"))
            )
            self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR,"i[data-icon-name='More']"))
            ).click()
            self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "button[name='Download']"))
            )
            self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "button[name='Download']"))
            ).click()
        time.sleep(2)
        self.logger.info(f"Tải file {site_url} thành công")
        while not self.__wait_download_finish():
            continue
        time.sleep(5)
        return (True,download_file)
            

            
                
                
                
            