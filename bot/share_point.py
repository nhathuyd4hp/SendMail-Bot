import time
import logging
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
class SharePoint:
    def __init__(
        self,
        username: str,
        password: str,
        timeout: int = 10,
        headless:bool=False,
        logger_name: str = __name__,
    ):
        options = webdriver.ChromeOptions()
        options.add_argument('--disable-notifications')
        # Disable log
        options.add_argument('--disable-logging')
        options.add_argument('--log-level=3')  #
        options.add_argument('--silent')
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        if headless: 
            options.add_argument('--headless=new')
        # Attribute
        self.logger = logging.getLogger(logger_name)
        self.browser = webdriver.Chrome(options=options)
        self.browser.maximize_window()
        self.timeout = timeout
        self.wait = WebDriverWait(self.browser, timeout)
        self.username = username
        self.password = password
        # Trạng thái đăng nhập
        self.authenticated = self.__authentication(username, password)
        
    def __del__(self):
        if hasattr(self,"browser"):
            self.browser.quit()
    
    def __authentication(self, username: str, password: str) -> bool:
        time.sleep(0.5)
        self.browser.get("https://login.microsoftonline.com/")
        # -- Wait load page
        self.wait.until(
            lambda driver: driver.execute_script("return document.readyState") == "complete"
        )
        if self.browser.current_url == "https://m365.cloud.microsoft/?auth=2":
            return True
        try:
            # -- Email, phone or Skype
            self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR,"input[id=\"i0116\"]")
                )
            ).send_keys(username)
            time.sleep(0.5)
            # -- Next
            self.wait.until(
                EC.presence_of_element_located(
                    (By.ID,"idSIButton9")
                )
            ).click()
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
                    (By.CSS_SELECTOR,"input[id=\"i0118\"]")
                )
            ).send_keys(password)
            time.sleep(0.5)
            # -- Sign in
            self.wait.until(
                EC.presence_of_element_located(
                    (By.ID,"idSIButton9")
                )
            ).click()
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
            try:
                self.wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR,"div[class='row text-title']")
                    )
                )
                self.wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR,"input[id='idSIButton9']")
                    )
                ).click()
                time.sleep(0.5)
                self.wait.until(
                    lambda driver: driver.execute_script("return document.readyState") == "complete"
                )
                if self.browser.current_url.startswith("https://m365.cloud.microsoft/?auth="):
                    self.logger.info("Xác thực thành công")
                    return True
                else:
                    return False
            except TimeoutException as e:
                self.logger.error(e)
                return False
        except TimeoutException as e:
            self.logger.error(e)
            return False
        except Exception as e:
            self.logger.error(e)
            return False