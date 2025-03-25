import time
import logging
from typing import Union
import pandas as pd
from functools import wraps
from selenium import webdriver
from selenium.common.exceptions import TimeoutException,NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait

def login_required(func):
    @wraps(func)
    def wrapper(self:"MailDealer", *args, **kwargs):
        if not self.authenticated:
            self.logger.error('❌ Yêu cầu xác thực')
            return []
        return func(self, *args, **kwargs)
    return wrapper

def switch_to_default_content(func):
    @wraps(func)
    def wrapper(self:"MailDealer", *args, **kwargs):
        self.browser.switch_to.default_content()
        result = func(self, *args, **kwargs)
        return result
    return wrapper

class MailDealer:
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
        options.add_argument('--disable-logging')
        options.add_argument('--log-level=3')  #
        options.add_argument('--silent')
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        if headless: 
            options.add_argument('--headless=new')

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
        
    @switch_to_default_content
    def __authentication(self, username: str, password: str) -> bool:
        time.sleep(0.5)
        self.browser.get('https://md29.maildealer.jp/')
        try:
            username_field = self.wait.until(
                EC.presence_of_element_located((By.ID, 'fUName')),
            )
            username_field.send_keys(username)

            password_field = self.wait.until(
                EC.presence_of_element_located((By.ID, 'fPassword')),
            )
            password_field.send_keys(password)

            self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "input[value='ログイン']"),
                ),
            )
            login_btn = self.wait.until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "input[value='ログイン']"),
                ),
            )
            login_btn.click()
            try:
                self.wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "div[class='d_error_area ']"),
                    ),
                )
                self.logger.error(
                    '❌ Xác thực thất bại! Kiểm tra thông tin đăng nhập.',
                )
                return False
            except TimeoutException:
                if self.browser.current_url.find("app") != -1:
                    self.logger.info('✅ Xác thục thành công!')
                    return True
                return False
        except Exception as e:
            self.logger.error(f'❌ Xác thực thất bại! {e}.')
            return False


    @login_required
    @switch_to_default_content
    def __open_mail_box(self, mail_box: str, tab: Union[str,None] = None) -> bool:
        # --------------#
        if not self.wait.until(
            EC.frame_to_be_available_and_switch_to_it(
                (By.CSS_SELECTOR, "iframe[id='ifmSide']"),
            ),
        ):
            self.logger.error('Không tìm thấy Frame MailBox')
            return False
        mail_boxs: list[str] = mail_box.split('/')
        for box in mail_boxs:
            try:
                span_box = self.wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, f"span[title='{box}']"),
                    ),
                )
                span_box.click()
                time.sleep(1)
            except TimeoutException:
                self.logger.error(f'Không tìm thấy hộp thư {box}')
                return False
        self.browser.switch_to.default_content()
        # --------------#
        if not self.wait.until(
            EC.frame_to_be_available_and_switch_to_it(
                (By.CSS_SELECTOR, "iframe[id='ifmMain']"),
            ),
        ):
            self.logger.error('Không thể tìm thấy mailbox')
            return False
        # --------------#
        if tab:
            try:
                self.wait.until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            f".//span[@class='olv-c-tab__name' and text()='{tab}']",
                        ),
                    ),
                ).click()
            except TimeoutException:
                self.logger.error(f'❌ Không tìm thấy Tab {tab}')
                return False
        self.browser.switch_to.default_content()
        return True

    @login_required
    @switch_to_default_content
    def mailbox(self, mail_box: str, tab_name: Union[str, None] = None) -> pd.DataFrame | None:
        try:
            self.__open_mail_box(mail_box, tab_name)
            time.sleep(2)
            if not self.wait.until(
                EC.frame_to_be_available_and_switch_to_it(
                    (By.CSS_SELECTOR, "iframe[id='ifmMain']"),
                ),
            ):
                self.logger.error(f'❌ Không thể tìm thấy Content Iframe!.')
                return None
            thead = self.wait.until(
                EC.presence_of_element_located(
                    (By.TAG_NAME,'thead')
                )
            )
            labels = thead.find_elements(By.TAG_NAME,'th')
            # Lọc lấy các thẻ label
            columns = []
            index_value = []
            for index,label in enumerate(labels):
                if label.find_elements(By.XPATH, "./*") and label.text:
                    columns.append(label.text)
                    index_value.append(index)
            df = pd.DataFrame(columns=columns)  
            try:
                self.wait.until(
                    EC.presence_of_element_located(
                        (By.XPATH,"//div[text()='条件に一致するデータがありません。']")
                    )
                )
                self.logger.info(f'✅ Hộp thư: {mail_box} rỗng')
                return df
            except:             
                tbodys = self.wait.until(
                    EC.presence_of_all_elements_located(
                        (By.TAG_NAME,'tbody')
                    )
                )
                for tbody in tbodys:
                    row = []
                    values:list[WebElement] = tbody.find_elements(By.TAG_NAME,"td")
                    values:list[WebElement] = [values[index] for index in index_value]
                    for value in values:
                        row.append(value.text)
                    df.loc[len(df)] = row                   
                self.logger.info(f'✅ Lấy hộp thư: {mail_box}, tab: {tab_name}: thành công')
                return df
        except Exception as e:    
            self.logger.error(f'❌ Không thể lấy được danh sách mail: {mail_box}, tab: {tab_name}: {e}')
            return None
            
    
    @login_required
    def read_mail(self, mail_box: str,mail_id:str,tab_name:str=None) -> str:
        try:
            content = ""
            if not self.browser.current_url.startswith('https://md29.maildealer.jp/app/'):
                self.__authentication(self.username, self.password)
            self.__open_mail_box(
                mail_box=mail_box,
                tab=tab_name,
            )
            self.wait.until(
                EC.frame_to_be_available_and_switch_to_it(
                    (By.CSS_SELECTOR, "iframe[id='ifmMain']"),
                ),
            )
            email_span = self.wait.until(
                EC.presence_of_element_located((By.XPATH,f"//span[text()='{mail_id}']"))
            )
            email_span.click()
            try:
                self.wait.until(
                    EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR,"iframe[id='html-mail-body-if']"))
                )
                ps = self.wait.until(
                    EC.presence_of_all_elements_located((By.TAG_NAME,'p'))
                )
                for p in ps:
                    content += p.text + "\n"
            except TimeoutException:
                body = self.wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR,"div[class='olv-p-mail-view-body']")
                    )
                )
                content = body.find_element(By.TAG_NAME,'pre').text
            self.logger.info(f'✅ Đã đọc được nội dung mail:{mail_id}. tab: {tab_name} ở box:{mail_box}')
            return content
        except Exception as e:
            self.logger.error(f'❌ Dọc nội dung mail:{mail_id} ở {mail_box} thất bại: {e}')
    
    @login_required
    def 一括操作(self,案件ID:any,このメールと同じ親番号のメールをすべて関連付ける:bool=False) -> tuple[bool,str]:
        try:
            popup = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR,"div[class='pop-panel__content']"))
            )
            input = popup.find_element(By.ID,'fMatterID_add')
            button = input.find_element(By.XPATH, "./ancestor::*[1]//button")
            このメールと同じ親番号のメールをすべて関連付ける_div = popup.find_element(
                By.XPATH,f"//div[text()='このメールと同じ親番号のメールをすべて関連付ける']"
            )
            div_checkbox = このメールと同じ親番号のメールをすべて関連付ける_div.find_element(By.XPATH, "./ancestor::*[1]//div")
            div_input = このメールと同じ親番号のメールをすべて関連付ける_div.find_element(By.XPATH, "./ancestor::*[1]//input")
            time.sleep(0.5)
            if div_input.is_selected() != このメールと同じ親番号のメールをすべて関連付ける:
                time.sleep(0.5)
                div_checkbox.click()
            time.sleep(1)
            input.clear()
            input.send_keys(案件ID)       
            time.sleep(1)
            button.click()
            time.sleep(2)
            # Check Result
            snackbar_div = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR,"div[class='snackbar__msg']"))
            )
            if snackbar_div.text == "案件との関連付けを行いました。":
                self.logger.info(f"Liên kết {案件ID}: {snackbar_div.text}")
                return True,snackbar_div.text
            else:
                self.logger.info(f"Liên kết {案件ID}: {snackbar_div.text}")
                return False,snackbar_div.text
            
        except TimeoutException as e:
            button = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR,"button[title='一括操作']"))
            )
            button = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button[title='一括操作']"))   
            )
            button.click()
            return self.一括操作(
                案件ID=案件ID,
                このメールと同じ親番号のメールをすべて関連付ける=このメールと同じ親番号のメールをすべて関連付ける,
            )
        except NoSuchElementException as e:
            self.logger.error(f'❌ Liên kết {案件ID} thất bại: {e}')
            return False,e
        except Exception as e:
            self.logger.error(f'❌ Liên kết {案件ID} thất bại: {e}')
            return False,e
        
    def reply(self,mail_id:str) -> bool:
        # Search
        search_input = self.wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR,"input[placeholder=\"このメールボックスのメール・電話を検索\"]")
            )
        )
        search_input.clear()
        search_input.send_keys(mail_id)
        # search
        self.wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR,"button[title=\"検索\"]")
            )
        )
        search_btn = self.wait.until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR,"button[title=\"検索\"]")
            )
        )
        search_btn.click()
        # Switch frame
        ifame = self.wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR,"iframe[id=\"ifmMain\"]")
            )
        )
        self.browser.switch_to.frame(ifame)
        try:
            mainApp = self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR,"div[id='mainApp'")
                )
            )
            form = mainApp.find_element(By.CSS_SELECTOR,"form[id='form-olv-p-viewmail']")
            
            # Check Status Reply "このメールには既に返信されています。" -> True
        except TimeoutException:
            return False
        except Exception as e:
            return False
             
          
        


__all__ = [MailDealer]