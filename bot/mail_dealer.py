import time
import logging
import re
import pywinauto
from pathlib import Path
from typing import Union
import pandas as pd
from pywinauto import Application
from pywinauto.application import Application,WindowSpecification
from selenium.common.exceptions import TimeoutException,ElementClickInterceptedException,NoSuchElementException,StaleElementReferenceException
from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from common import HandleException

email_to_lastname = {
    "junichiro.kawakatsu@kaneka.co.jp":"川勝",
    "hideaki.takeshiro@kaneka.co.jp":"武城",
    "haruna.kuge@kaneka.co.jp":"久家",
    "shigeru.yokota@kaneka.co.jp":"横田",
    "takafumi.miki@kaneka.co.jp":"三木",
    "chie.nakamura@kaneka.co.jp":"中村",
    "kazuki.masuda1@kaneka.co.jp":"増田",
    "miku.ichinotsubo@kaneka.co.jp":"一ノ坪",
    "shigeru.yokota@kaneka.co.jp":"横田",
    "yuki.date@kaneka.co.jp":"伊達",
}
                
class MailDealer:
    # Init
    def __init__(
        self,
        url:str,
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
        self.url = url
        self.wait = WebDriverWait(self.browser, timeout)
        self.username = username
        self.password = password
        # Trạng thái đăng nhập
        self.authenticated = self.__authentication(username, password)
    # Delete   
    def __del__(self):
        if hasattr(self,"browser") and isinstance(self.browser,WebDriver):
            try:
                self.browser.quit()
            except:
                pass


    @HandleException
    def switch_to_default_content(func) : 
        def wrapper(self, *args, **kwargs):
            self.browser.switch_to.default_content()
            return func(self, *args, **kwargs)
        return wrapper 
    
    @HandleException
    def login_required(func) : 
        def wrapper(self, *args, **kwargs):
            if not self.authenticated:
                self.logger.error('❌ Yêu cầu xác thực')
                return None
            return func(self, *args, **kwargs)  
        return wrapper 
        
            
    
    # Method  
    @switch_to_default_content
    @HandleException
    def __authentication(self, username: str, password: str) -> bool:
        time.sleep(1)
        self.browser.get(self.url)
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
                self.logger.info('✅ Xác thực thành công!')
                return True
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
            
    @HandleException
    @login_required
    def read_mail(self, mail_box: str,mail_id:str,tab_name:str=None) -> str:
        content = ""
        if not self.browser.current_url.endswith('/app/'):
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

    @HandleException
    @switch_to_default_content
    @login_required
    def get_status(self,mail_id:str) -> str | None:
        search_input = self.wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR,"input[placeholder='このメールボックスのメール・電話を検索']"))
        )
        search_input.clear()
        search_input.send_keys(mail_id)
        # search
        self.wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR,"button[title='検索']"))
        )
        search_btn = self.wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR,"button[title='検索']"))
        )
        search_btn.click()
        # Wait loader
        time.sleep(0.5)
        WebDriverWait(self.browser, 120).until(
            EC.invisibility_of_element_located((By.CSS_SELECTOR, "div[class='loader']"))
        )
        time.sleep(0.5)
        # Switch frame
        ifame = self.wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR,"iframe[id='ifmMain']"))
        )
        self.browser.switch_to.frame(ifame)
        # Check Maillist
        try:
            self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR,"span[id='maillist']"))
            )
            self.logger.info(f"Không tìm thấy mail (id:{mail_id})")
            return "NotFound"
        except TimeoutException:
            pass
        olv_p_mail_page__content = self.wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR,"div[class='olv-p-mail-page__content']"))
        )
        snackbar__wrap = olv_p_mail_page__content.find_elements(By.CSS_SELECTOR,"div[class='snackbar__wrap']")
        if snackbar__wrap:
            msg = ", ".join([e.text for e in snackbar__wrap])
            self.logger.info(f"Mail Status {mail_id}: {msg}")
            return msg
        return None
        

    @HandleException
    @switch_to_default_content
    @login_required 
    def reply(self,
        mail_id:str,
        返信先:str = None, # Radio
        引用:list[str] = [], # CheckBox
        ToFrom:str = None, # Dropdown
        署名:str = None, # Dropdown
        テンプレート:list[str] = [], # list[Dropdown]
        attached:list[str] = [] # list[Attached]
    ) -> tuple[bool,str]:
        options = {key: value for key, value in locals().items() if value and key not in ["self", "mail_id","attached"]}
        self.logger.info(f"Reply Mail {mail_id}: {options}")
        # Search
        search_input = self.wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR,"input[placeholder='このメールボックスのメール・電話を検索']"))
        )
        search_input.clear()
        search_input.send_keys(mail_id)
        # search
        self.wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR,"button[title='検索']"))
        )
        search_btn = self.wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR,"button[title='検索']"))
        )
        search_btn.click()
        # Wait loader
        time.sleep(0.5)
        WebDriverWait(self.browser, 120).until(
            EC.invisibility_of_element_located((By.CSS_SELECTOR, "div[class='loader']"))
        )
        time.sleep(0.5)
        # Switch frame
        ifame = self.wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR,"iframe[id='ifmMain']"))
        )
        self.browser.switch_to.frame(ifame)
        # Check Maillist
        try:
            self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR,"span[id='maillist']"))
            )
            self.logger.info(f"Không tìm thấy mail (id:{mail_id})")
            return (False,f"Không tìm thấy mail (id:{mail_id})")
        except TimeoutException:
            pass
        olv_p_mail_page__content = self.wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR,"div[class='olv-p-mail-page__content']"))
        )
        snackbar__wrap = olv_p_mail_page__content.find_elements(By.CSS_SELECTOR,"div[class='snackbar__wrap']")
        if snackbar__wrap:
            msg = ", ".join([e.text for e in snackbar__wrap])
            if not msg == 'このメールは「返信処理中」です。':
                self.logger.info(f"Mail {mail_id}: {msg}")
                return (False,msg)
        olv_p_mail_ops__act_reply = self.wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR,"div[class='olv-p-mail-ops__act-reply']"))
        )
        # accent_btn__btn_wrap = olv_p_mail_ops__act_reply.find_element(By.CSS_SELECTOR,"span[class='accent-btn__btn-wrap']") What is this?
        menu = olv_p_mail_ops__act_reply.find_element(By.CSS_SELECTOR,"div[class='menu']")
        if options:
            menu.click()
            time.sleep(2)  
            menu = olv_p_mail_ops__act_reply.find_element(By.CSS_SELECTOR,"div[class='menu']")
            sections = menu.find_elements(By.XPATH,".//section")
            sections = [section for section in sections if section.find_element(By.TAG_NAME,'label').text in options]
            for section in sections:
                div_control = section.find_element(By.XPATH,"./div")
                # Radio
                if ops := div_control.find_elements(By.CSS_SELECTOR,"label[class^='radio']"):
                    label = section.find_element(By.XPATH, "./label") 
                    target_option = options.get(label.text)
                    target_op = next((op for op in ops if op.text == target_option), None)
                    if target_op:
                        target_op.click()
                # Checkbox                            
                if ops := div_control.find_elements(By.CSS_SELECTOR,"label[class^='checkbox']"):
                    label = section.find_element(By.XPATH, "./label")
                    pass
                # Dropdown
                if ops := div_control.find_elements(By.CSS_SELECTOR,"div[class^='dropdown ']"):
                    label = section.find_element(By.XPATH, "./label")
                    target_option = [options.get(label.text)] if isinstance(options.get(label.text), str) else options.get(label.text)
                    for index,op in enumerate(ops):
                        label.click()
                        op.click()
                        time.sleep(1)
                        lis = self.browser.find_elements(By.XPATH,f".//li[text()='{target_option[index]}']") 
                        if len(lis) == 1:
                            lis[0].click()
                        else:
                            self.logger.error(f"Không tìm thấy {label.text}:{target_option[index]}")
                            return (False,f"Không tìm thấy {label.text}:{target_option[index]}")
        # メール作成
        # ---- #
        button = menu.find_element(By.TAG_NAME,'footer').find_element(By.TAG_NAME,'button')
        self.wait.until(EC.element_to_be_clickable(button))
        current_windows = self.browser.window_handles
        button.click()
        # Wait for Open New Windows
        while len(self.browser.window_handles) <= len(current_windows):
            continue
        window_ids:set= set(self.browser.window_handles) - set(current_windows)
        while window_ids:
            email = None
            last_name = None
            id = window_ids.pop()
            self.browser.switch_to.window(id)
            self.browser.maximize_window()
            # 
            time.sleep(5)
            olv_p_mail_edit_sections = self.wait.until(
                EC.presence_of_all_elements_located(
                    (By.CSS_SELECTOR,"section[class='olv-p-mail-edit__section']")
                )
            )
            for olv_p_mail_edit_section in olv_p_mail_edit_sections:
                olv_p_mail_edit__dl_olv_u_mbs:list[WebElement] = olv_p_mail_edit_section.find_elements(By.CSS_SELECTOR,"div[class^='olv-p-mail-edit__dl olv-u-mb--']")
                for olv_p_mail_edit__dl_olv_u_mb in olv_p_mail_edit__dl_olv_u_mbs:
                    if olv_p_mail_edit__dl_olv_u_mb.find_elements(By.XPATH, ".//span[text()='To']"):
                        olv_p_mail_edit__dd = olv_p_mail_edit__dl_olv_u_mb.find_element(By.CSS_SELECTOR,"div[class='olv-p-mail-edit__dd']")
                        olv_p_mail_edit__dd_input = olv_p_mail_edit__dd.find_element(By.TAG_NAME,"input")
                        ToEmail = olv_p_mail_edit__dd_input.get_attribute('value')
                        match = re.search(r"<([^<>]+)>", ToEmail)
                        if match:
                            email:str = match.group(1)  
                            last_name = email_to_lastname.get(email.lower(),"ご担当者")
                        break
                

            if not email or not last_name:
                cancelButton = self.wait.until(
                    EC.presence_of_element_located((By.XPATH, ".//button[text()='キャンセル']"))
                )
                self.wait.until(
                    EC.element_to_be_clickable(cancelButton)
                ).click()
                time.sleep(2)
                try:
                    alert = self.wait.until(EC.alert_is_present())
                    alert.accept()
                except Exception:
                    pass
                self.browser.switch_to.window(current_windows[0])     
                self.logger.info(f"Reply Mail {mail_id} thất bại: {email} không phù hợp")    
                return False,f"Reply Mail {mail_id} thất bại: {email} không phù hợp"
            # -- Content
            content = None
            self.wait.until(
                EC.frame_to_be_available_and_switch_to_it((By.TAG_NAME,'iframe'))
            )
            body = self.wait.until(
                EC.presence_of_element_located((By.TAG_NAME,'body'))
            )
            if divs := body.find_elements(By.XPATH,'./div'):
                for div in divs:
                    if "○○" in div.get_attribute("innerHTML"):
                        new_html = div.get_attribute("innerHTML").replace("○○", last_name)
                        self.browser.execute_script("arguments[0].innerHTML = arguments[1];", div, new_html)
                        content = new_html.replace("<br>","\n")
                        time.sleep(2)
                        break
            self.browser.switch_to.default_content()
            # -- Attached
            if attached:
                time.sleep(1)
                for path in attached:
                    success = True
                    error = ""
                    button = self.wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR,"button[title='添付ファイル']"))
                    )
                    while True:
                        try:
                            self.wait.until(
                                EC.element_to_be_clickable(button)
                            ).click()
                            break
                        except ElementClickInterceptedException:
                            continue
                    time.sleep(1)
                    
                    open_application: Application = Application(backend="win32").connect(title="Open")
                    dialog_file_upload: WindowSpecification = open_application.window(title="Open")      
                                    
                    dialog_file_upload.child_window(class_name="Edit").type_keys(str(Path(path)).replace("(", "{(}").replace(")", "{)}"),with_spaces=True)
                    time.sleep(2)
                    dialog_file_upload.child_window(title="&Open", class_name="Button").wrapper_object().click() 
                    time.sleep(2)   
                    try:    
                        alert = self.wait.until(
                            EC.alert_is_present()
                        )
                        if alert.text == "添付可能なファイル容量の上限は20MBです。":
                            self.logger.info(f"Đính kèm {path} thất bại: 添付可能なファイル容量の上限は20MBです。")
                            alert.accept()
                            success = False
                            error = "添付可能なファイル容量の上限は20MBです。"
                            break
                    except TimeoutException:
                        pass
                        
                    while open_windows := pywinauto.findwindows.find_elements(title="Open"):
                        success = False
                        error = f"{path} not found"
                        for w in open_windows:
                            w: WindowSpecification = Application().connect(handle=w.handle).window(handle=w.handle)
                            if w.child_window(title="OK", class_name="Button").exists():
                                w.child_window(title="OK", class_name="Button").wrapper_object().click()
                            if w.child_window(title="Cancel", class_name="Button").exists():
                                w.child_window(title="Cancel", class_name="Button").wrapper_object().click()
                    #----------#
                    if not success:
                        self.logger.info(f"Đính kèm {path} thất bại")
                        return success,error
                    self.logger.info(f"Đính kèm {path} thành công")    
            
            while True:
                try:
                    div_menu = self.wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='menu']"))
                    )
                    children = div_menu.find_elements(By.XPATH,'./*')
                    if(len(children)) == 1:
                        child = children[0]
                        button = child.find_element(By.TAG_NAME,'button')
                        self.wait.until(EC.element_to_be_clickable(button)).click()
                        continue
                    else:
                        button = self.wait.until(
                            EC.presence_of_element_located((By.XPATH, ".//button[text()='一時保存']"))
                        )
                        self.wait.until(EC.element_to_be_clickable(button)).click()
                        break
                except TimeoutException:
                    continue
            
            self.browser.close()
            self.browser.switch_to.window(current_windows[0])     
            self.logger.info(f"Reply Mail {mail_id} thành công")
            return True,content      
        
            
              

__all__ = [MailDealer]