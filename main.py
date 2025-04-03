import os
import re
import logging
import time
import zipfile
import pandas as pd
from datetime import datetime
from bot import MailDealer
from bot import GoogleSheet
from bot import SharePoint
from enum import StrEnum


logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    encoding='utf-8',
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler('bot.log', mode='a', encoding='utf-8'),
        logging.StreamHandler(),
    ],
)
class ColorMapping(StrEnum):
    YELLOW = "#ffff00"
    WHITE = "#ffffff"
    BLUE = "#00ffff"
    GRAY = "#999999"
    GREEN = "#93c47d"
    
class SheetMapping(StrEnum):
    KENTEC_1_T1_T6 = "976467386"
    KENTEC_2_T1_T6 = "1303713830"
    パネル_T1_T6 = "1709675894"
    
    def __str__(self):
        return self.value
       
APP_SCRIPTS_URL = "https://script.google.com/macros/s/AKfycbzQ_W_YHoffYZc-V-AqFOayYk_RTIFgr_owPKLMzsf5pZDCUoMIH331sAHuFlLyJaVT/exec"
SHAREPOINT_DOWNLOAD_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),"SharePoint")
TIME_OUT = 10
HEADLESS = False

if __name__ == "__main__":
    while True:
        google_sheet = GoogleSheet(
            application_scripts_url=APP_SCRIPTS_URL,
            timeout=TIME_OUT,
            logger_name="GoogleSheet"
        )
        mail_dealer = MailDealer(
            url="https://mds3310.maildealer.jp/",
            username="vietnamrpa",
            password="nsk159753",
            timeout=TIME_OUT,
            headless=HEADLESS,
            logger_name="MailDealer"
        )
        share_point = SharePoint(
            url="https://nskkogyo.sharepoint.com/",
            username="vietnamrpa@nskkogyo.onmicrosoft.com",
            password="Robot159753",
            timeout=TIME_OUT,
            headless=HEADLESS,
            logger_name="SharePoint",
            download_directory=SHAREPOINT_DOWNLOAD_PATH,
        )
        if not (share_point.authenticated and mail_dealer.authenticated):
            continue
        # Process
        MAILS:list[dict] = []
        for sid in SheetMapping:
            sheet_name,sheet = google_sheet.getSheet(sid)
            sheet.index = sheet.index + 2 # Tăng index lên 2 đơn vị (Excel bắt đầu từ 1 và Mất 1 dòng header)
            # Loại các các dòng có ID MAIL rỗng
            sheet = sheet[
                sheet["ID MAIL"] != ""
            ]
            # Kiểm tra từng Mail ID
            MAIL_IDS = sheet['ID MAIL'].unique()
            for ID in MAIL_IDS:
                sheet_based_on_id = sheet[
                    sheet['ID MAIL'] == ID
                ]
                backgrounds: list[list[str]] = sheet_based_on_id['background'].to_list()
                if all(background[0] == ColorMapping.BLUE.value for background in backgrounds):
                    MAILS.append({
                        "mail_id":ID,
                        "sheet_id":sid.value,
                        "sheet_name":sid.name,
                        "row":sheet_based_on_id.index.to_list(),
                        "type":1,
                    })
                if all(background[0] == ColorMapping.GREEN.value for background in backgrounds):
                    MAILS.append({
                        "mail_id":ID,
                        "sheet_id":sid.value,
                        "sheet_name":sid.name,
                        "row":sheet_based_on_id.index.to_list(),
                        "type":2,
                    })
        data = []
        for mail in MAILS:
            if mail.get("type") == 1:
                mail_id = mail.get("mail_id")
                if mail.get("sheet_name") == "KENTEC_1_T1_T6":
                    status = mail_dealer.get_status(mail_id)
                    if status == "このメールは「返信処理中」です。" or status == None:
                        result,note = mail_dealer.reply(
                            mail_id=mail.get("mail_id"),
                            返信先="全員に返信",
                            署名="署名なし",
                            テンプレート=[
                                "kaneka",
                                "KKT Box"
                            ],
                        )
                        data.append({**mail, "result": result, "note": note})
                    else:
                        data.append({**mail, "result": False, "note": status})
                if mail.get("sheet_name") == "KENTEC_2_T1_T6":
                    status = mail_dealer.get_status(mail_id)
                    if status == "このメールは「返信処理中」です。" or status == None:
                        result,note = mail_dealer.reply(
                            mail_id=mail.get("mail_id"),
                            返信先="全員に返信",
                            署名="署名なし",
                            テンプレート=[
                                "kaneka",
                                "KFP Box"
                            ],
                        )
                        data.append({**mail, "result": result, "note": note})
                    else:
                        data.append({**mail, "result": False, "note": status})
            if mail.get("type") == 2:
                mail_id:str = mail.get("mail_id")
                folder = mail_id.replace("-1","")
                # ----- #
                status = mail_dealer.get_status(mail_id)
                if status == "このメールは「返信処理中」です。" or status == None:
                    download_files = []
                    site_url,files = share_point.get_files_in_folder(
                        site_folder="/sites/2021/Shared Documents/は行/KANEKA JOB",
                        folder=folder,
                    )
                    folders = [f for f in files if '.' not in f]
                    for folder in folders:
                        result,files = share_point.download_file(
                            site_url=site_url,
                            file_pattern=f"{folder}/.*.(pdf|xlsm|xlsx)$",
                        )
                        download_files.extend(files)
                    for file in os.listdir(SHAREPOINT_DOWNLOAD_PATH):
                        if file.lower().endswith('.zip'):
                            file_path = os.path.join(SHAREPOINT_DOWNLOAD_PATH, file)
                            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                                zip_ref.extractall(SHAREPOINT_DOWNLOAD_PATH)
                            os.remove(file_path)
                    time.sleep(1)
                    if download_files:
                        result,note = mail_dealer.reply(
                            mail_id=mail.get("mail_id"),
                            返信先="全員に返信",
                            署名="署名なし",
                            テンプレート=[
                                "kaneka",
                                "KFP file attach"
                            ],
                            attached=download_files,
                        )
                        data.append({**mail, "result": result, "note": note})           
                else:
                    data.append({**mail, "result": False, "note": status})           
                                
        data = pd.DataFrame(data)
        data.to_excel(f"{datetime.today().strftime("%Y-%m-%d_%H-%M-%S") }.xlsx",index=False)
        break
