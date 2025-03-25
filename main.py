#!.venv\Scripts\python.exe
import pandas as pd
from bot import GoogleSheet
from bot import SharePoint
from enum import StrEnum
import logging

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

class Color(StrEnum):
    Blue = "#00ffff"
    Yellow = "#f1c232"
    White = "#ffffff"
    Gray = "#999999"
    Pink = "#ead1dc"
    Green = "#93c47d"
    
class Status(StrEnum):
    Completed = "#999999"  
    ProcessType1 = "#00ffff"  # Type 1
    ProcessType2 = "#93c47d"  # Type 2
    
class SheetMapping(StrEnum):
    KENTEC_1_T1_T6 = "976467386"
    KENTEC_2_T1_T6 = "1303713830"
    パネル_T1_T6 = "1709675894"
    
    def __str__(self):
        return self.value
   
APP_SCRIPTS_URL = "https://script.google.com/macros/s/AKfycbzQ_W_YHoffYZc-V-AqFOayYk_RTIFgr_owPKLMzsf5pZDCUoMIH331sAHuFlLyJaVT/exec"
TIME_OUT = 10

if __name__ == "__main__":
    google_sheet = GoogleSheet(
        application_scripts_url=APP_SCRIPTS_URL,
        timeout=TIME_OUT,
        logger_name="GoogleSheet"
    )
    for sid in SheetMapping:
        sheet_name,sheet = google_sheet.getSheet(sid)
        sheet.index = sheet.index + 2 # Tăng index lên 2 đơn vị (Excel bắt đầu từ 1 và Mất 1 dòng header)
        # Loại các các dòng có ID MAIL rỗng
        sheet = sheet[sheet["ID MAIL"] != ""] 
        # Duyệt qua các MAIL_ID
        MAIL_IDS = sheet["ID MAIL"].unique()
        for mail_id in MAIL_IDS:
            sheet_mail = sheet[sheet["ID MAIL"] == mail_id]
            bgs = sheet_mail['background'].to_list()
            
        
    # sharepoint = SharePoint(
    #     username="vietnamrpa@nskkogyo.onmicrosoft.com",
    #     password="Robot159753",
    #     logger_name="SharePoint",
    #     timeout=TIME_OUT,
    #     headless=True,
    # )