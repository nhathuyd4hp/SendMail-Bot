import pandas as pd
from bot import GoogleSheet
from bot import SharePoint
from enum import Enum
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

class Color(Enum):
    Blue = "#00ffff"
    Yellow = "#f1c232"
    White = "#ffffff"
    Gray = "#999999"
    Pink = "#ead1dc"
    Green = "#93c47d"
    
from enum import Enum

class Status(Enum):
    Completed = "#999999"  
    ProcessType1 = "#00ffff"  # Type 1
    ProcessType2 = "#93c47d"  # Type 2
    
class SheetMapping(Enum):
    KENTEC_1_T1_T6 = "976467386"
    KENTEC_2_T1_T6 = "1303713830"
    パネル_T1_T6 = "1709675894"
    
    def __str__(self):
        return self.value
    
# def pre_processing(df:pd.DataFrame) -> pd.DataFrame:
#     return df[df["ID MAIL"] != ""]

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
        # Loại các các dòng có ID MAIL rỗng
        sheet = sheet[sheet["ID MAIL"] != ""] 
        # Chỉ lấy các dòng là Status.ProcessType1 hoặc Status.ProcessType2
        sheet = sheet[sheet["background"].apply(lambda bg_list: 
            (len(bg_list) == 1 and bg_list[0] in [Status.ProcessType1.value, Status.ProcessType2.value]) 
            or 
            (len(bg_list) == 2 and any(color in [Status.ProcessType1.value, Status.ProcessType2.value] for color in bg_list) and Color.White.value in bg_list)
        )]
        sheet.to_excel(f"{sheet_name}.xlsx",index_label="STT")    
    # sharepoint = SharePoint(
    #     username="vietnamrpa@nskkogyo.onmicrosoft.com",
    #     password="Robot159753",
    #     logger_name="SharePoint",
    #     timeout=TIME_OUT,
    #     headless=True,
    # )