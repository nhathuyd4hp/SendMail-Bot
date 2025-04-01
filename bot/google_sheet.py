import pandas as pd
import logging
import json
import requests
from typing import Tuple,Union
from enum import StrEnum

class GoogleSheet:
    def __init__(self,application_scripts_url:str,timeout:int=10,logger_name:str=__name__,):
        self.application_scripts_url = application_scripts_url
        self.timeout = timeout
        self.logger = logging.getLogger(logger_name)
        self.client = requests.Session()
        self.client.headers.update(
            {"Content-Type": "application/json"},
        )  
        
    def getSheet(self,sheet_id: Union[StrEnum,str]) -> Tuple[str,pd.DataFrame] | None:
        """ 
        Args:
            sheet_id (_type_): sheet_id:any
        Returns:
            Tuple[str,pd.DataFrame]: sheet_name, data 
            None: None
        """
        if isinstance(sheet_id,StrEnum):
            sheet_id = sheet_id.value
        
        response = self.client.get(
            url = f"{self.application_scripts_url}?gid={sheet_id}",
            timeout = self.timeout,
        )
        
        if response.status_code == 200:
            content = json.loads(response.content.decode("utf-8"))
            if "error" in content:
                self.logger.error(content['error'])
                return None
            else:
                # ------- #
                columns = content['data'][0]
                data = content['data'][1:]
                sheet_name = content['sheet_name']
                # ------- #
                sheet = pd.DataFrame(data,columns=columns)
                # ------- #
                sheet['background'] = [background for background in content['backgrounds'][1:]]
                return (sheet_name,sheet)
        else:
            self.logger.error("Lỗi đọc dữ liệu từ Google Sheet")
            return None
        
    def setColor(self,sheet_id: Union[StrEnum,str],row,color:str) -> bool:
        self.client.post()
        
            