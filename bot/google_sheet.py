import pandas as pd
import logging
import json
import requests
from typing import Tuple

class GoogleSheet:
    def __init__(self,application_scripts_url:str,timeout:int=10,logger_name:str=__name__,):
        self.application_scripts_url = application_scripts_url
        self.timeout = timeout
        self.logger = logging.getLogger(logger_name)
        
    def getSheet(self,sheet_id) -> Tuple[str,pd.DataFrame] | None:
        """ 
        Args:
            sheet_id (_type_): sheet_id:any

        Returns:
            Tuple[str,pd.DataFrame]: sheet_name, data 
            None: None
        """
        sheet_id = str(sheet_id)
        url = f"{self.application_scripts_url}?gid={sheet_id}"
        response = requests.get(
            url = url,
            timeout=self.timeout,
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
                sheet['background'] = [list(set(background)) for background in content['backgrounds'][1:]]
                return (sheet_name,sheet)
        else:
            self.logger.error("Lỗi đọc dữ liệu từ Google Sheet")
            return None
            