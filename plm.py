import requests
import os
from datetime import date, timedelta
import pandas as pd
import json

class PLM:
    def __init__(self, cookies):
        self.session_request = requests.session()

        self.header = {
            'Accept': 'application/x-ms-application, image/jpeg, application/xaml+xml, image/gif, image/pjpeg, application/x-ms-xbap, application/x-shockwave-flash, */*'
            , 'Accept-Language': 'en-US'
            ,'Host': 'splm.sec.samsung.net'
            ,
            'User-Agent': 'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; WOW64; Trident/7.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; InfoPath.3; .NET4.0C; .NET4.0E)'
            , 'Accept-Encoding': 'gzip, deflate'
            , 'DNT': '1'
            , 'Connection': 'Keep-Alive'
            , 'Content-Type': 'application/x-www-form-urlencoded'
            , 'Referer': 'http://splm.sec.samsung.net/wl/tqm/defect/defectreg/getDefectSearchDetail.do'
        }

        self.url = "http://splm.sec.samsung.net/wl/tqm/defect/defectreg/getDefectCombinedExcel.do"

        self.cookies = cookies

        self.data = ""

    def get_my_issues(self):
        if not self.cookies:
            return "Cannot get cookie"
        response = requests.post(self.url, headers=self.header, cookies=self.cookies, data=self.data)
        print(response)

        file = open("out.xls", "wb")
        file.write(response.content)
        file.close()
        return ""

    def getUserIdbyKnoxId(self, id):
        self.url = 'http://splm.sec.samsung.net/wl/com/ums/umsGetUserListMultiWithFunction.do'
        self.data = 'category=singleID&beforeParseList=%s&isEpSearch=false&contentType=html&callBack=&popup=Y&divLimitYN=N&retired=&searchAll=Y&cond=id&searchWord=%s' % (id, id)
        response = requests.post(self.url, headers=self.header, cookies = self.cookies, data = self.data)
        html_data = str(response.content)
        # We are going to get user id from html text. In Html text, there is something like
        # userId = "D160115094933C103239";
        # mail= "ba.lv@samsung.com";

        user_id = self.findVarFromText(html_data, 'userId = "', '";')
            
        return user_id

    def findVarFromText(self, source, str, end_c):
        start_index = source.find(str)
        source = source[start_index:]
        end_index = source.find(end_c)
        return source[len(str):end_index]