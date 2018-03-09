import requests
import json
from cookie import Cookie
import re

class PLM:
    def __init__(self, cookie_raw):
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
        }
        self.url = "http://splm.sec.samsung.net/wl/tqm/defect/defectsol/getDefectSolDefaultData.do"
        self.cookies = Cookie(cookie_raw, 'PLM').get_cookie()

    def get_my_issues(self):
        if not self.cookies:
            return "Cannot get cookie"
        response = requests.post(self.url, headers=self.header, cookies=self.cookies)
        html_data = str(response.content)
        start_index = html_data.find('&mainInChargeUserId=')
        html_data = html_data[start_index:]
        end_index = html_data.find('"')
        user_id = html_data[len('&mainInChargeUserId='):end_index]
        print(user_id)
        return html_data
