import requests
import json
import psutil

class KnoxPortal:
    def __init__(self, cookies):
        self.session_request = requests.session()

        self.header = {
            'Accept': 'application/x-ms-application, image/jpeg, application/xaml+xml, image/gif, image/pjpeg, application/x-ms-xbap, application/x-shockwave-flash, */*'
            , 'Accept-Language': 'en-US'
            ,'Host': 'www.samsung.net'
            ,
            'User-Agent': 'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; WOW64; Trident/7.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; InfoPath.3; .NET4.0C; .NET4.0E)'
            , 'Accept-Encoding': 'gzip, deflate'
            , 'DNT': '1'
            , 'Connection': 'Keep-Alive'
            , 'Content-Type': 'application/x-www-form-urlencoded'
            , 'Referer' : 'http://www.samsung.net/portal/default.jsp'
        }

        self.s_header = {
            'Accept': 'application/json, text/javascript, */*; q=0.01'
            , 'Content-Type': 'application/json'
            , 'Accept-Language':'en-US'
            , 'X-Requested-With': 'XMLHttpRequest'
            , 'Referer': 'http://kr2.samsung.net/employee/empsearch/rest/v1/page/search/empSearchPopup?language=en&tabId=EMP'
            , 'User-Agent': 'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; WOW64; Trident/7.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; InfoPath.3; .NET4.0C; .NET4.0E)'
            , 'Accept-Encoding': 'gzip, deflate'
            , 'DNT': '1'
            , 'Connection': 'Keep-Alive'
            , 'Content-Length': '174'
            , 'Cache-Control': 'no-cache'
        }

        self.url = "http://kr2.samsung.net/employee/empsearch/rest/v1/emp/search"
        self.cookies = cookies
        self.data = {
            "queryBound":"TOTAL",
            "queryString":"a",
            "queryScope":"ALL",
            "sortType":"default",
            "empListCount":50,
            "currentPage":1,
            "attributes":[],"lang":"en",
            "adminSearchYn":"N",
            "englishOnly":"false"
        }

    # TODO
    def autoLogin(self, user, password):
        # self.data = "mode="
        # self.data += "&ICODE=N"
        # self.data += "&MCODE=N"
        # self.data += "&USERID=" + user
        # self.data += "&LOGOUTURL="
        # self.data += "&LOGINIMGUSETIME=0"
        # self.data += "&EPTRAYDATA="
        # self.data += "&LOGINTOTALVIEWTIME=0"
        # self.data += "&LANG=en_US.EUC-KR"
        # self.data += "&USERPASSWORD=" + password
        # self.data += "&noSessionPrc=Y"
        # self.data += "&login=Y"
        # self.data += "&domainGlobalPosition=KR"
        # self.data += "&mHost=cas102.samsung.com"
        # self.data += "&userBrnchCd=KR"
        # self.data += "&COMPID=C10"
        # self.data += "&SORGID=SEAHQ"
        # self.data += "&DEPTID=C10DA13DA130370"
        # self.data += "&SECID=5"
        # self.data += "&EPID=D120628112521C102840"
        # response = requests.post('http://www.samsung.net/employee/memberadmin/rest/v1/page/nosession/personalinfo/agree', headers=self.header, cookies = self.cookies, data = self.data)
        # print(str(response.content))

        self.data  = "EPTRAYDATA="
        self.data += "&ICODE=N"
        self.data += "&LANG=en_US.EUC-KR"
        self.data += "&LOGINIMGUSETIME=0"
        self.data += "&LOGINTOTALVIEWTIME=0"
        self.data += "&LOGOUTURL="
        self.data += "&MCODE=N"
        self.data += "&USERID=" + user
        self.data += "&USERPASSWORD=" + password
        self.data += "&mode=logout"
        self.data += "&noSessionPrc=Y"
        self.data += "&login=Y"
        self.data += "&domainGlobalPosition=KR"
        self.data += "&mHost=cas102.samsung.com"
        self.data += "&userBrnchCd=KR"
        self.data += "&agreehistory=Y"
        self.data += "&agreehistory=Y"
        self.data += "&AUTHCODE=Y"
        self.data += "&checkcmd="
        self.data += "&locale=en_US.EUC-KR"
        self.data += "&myinfo=Y"
        self.data += "&name=" + user
        self.data += "&rdmyinfo3=on"
        self.data += "&rdmyinfo4=on"
        self.data += "&rdmyinfo5=on"
        self.data += "&cmd=Login"
        response = requests.post('http://www.samsung.net/portal/login/login.do', headers=self.header, cookies = self.cookies, data = self.data)
        print(str(response.content))

    def verifyMemberById(self, single_id):
        data = self.data
        del data["queryBound"]
        data["queryScope"] = "ID_ENTER"
        data["queryString"] = single_id

        result = requests.post(self.url, headers=self.s_header, cookies=self.cookies, data=json.dumps(data))
        try: 
            if (result.json()['totalElements'] > 0):
                return True
        except:
            return False
        return False
    
    def isPortalLoggedIn(self):
        result = self.verifyMemberById('duc.quynh')
        if result:
            for proc in psutil.process_iter():
                if proc.is_running() and 'EpTray.exe' in proc.name():
                    print('Knox Portal has been logged in')
                    return True

        return False
    