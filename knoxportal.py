import requests
import json
import string
from cookie import Cookie


class KnoxPortal:
    def __init__(self, cookie_raw):
        self.session_request = requests.session()
        self.data = {
            "queryBound": "TOTAL",
            "queryString": "a",
            "queryScope": "ALL",
            "sortType": "default",
            "empListCount": 50,
            "currentPage": 1,
            "attributes": [], "lang": "en",
            "adminSearchYn": "N",
            "englishOnly": "false"
        }
        self.header = {
            'Accept': 'application/json, text/javascript, */*; q=0.01'
            , 'Content-Type': 'application/json'
            , 'Accept-Language': 'en-US'
            , 'X-Requested-With': 'XMLHttpRequest'
            ,
            'Referer': 'http://kr2.samsung.net/employee/empsearch/rest/v1/page/search/empSearchPopup?language=en&tabId=EMP'
            ,
            'User-Agent': 'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; WOW64; Trident/7.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; InfoPath.3; .NET4.0C; .NET4.0E)'
            , 'Accept-Encoding': 'gzip, deflate'
            , 'DNT': '1'
            , 'Connection': 'Keep-Alive'
            , 'Content-Length': '174'
            , 'Cache-Control': 'no-cache'
        }
        self.login_url = "http://kr2.samsung.net/employee/empsearch/rest/v1/page/search/empSearchPopup?language=en&tabId=EMP"
        self.url = "http://kr2.samsung.net/employee/empsearch/rest/v1/emp/search"
        self.cookies = Cookie(cookie_raw, 'Portal').get_cookie()

    def get_member_by_c(self, c):
        response = []
        data = self.data
        data["queryString"] = c
        page_number = 0
        total_page = 1
        while page_number < total_page:
            data["currentPage"] = page_number + 1
            result = requests.post(self.url, headers=self.header, cookies=self.cookies, data=json.dumps(data))
            json_data = result.json()
            # json_data = json.load(open("output.json"))

            if "totalElements" in json_data:
                total_element = json_data["totalElements"]
            else:
                return response
            if total_element > 0:
                elements = json_data["elements"]
                for e in elements:
                    response.append((e["compTelNo"], e["deptNm"], e["emailAddr"], e["empNo"], e["enChagBizCn"],
                                     e["enCompNm"], e["enDeptNm"], e["enFnm"], e["enJobgrdNm"], e["enJobplAddr"],
                                     e["enJobpoNm"], e["mphonNo"], e["userId"]))

                page_number = json_data["pageNumber"]
                total_page = json_data["totalPages"]
                print("page number = %s" % page_number)
                print("total page = %s" % total_page)
            else:
                page_number = total_page

        return response

        # return json.dumps(result.json(), indent=4, sort_keys=True)
        # return result.json()
        # print(result.json()['totalElements'])

    def get_mem_by_id(self, jsonData):
        data = self.data
        del data["queryBound"]
        data["queryScope"] = "ID_ENTER"
        data["queryString"] = jsonData["id"]

        result = requests.post(self.url, headers=self.header, cookies=self.cookies, data=json.dumps(data))
        return json.dumps(result.json(), indent=4, sort_keys=False)

    def verify_mem_by_id(self, single_id):
        data = self.data
        del data["queryBound"]
        data["queryScope"] = "ID_ENTER"
        data["queryString"] = single_id

        result = requests.post(self.url, headers=self.header, cookies=self.cookies, data=json.dumps(data))

        if (result.json()['totalElements'] > 0):
            return True
        return False

    def verify_multi_ids(self, data):
        response = {"ids": []}
        for id in data["ids"]:
            if self.verify_mem_by_id(id):
                response["ids"].append(id)

        return json.dumps(response)