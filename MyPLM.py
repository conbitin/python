import requests
import os
from datetime import date, timedelta
import pandas as pd
import ReadDataFromExcel as DataEXL
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

    def downloadPLMissues(self, userList, dest_file, user_id):
        
        # prepare data
        converted_user_list =""
        last_user = userList[len(userList) - 1]
        for user in userList:
            converted_user_list += user
            if user != last_user:
                converted_user_list += '%2C'
        
        today = date.today().strftime("%Y.%m.%d")
        last_90_days = (date.today() - timedelta(days=90)).strftime("%Y.%m.%d")

        file = open("sample input/plm_ids.json", 'r')
        json_data = json.load(file)
        create_user = json_data[user_id]['id']
        mem_group_id = json_data[user_id]['group_id']

        # get object references. Maybe PLM use them to keep data consistency.
        # If we use wrong object references, we still get data but they are missing something
        self.url = 'http://splm.sec.samsung.net/wl/tqm/statistics/getDefectByUserIFrame.do?pagingYn=Y'

        self.data = "refObjectId="
        self.data += "&resolveUser="
        self.data += "&searchYn=Y"
        self.data += "&orderBy="
        self.data += "&orderByType="
        self.data += "&jspPathName=tqm.statistics.statistics.jsp.DefectByUser.jsp"
        self.data += "&memberGroupChangeYn=N"
        self.data += "&checkJobToken="
        self.data += "&checkJobTokenType="
        self.data += "&mailReceiver="
        self.data += "&additionalHtmlContents="
        self.data += "&extSearchYn=N"
        self.data += "&pageIndex=1"
        self.data += "&pageUnit=10"
        self.data += "&pageSizeChangeYn=N"
        self.data += "&refObjectIdPbs="
        self.data += "&refObjectIdDev="
        self.data += "&refObjectIdMfg="
        self.data += "&refObjectIdMaint="
        self.data += "&refObjectIdEtc="
        self.data += "&projectTypeAllCheck=MULTIALLCHECK"
        self.data += "&projectType=PRE"
        self.data += "&projectType=BASIC"
        self.data += "&projectType=SET"
        self.data += "&projectType=SUPPORT"
        self.data += "&projectType=SW"
        self.data += "&projectType=MAINT"
        self.data += "&projectType=ETC"
        self.data += "&projectStatus=NP"
        self.data += "&projectStatus=DP"
        self.data += "&projectStatus=NC"
        self.data += "&projectStatus=DC"
        self.data += "&isDevVerify=A"
        self.data += "&ownerType=Y"
        self.data += "&regStartDate=" + last_90_days
        self.data += "&regEndDate=" + today
        self.data += "&memberGroupId=" + mem_group_id
        self.data += "&delayStandard="
        self.data += "&tag="
        self.data += "&testUnit="
        self.data += "&testCategory="
        self.data += "&testItem="
        self.data += "&functionBlock="
        self.data += "&detailFunctionclass="
        self.data += "&occurType="
        self.data += "&swVersion="
        self.data += "&swResolveVersion="
        self.data += "&closeVersion="
        self.data += "&overlapCheck="
        self.data += "&afterMgmtYn="
        self.data += "&pageSize=10"
        # self.data += "&perfomCheckStime=1522837723006"
        self.data += "&browserTimezone=GMT%2B7"
        self.data += "&perfomCheckurl="

        response = requests.post(self.url, headers=self.header, cookies = self.cookies, data = self.data)
        html_data = str(response.content)

        refObjectIdPbs = self.findVarFromText(html_data, '"refObjectIdPbs" : "', '",')
        refObjectIdDev = self.findVarFromText(html_data, '"refObjectIdDev" : "', '",')
        refObjectIdMfg = self.findVarFromText(html_data, '"refObjectIdMfg" : "', '",')
        refObjectIdMaint = self.findVarFromText(html_data, '"refObjectIdMaint" : "', '",')
        refObjectIdEtc = self.findVarFromText(html_data, '"refObjectIdEtc" : "', '",')

        refObjectIdPbs = refObjectIdPbs.replace(',', '%2C')
        refObjectIdDev = refObjectIdDev.replace(',', '%2C')
        refObjectIdMfg = refObjectIdMfg.replace(',', '%2C')
        refObjectIdMaint = refObjectIdMaint.replace(',', '%2C')
        refObjectIdEtc = refObjectIdEtc.replace(',', '%2C')
        

        # Now, download data
        self.url = 'http://splm.sec.samsung.net/wl/tqm/defect/defectreg/getDefectCombinedExcel.do'

        self.data = "cBox="
        self.data += "&isDeleteY="
        self.data += "&pageIndex=1"
        self.data += "&pageUnit=10"
        self.data += "&orderBy="
        self.data += "&orderByType="
        self.data += "&menuGubun="
        self.data += "&paramGubun="
        self.data += "&frmSearchFilterYn=N"
        self.data += "&paramOnePath="
        self.data += "&paramGubunEx=N"
        self.data += "&projectModelType="
        self.data += "&displayType=LIST"
        self.data += "&workType=DETECTION"
        self.data += "&objectId="
        self.data += "&projectId="
        self.data += "&deptId="
        self.data += "&deptName="
        self.data += "&displayDept="
        self.data += "&importance="
        self.data += "&nextClReleaseYn="
        self.data += "&tempEtc1="
        self.data += "&swVersionSel="
        self.data += "&status=ALL"
        self.data += "&openSubStatus="
        self.data += "&testPlanId="
        self.data += "&occurPhase="
        self.data += "&occurType="
        self.data += "&occurTypeName="
        self.data += "&occurTypeIsNull="
        self.data += "&swVersion="
        self.data += "&functionBlock="
        self.data += "&functionBlockIsNull="
        self.data += "&detailCategory="
        self.data += "&detailCategoryfrm="
        self.data += "&afterMgmtYn="
        self.data += "&afterStatus="
        self.data += "&searchDefectCategory="
        self.data += "&defectType="
        self.data += "&defectCategory="
        self.data += "&defectCategoryStr="
        self.data += "&testUnit="
        self.data += "&testItemId="
        self.data += "&testItemIdIsNull="
        self.data += "&swResolveType="
        self.data += "&hwVersion="
        self.data += "&hwVerIsNull=+"
        self.data += "&userId="
        self.data += "&defaultDefectId="
        self.data += "&popup=Y"
        self.data += "&isPopUp=Y"
        self.data += "&callType="
        self.data += "&defectSelectable=Y"
        self.data += "&testTypeSqlStr="
        self.data += "&selectRowId="
        self.data += "&statusStr="
        self.data += "&createUser=" + create_user
        self.data += "&periodType="
        self.data += "&period="
        self.data += "&isSwDefect="
        self.data += "&addOption="
        self.data += "&ddTestType="
        self.data += "&mainInChargeUserId="
        self.data += "&mainInChargeUserIds=" + converted_user_list
        self.data += "&frmDefectStat="
        self.data += "&mainInChargeYn=Y"
        self.data += "&createUser=" + create_user
        self.data += "&createStartDate="
        self.data += "&createEndDate="
        self.data += "&subInChargeUserId="
        self.data += "&regDate="
        self.data += "&changeId="
        self.data += "&isDevVerify="
        self.data += "&frmDetailYn=Y"
        self.data += "&excelMaxCount=50000"
        self.data += "&sliceUnit=10000"
        self.data += "&maintItemName="
        self.data += "&swCategory="
        self.data += "&pjtProjectClass="
        self.data += "&tqmDefectRange=T"
        self.data += "&defectCode="
        self.data += "&devModelName="
        self.data += "&mfgModelCode="
        self.data += "&title="
        self.data += "&projectName="
        # self.data += "&regStartDate=" + i[0]
        # self.data += "&regEndDate=" + i[1]
        self.data += "&resolveStartDate="
        self.data += "&resolveEndDate="
        self.data += "&closeStartDate="
        self.data += "&closeEndDate="
        self.data += "&rejStartDate="
        self.data += "&rejEndDate="
        self.data += "&resStartDate="
        self.data += "&resEndDate="
        self.data += "&createUserName="
        self.data += "&createUserIds="
        self.data += "&createUserId="
        self.data += "&inChargeUserName="
        self.data += "&inChargeUserIds="
        self.data += "&inChargeUserId="
        self.data += "&resolveUserName="
        self.data += "&resolveUserIds="
        self.data += "&resolveUserId="
        self.data += "&hierarchySearchDeptYn="
        self.data += "&occurRateType="
        self.data += "&occurLanguage="
        self.data += "&reqVer="
        self.data += "&swResolveVer="
        self.data += "&checksum="
        self.data += "&swCloseVer="
        self.data += "&cqId="
        self.data += "&testCaseId="
        self.data += "&testCaseYn="
        self.data += "&searchContent="
        self.data += "&searchReason="
        self.data += "&searchCountermeasure="
        self.data += "&detailProblemType="
        self.data += "&detailProblemTypeName="
        self.data += "&overlapCheck="
        self.data += "&sideEffectYn="
        self.data += "&failCaseYn="
        self.data += "&failureType="
        self.data += "&failureTypes="
        self.data += "&partCode="
        self.data += "&commonManage="
        self.data += "&commonCategory="
        self.data += "&categoryDetail="
        self.data += "&projectNameStr="
        self.data += "&projectStatus="
        self.data += "&projectStatusArr=NP"
        self.data += "&projectStatusArr=DP"
        self.data += "&projectStatusArr=NC"
        self.data += "&projectStatusArr=DC"
        self.data += "&frmTgStatus=N"
        self.data += "&tgUserId="
        self.data += "&pjtObjectId="
        self.data += "&hwVer="
        self.data += "&rejectYn="
        self.data += "&tqmAuthYn=N"
        self.data += "&mktProjectId="
        self.data += "&modelGroupYn="
        self.data += "&checkJobTokenType=file"
        # self.data += "&checkJobToken=5366543401770"
        self.data += "&deptCode="
        self.data += "&tgLevel1="
        self.data += "&tgLevel2="
        self.data += "&delayPeriod="
        self.data += "&resolveUnopen="
        self.data += "&openAuthYn="
        self.data += "&notExcel=N"
        self.data += "&deptType="
        self.data += "&searchFilter1AllChecked=N"
        self.data += "&searchFilter2AllChecked=N"
        self.data += "&searchFilter3AllChecked=N"
        self.data += "&searchFilter4AllChecked=N"
        self.data += "&searchFilter5AllChecked=N"
        self.data += "&searchFilter6AllChecked=N"
        self.data += "&searchFilter6_2AllChecked=N"
        self.data += "&searchFilter7AllChecked=N"
        self.data += "&searchFilter8AllChecked=N"
        self.data += "&searchFilter9AllChecked=N"
        self.data += "&searchFilter10AllChecked=N"
        self.data += "&searchFilter17AllChecked=N"
        self.data += "&searchFilter18AllChecked=N"
        # self.data += "&refObjectIdPbs=00GTYH8IDtPMWL1000%2C001N0X5RXtPMWL1000"
        self.data += "&refObjectIdPbs=" + refObjectIdPbs
        # self.data += "&refObjectIdDev=00GZRIUTVtPMWL1000%2C00GP1QR9UtPMWL1000%2C00GNJBSYHtPMWL1000%2C00GOYWE9ZtPMWL1000%2C00FUW2ST6tPMWL1000%2C00G8Q9ZYPtPMWL1000%2C00GIT0HTBtPMWL1000%2C00GIT0D9QtPMWL1000%2C00GTOHO25tPMWL1000%2C00H8I0N4DtPMWL1000%2C00GNVICRWtPMWL1000%2C00GNVI9CEtPMWL1000%2C00H8AEG0OtPMWL1000%2C00H8AEM1WtPMWL1000%2C00H8AFUZGtPMWL1000%2C00H8AEHQKtPMWL1000%2C00H8AEJA0tPMWL1000%2C00H8AG2B3tPMWL1000%2C00H8AEGN2tPMWL1000%2C00HAAFSI7tPMWL1000%2C00H8ADZ3FtPMWL1000%2C00H8AE6GTtPMWL1000%2C00H8AE2DZtPMWL1000%2C00H8AE0U6tPMWL1000%2C00H8AE3NMtPMWL1000%2C00H8ADZNVtPMWL1000%2C00FX2BJSHtPMWL1000%2C00GM23HIStPMWL1000%2C00GM23LFYtPMWL1000%2C00GTOHJV0tPMWL1000%2C00H8I0P74tPMWL1000%2C00GP4O7Y5tPMWL1000%2C00GJNYJAJtPMWL1000%2C00GX0034NtPMWL1000%2C00G86P2WWtPMWL1000%2C00GQDCIGFtPMWL1000%2C00GGLDRF8tPMWL1000%2C00GW3ZNWWtPMWL1000%2C00GMF666LtPMWL1000%2C00GM23QRYtPMWL1000%2C00GL1P5H4tPMWL1000%2C00GN64JBCtPMWL1000%2C00GM25RVCtPMWL1000%2C00GQ1UDX5tPMWL1000%2C00H4OK9FAtPMWL1000%2C00GUGR3H7tPMWL1000%2C00GPGVA6TtPMWL1000%2C00H23Q3EXtPMWL1000%2C00GPGVEXLtPMWL1000%2C00GM1OGAMtPMWL1000%2C00GM1NR1ZtPMWL1000%2C00GM1M24ItPMWL1000%2C00GL1PWTHtPMWL1000%2C00GM1T31AtPMWL1000%2C00GM23MA9tPMWL1000%2C00GQ1TVQ7tPMWL1000%2C00H23LKTYtPMWL1000%2C00EIFXIDBtPMWL1000%2C00EIFX798tPMWL1000%2C00EIFXA08tPMWL1000%2C00GQXO589tPMWL1000%2C00GRARC0EtPMWL1000%2C00FPDDB3WtPMWL1000%2C00FPDD6TYtPMWL1000%2C00FPDD5D9tPMWL1000%2C00FATQM5OtPMWL1000%2C00FPDD84AtPMWL1000%2C00GLXNWSFtPMWL1000%2C00H1B551XtPMWL1000%2C00H1B4WA0tPMWL1000%2C00H1B503FtPMWL1000%2C00GWYSFKOtPMWL1000%2C00GLX22Q4tPMWL1000%2C00GWYSPSOtPMWL1000%2C00GWYSNUBtPMWL1000%2C00GWYSLQOtPMWL1000%2C00GFW2M2NtPMWL1000%2C00FUC6R59tPMWL1000%2C00GFW2OEFtPMWL1000%2C00GFW2PJGtPMWL1000%2C00FJ47FFYtPMWL1000%2C00F9ANF8WtPMWL1000%2C00FJ47DYKtPMWL1000%2C00FJ47CD9tPMWL1000%2C00FLAM2YKtPMWL1000%2C00G70DBXVtPMWL1000%2C00FRW7D93tPMWL1000%2C00FTXYFLRtPMWL1000%2C00FS30LV9tPMWL1000%2C00GFTD8QVtPMWL1000%2C00GFTFHCYtPMWL1000%2C00GFTEIC6tPMWL1000%2C00GFTDWXVtPMWL1000%2C00FEA9JYLtPMWL1000%2C00GO9DH6StPMWL1000%2C00GL7UNNAtPMWL1000%2C00G0LOTG0tPMWL1000%2C00GX22W0WtPMWL1000%2C00GVUYXWVtPMWL1000%2C00GNKK9SZtPMWL1000%2C00GJNVVHJtPMWL1000%2C00GIT084ItPMWL1000%2C00GISZZWWtPMWL1000%2C00FPLSK4GtPMWL1000%2C00GGLCJ2JtPMWL1000%2C00H45C830tPMWL1000%2C00H2Y4P0OtPMWL1000%2C00GFPKHNBtPMWL1000"
        self.data += "&refObjectIdDev=" + refObjectIdDev
        self.data += "&refObjectIdMfg=" + refObjectIdMfg
        # self.data += "&refObjectIdMaint=00DH3YB26tPMWL1000%2C00GNW0Z7YtPMWL1000%2C00GOB8OA2tPMWL1000%2C00GU1OY6UtPMWL1000%2C00GQEFQYItPMWL1000%2C00GG8HBJOtPMWL1000%2C00H2YV1ABtPMWL1000%2C00H1R0EX7tPMWL1000%2C00H1SAAMFtPMWL1000%2C00ET93JO5tPMWL1000%2C00GT9HBAMtPMWL1000%2C00H2JEFZVtPMWL1000%2C00GOKSP3TtPMWL1000%2C00F0POYDMtPMWL1000%2C00GOKSTH1tPMWL1000%2C00GJMFU1DtPMWL1000%2C00H76ML80tPMWL1000%2C00F6LJ9PTtPMWL1000%2C00GR6IL2KtPMWL1000%2C00H7069QWtPMWL1000%2C00FTOIVGYtPMWL1000%2C00FT9HYZLtPMWL1000%2C00F80WZHGtPMWL1000%2C00FL87M08tPMWL1000%2C00G5GF58DtPMWL1000"
        self.data += "&refObjectIdMaint=" + refObjectIdMaint
        # self.data += "&refObjectIdEtc=00AK8H2AStPMWL1000%2C00DSQYAADtPMWL1000%2C00DGT49ABtPMWL1000%2C00EPMREFPtPMWL1000%2C00DW4XJVYtPMWL1000%2C00DY3OJ38tPMWL1000%2C00DIQ5DI6tPMWL1000%2C00DLMLHBDtPMWL1000%2C00AKOTBRAtPMWL1000%2C00EUTA1TQtPMWL1000%2C00FBNJ2VGtPMWL1000%2C00EK4Z3MFtPMWL1000%2C00EK4ZQDQtPMWL1000%2C00HG9W38FtPMWL1000%2C00FT6FUUKtPMWL1000%2C00FZF6HEEtPMWL1000%2C00G1EPQUXtPMWL1000%2C00FWPNLHStPMWL1000%2C00GW81T3MtPMWL1000%2C00HA0YHN6tPMWL1000%2C00HA0YOP9tPMWL1000%2C00H9ZO5QVtPMWL1000%2C00H9V682LtPMWL1000%2C00HACBE1WtPMWL1000%2C00HEI745OtPMWL1000%2C00H8VJJ6MtPMWL1000%2C00HHSZ6HBtPMWL1000%2C00HAH1RHNtPMWL1000%2C00HB8CDHXtPMWL1000%2C00DX8XVA5tPMWL1000%2C00ELC80TRtPMWL1000%2C00FZF7DK3tPMWL1000%2C00G1EPW91tPMWL1000%2C00FWP6SMTtPMWL1000%2C00GGA5M2RtPMWL1000%2C00ESDQB36tPMWL1000%2C00EN3QINTtPMWL1000%2C00GOAOE8PtPMWL1000%2C00FZ3RQ24tPMWL1000%2C009GN56RRtPMWL1000%2C00976JH5LtPMWL1000%2C00DGPBFXOtPMWL1000%2C00BM6R4ETtPMWL1000%2C00CJ7511NtPMWL1000%2C00EYEVK6UtPMWL1000%2C00ESX5QFWtPMWL1000%2C00FBMB1HHtPMWL1000%2C00FAV58L6tPMWL1000%2C00EUQRDHDtPMWL1000%2C00G8O57I5tPMWL1000%2C00F6P0GQFtPMWL1000%2C00FB9Z6NCtPMWL1000%2C00FPKTI37tPMWL1000%2C00EIJ4PSTtPMWL1000%2C00GNG7D23tPMWL1000%2C00HK4NBSPtPMWL1000%2C00GUI3QSGtPMWL1000%2C00GBCS1FRtPMWL1000%2C00GAZ2FQStPMWL1000%2C00H20EGEOtPMWL1000%2C00DR2YSC4tPMWL1000%2C00HI4UV6VtPMWL1000%2C00FRIKBYZtPMWL1000%2C00DEE5OB0tPMWL1000%2C008OYWMOBtPMWL1000%2C008TCEES9tPMWL1000%2C00CV76EVAtPMWL1000%2C00D5OOVS3tPMWL1000%2C00CV74YTXtPMWL1000%2C00D4PMK38tPMWL1000%2C00CTD6BM4tPMWL1000%2C00D4PLKBBtPMWL1000%2C00DNYE1CStPMWL1000%2C00CXPIXN4tPMWL1000%2C00DBZD3ZFtPMWL1000%2C00CXPXEDUtPMWL1000%2C00D6392SUtPMWL1000%2C00EL5GW8MtPMWL1000%2C00CTONL06tPMWL1000%2C00DGUNSSOtPMWL1000%2C00CQ9DFAPtPMWL1000%2C00C5Q82RHtPMWL1000%2C00GW7UOFVtPMWL1000%2C00GXVRDITtPMWL1000%2C008A4D5P7tPMWL1000%2C00CB4Z5PUtPMWL1000%2C00CPWT665tPMWL1000%2C00CVO3DD4tPMWL1000%2C00CQ9DAAVtPMWL1000%2C00HEI3LCGtPMWL1000%2C00CLXNE40tPMWL1000%2C00GWADX65tPMWL1000%2C00HEG8PKVtPMWL1000%2C00CB4Z99YtPMWL1000%2C00CPWT2CStPMWL1000%2C00HH9IPSJtPMWL1000%2C00DGICMBUtPMWL1000%2C00CRE58PJtPMWL1000%2C00DOVFX0HtPMWL1000%2C00GJLS1HRtPMWL1000%2C00H4OJONDtPMWL1000%2C00DCY7IZQtPMWL1000%2C00GT6SI30tPMWL1000%2C00HK2RVVTtPMWL1000%2C00GISZYOMtPMWL1000%2C00FBLQNUStPMWL1000%2C00GIQ1JAGtPMWL1000%2C00DBKXRDFtPMWL1000%2C00H4OJVKGtPMWL1000%2C00GISE9XPtPMWL1000%2C00DCR6BK5tPMWL1000%2C00GWJUYKQtPMWL1000%2C00H3P7DV2tPMWL1000%2C00FWWBUG4tPMWL1000%2C00GRLFT8WtPMWL1000%2C00G6XRILPtPMWL1000%2C00H7IKCAMtPMWL1000%2C00H272G7LtPMWL1000%2C00GAMA62HtPMWL1000%2C00FRMX77JtPMWL1000%2C00GK1H1VPtPMWL1000%2C00GAIBQF6tPMWL1000%2C00FWWBVKXtPMWL1000%2C00H7AZKB3tPMWL1000%2C00H9UJ023tPMWL1000%2C00FC9YIXPtPMWL1000%2C00H1CCQWJtPMWL1000%2C00FZW5DMYtPMWL1000%2C00G8JKRFWtPMWL1000%2C00H6Z0BKStPMWL1000%2C00H6YV8DGtPMWL1000%2C00H9UJ6UJtPMWL1000%2C00G7D5OLAtPMWL1000%2C00G6VAMJ1tPMWL1000%2C008RHAKC5tPMWL1000%2C008O2PNDAtPMWL1000%2C008X0J29DtPMWL1000%2C00B5YE7AGtPMWL1000%2C00BM6QW1UtPMWL1000%2C00HAT3GETtPMWL1000%2C00HAT25F7tPMWL1000%2C00AZW9ZG6tPMWL1000%2C007TUSUGHtPMWL1000%2C007DZWK9FtPMWL1000%2C00HB5S9GMtPMWL1000%2C00FJ25ZEFtPMWL1000%2C00G2MY3C8tPMWL1000%2C00GA3JCO0tPMWL1000%2C00H4M6ZLRtPMWL1000%2C00HJNGEQ8tPMWL1000%2C00FM5E9BFtPMWL1000%2C00GA3K849tPMWL1000%2C00HAT33C2tPMWL1000%2C00HDZCUS6tPMWL1000%2C00HCXP1URtPMWL1000%2C00ETIO24AtPMWL1000%2C00EAWVWB8tPMWL1000%2C00F5WTWBStPMWL1000%2C00E53NMPDtPMWL1000%2C00EMFPKUAtPMWL1000%2C00HDZBCBJtPMWL1000%2C00HH3MK20tPMWL1000%2C00H19VQEYtPMWL1000%2C00G6XYFN9tPMWL1000%2C009PRKB17tPMWL1000%2C00FIC5ZNJtPMWL1000%2C00ESLB6JVtPMWL1000%2C00DZHVMWGtPMWL1000%2C00E8090TCtPMWL1000%2C00DLSIP09tPMWL1000%2C00E782JKJtPMWL1000%2C00E0C0XU3tPMWL1000%2C00HGN5O42tPMWL1000%2C00HDY57PFtPMWL1000%2C00HGB3G0BtPMWL1000%2C00AWKV2YKtPMWL1000%2C009DCGUWTtPMWL1000%2C00ASQ2V50tPMWL1000%2C00AGT41VLtPMWL1000%2C009K8T1ULtPMWL1000%2C009PDTPPOtPMWL1000%2C00F2WWUHUtPMWL1000%2C00EKWQS63tPMWL1000%2C00EFN5J2StPMWL1000%2C00E4RWX10tPMWL1000%2C00F3MB5J6tPMWL1000%2C00EMT6VLAtPMWL1000%2C00HACD04QtPMWL1000%2C00GV0P0U8tPMWL1000%2C00H5E8312tPMWL1000%2C00H687EFZtPMWL1000%2C00HD7QRUQtPMWL1000%2C00HD7XQPRtPMWL1000%2C00DMBA5TPtPMWL1000%2C00DMAM7AWtPMWL1000%2C00DF4MWBItPMWL1000%2C00EJS8PH4tPMWL1000%2C00E0K6XQCtPMWL1000%2C00FU406B6tPMWL1000%2C00E0K9AUZtPMWL1000%2C00HHPIBVZtPMWL1000%2C00H1AF5SNtPMWL1000%2C008WRRIW6tPMWL1000%2C008THDFDUtPMWL1000%2C00D1X57MItPMWL1000%2C00DWE0WY0tPMWL1000%2C00DH01NG1tPMWL1000%2C00DGUNPXCtPMWL1000%2C008X1ORSWtPMWL1000%2C007EOXQSNtPMWL1000%2C00DADHYUAtPMWL1000%2C00HCOA8MFtPMWL1000%2C00EBXFWDYtPMWL1000%2C00EBXG380tPMWL1000%2C00HAT0DP9tPMWL1000%2C00EBXG05YtPMWL1000%2C00EXR1K37tPMWL1000%2C00GNHFVMWtPMWL1000%2C00EVK65UKtPMWL1000%2C00GSVV390tPMWL1000%2C00EXD427LtPMWL1000%2C00GZ0N4CXtPMWL1000%2C00GZCEYG2tPMWL1000%2C00HG4W3LPtPMWL1000%2C00EO3ECCRtPMWL1000%2C00F84M8BWtPMWL1000%2C00E8D1PZOtPMWL1000%2C00DO4HX0YtPMWL1000%2C009FEYLVJtPMWL1000%2C00EVU4XTKtPMWL1000%2C00FG5RMA9tPMWL1000%2C00FKY1108tPMWL1000%2C00EXC13ZZtPMWL1000%2C00FLCJRZItPMWL1000%2C00F9D6G78tPMWL1000%2C00F9D8KNUtPMWL1000%2C00E0CXJ4FtPMWL1000%2C00DMPOB1JtPMWL1000%2C00DY953N7tPMWL1000%2C00D1YVZPPtPMWL1000%2C00DLM3GB8tPMWL1000%2C00D6TLZUXtPMWL1000%2C00D1DM9KXtPMWL1000%2C00DQFWSF7tPMWL1000%2C00H1KYRP0tPMWL1000%2C00H291PVUtPMWL1000%2C00GXEEU7DtPMWL1000%2C00HD51LY1tPMWL1000%2C00GW5WQWHtPMWL1000%2C00HD5JQ46tPMWL1000%2C00HDOC2AOtPMWL1000%2C00GRWMZNXtPMWL1000%2C00GW6JHM4tPMWL1000%2C00GXF2NFWtPMWL1000%2C00HDZC07PtPMWL1000%2C00HI4WMC4tPMWL1000%2C006E1801WtPMWL1000%2C00CCDBSL8tPMWL1000%2C00EYP4IHNtPMWL1000%2C00G5DLPNPtPMWL1000%2C00GL7SZDKtPMWL1000%2C0096VG5H7tPMWL1000%2C0097N1WQItPMWL1000%2C00956I0GYtPMWL1000%2C0097MU72UtPMWL1000%2C0097N22MAtPMWL1000%2C0097MQ884tPMWL1000%2C0097N2AMLtPMWL1000%2C0097N2BSPtPMWL1000%2C00CXY2NOPtPMWL1000%2C00GQ1I7BGtPMWL1000%2C00HDB0XYFtPMWL1000%2C00FYROJTVtPMWL1000%2C00GTLTBOZtPMWL1000%2C00GG9KWT2tPMWL1000%2C00GWKJPZLtPMWL1000%2C002KCOMQItPMWL1000%2C009JRNWB9tPMWL1000%2C00GOCOU7EtPMWL1000%2C00FV8XYMWtPMWL1000%2C00FYRZ64UtPMWL1000%2C002L4KSIDtPMWL1000%2C00D0BHRDXtPMWL1000%2C00D0BHJL7tPMWL1000%2C00AQKZG83tPMWL1000%2C00B295BEXtPMWL1000%2C009236E2OtPMWL1000"
        self.data += "&refObjectIdEtc=" + refObjectIdEtc
        self.data += "&classification="
        self.data += "&testUnitSw=00036UZ8QtPMWL0000%2C005YBJ0I1tPMWL1000%2C00ASOPQ0ZtPMWL1000"
        self.data += "&reviewDept="
        self.data += "&reviewerName="
        self.data += "&reviewerIds="
        self.data += "&reviewResult="
        self.data += "&searchType="
        self.data += "&category="
        self.data += "&defectId="
        self.data += "&allInChargeUserName="
        self.data += "&allInChargeUserIds="
        self.data += "&createUserVendorName="
        self.data += "&createUserVendorIds="
        self.data += "&inChargeUserVendorName="
        self.data += "&inChargeUserVendorIds="
        self.data += "&resolveUserVendorName="
        self.data += "&resolveUserVendorIds="
        self.data += "&regDeptCode="
        self.data += "&solDeptCode="
        self.data += "&ccListName="
        self.data += "&ccListIds="
        self.data += "&defectSearchType="
        self.data += "&resolveStatusIds="
        self.data += "&resolveStatus2Ids="
        self.data += "&tgCategory1="
        self.data += "&tgCategory2Ids="
        self.data += "&statusIds="
        self.data += "&tag="
        self.data += "&comment="
        self.data += "&logComment="
        self.data += "&occurPhaseIds="
        self.data += "&testCategoryIds="
        self.data += "&testItemIds="
        self.data += "&functionBlockIds="
        self.data += "&detailFunctionIds="
        self.data += "&importanceIds="
        self.data += "&salesRegionIds="
        self.data += "&reviewResultIds="
        self.data += "&tgEmpty="
        self.data += "&swVersionSubmitYn="
        self.data += "&opVersionInfoId="
        self.data += "&swVersionName="
        self.data += "&searchUserType="
        self.data += "&delayReleasePeriodFrom="
        self.data += "&delayReleasePeriodTo="
        self.data += "&ddReqReason="
        self.data += "&resourceCode="
        self.data += "&testItemName="
        self.data += "&modelNumber="
        self.data += "&osVersion="
        self.data += "&nation="
        self.data += "&isCommentExcel=N"
        self.data += "&deletedDefectYn="
        self.data += "&convergenceYn="
        self.data += "&convergenceDetail="
        self.data += "&confirmStatus="
        self.data += "&listOrder=*"
        self.data += "&searchOption=TITLE"
        self.data += "&searchText="
        self.data += "&pageSize=10"

        # download one by one file
        count_temp = 0
        for i in range(90, 0, -15):
            start_reg_day = (date.today() - timedelta(days=i)).strftime("%Y.%m.%d")
            end_reg_day = (date.today() - timedelta(days=i-14)).strftime("%Y.%m.%d")
            if i == 15:
                end_reg_day = today

            query_data = self.data
            query_data += "&regStartDate=" + start_reg_day
            query_data += "&regEndDate=" + end_reg_day
            response = requests.post(self.url, headers=self.header, cookies = self.cookies, data = query_data)
            temp_file = '%d.xls' %count_temp
            print("Downloading data %s from %s to %s ..." % (temp_file, start_reg_day, end_reg_day))
            data_type = response.headers['Content-Type']
            if data_type and not 'application/octet' in data_type:
                raise ValueError('Opps! Have problem when downloading from PLM')
            file = open(temp_file, "wb")
            file.write(response.content)
            file.close()
            count_temp += 1

        # merge multi excel files into one
        all_data = pd.DataFrame()
        for i in range(count_temp):
            temp_file = '%d.xls' %i
            df, length = DataEXL.read_excel_issue(temp_file)
            all_data = all_data.append(df, ignore_index=True)
            
        # wrtie to dest file. But first remove existing file
        if os.path.exists(dest_file):
            os.remove(dest_file)
        
        writer = pd.ExcelWriter(dest_file)
        all_data.to_excel(writer, sheet_name='DEFECT', startrow = 2, index=False)
        writer.save()

        #remove temp files
        for i in range(count_temp):
            temp_file = '%d.xls' %i
            os.remove(temp_file)

    def findVarFromText(self, source, str, end_c):
        start_index = source.find(str)
        source = source[start_index:]
        end_index = source.find(end_c)
        return source[len(str):end_index]