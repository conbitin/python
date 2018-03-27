import sys, os, time, shutil
from selenium import webdriver
import IEHelper
from WireShark import PLMWireShark as PWS
from MyPLM import PLM
import ReadDataFromExcel as DataEXL
import json

need_update_user_list = False  # set True when our organization is updated (for example: add new members)

class AutoDownloadPLM:
    def __init__(self, user_name, pass_word):
        self.user_name = user_name
        self.pass_word = pass_word

        self.excel_file_name = "DEFECT_LIST_Today_Basic.xls"
        self.dir_sample_input = os.getcwd() + "\\sample input\\"

        self.autoLoginKnox()
        self.cookies = self.getPLMCookie()
        self.list_plm_id = self.getUserPLMidbyKnoxId()
        self.downloadPLMissues()

    def autoLoginKnox(self):
        IEHelper.set_zoom_100()
        self.driver = webdriver.Ie()
        self.driver.maximize_window()
        
        """ fill in user name and pass """

        self.driver.get("http://samsung.net/")
        if self.driver.current_url == 'http://kr2.samsung.net/portal/desktop/main.do':
            print("Knox has been login ")
            pass
        else:
            user = self.driver.find_element_by_id('USERID')
            user.send_keys(self.user_name)
            password = self.driver.find_element_by_id('USERPASSWORD')
            password.send_keys(self.pass_word)
            button = self.driver.find_element_by_class_name('btnLogin')
            button.click()

    def getPLMCookie(self):
        cookie = {}
        retry = 0

        while not 'EP_LOGINID' in cookie and retry < 5:
            retry += 1
            PWS().execute()
            time.sleep(10)
            cookie = PWS().get_cookie_raw()
        
        if not 'EP_LOGINID' in cookie:
            print("Please log in mysingle first and try again")
            time.sleep(3)
            self.restart_program()

        return cookie

    def getUserPLMidbyKnoxId(self):
        global need_update_user_list
        dict_single_id_plm_id = {}
        list_user_id = []

        if need_update_user_list:
            list_single_ids = DataEXL.get_list_single_id_member()

            for each_id in list_single_ids:
                user_plm_id = PLM(self.cookies).getUserIdbyKnoxId(each_id)
                print(each_id, user_plm_id)
                if user_plm_id:
                    dict_single_id_plm_id[each_id] = user_plm_id
                    list_user_id.append(user_plm_id)
            
            file = open(self.dir_sample_input + 'plm_user.json', 'w')
            file.write(json.dumps(dict_single_id_plm_id))
            file.close()
            need_update_user_list = False
            return list_user_id
        
        file = open(self.dir_sample_input + 'plm_user.json', 'r')
        dict_single_id_plm_id = json.load(file)
        for key in dict_single_id_plm_id.keys():
            list_user_id.append(dict_single_id_plm_id[key])

        return list_user_id

    def downloadPLMissues(self):
        """ start download PLM """
        dest = self.dir_sample_input + self.excel_file_name
        PLM(self.cookies).downloadPLMissues(self.list_plm_id, dest)        

    def restart_program(self):
        """ restart program Main again """
        try:
            self.driver.quit()
        except:
            pass
        time.sleep(5)
        python = sys.executable
        os.execv(python, ['python'] + ['V2_Home.py'])
