
class Cookie:
    WMONID = 'WMONID'
    ajs_user_id ='ajs_user_id'
    ajs_group_id = 'ajs_group_id'
    ajs_anonymous_id = 'ajs_anonymous_id'
    EPV7PTSID = 'EPV7PTSID'
    EP_LOGINID = '@EP_LOGINID'
    EPV7EMSID = 'EPV7EMSID'
    EPV7APSID = 'EPV7APSID'
    CLLIVESESSIONID = 'CLLIVESESSIONID'
    EPV7MLSID = 'EPV7MLSID'
    EP6_UTOKEN = 'EP6_UTOKEN'
    EP_BROWSERID = 'EP_BROWSERID'
    EP_LOGINID = 'EP_LOGINID'
    portal_token_key = 'portal_token_key'
    eBS = 'eBS'
    PORTALSESSIONID = 'PORTALSESSIONID'
    WLSESSIONID = 'WLSESSIONID'
    WLPBOMSESSID = 'WLPBOMSESSID'

    def __init__(self, cookie_dict, src):
        if self.WMONID in cookie_dict and self.EPV7PTSID in cookie_dict and self.EPV7EMSID in cookie_dict and self.EP6_UTOKEN in cookie_dict and self.EP_LOGINID in cookie_dict:
            tlogin = 'tlogin_' + cookie_dict[self.EP_LOGINID]
            if 'Portal' in src:
                self.cookies = {'approvalPanelType': 'vertical', 'approvalPanelWidth': '563', 'WMONID': cookie_dict[self.WMONID],
                                'logimage.index': '0',
                                'saveLanguage': 'en_US.EUC-KR', 'ajs_user_id': cookie_dict[self.ajs_user_id], 'ajs_group_id': cookie_dict[self.ajs_group_id],
                                'ajs_anonymous_id': cookie_dict[self.ajs_anonymous_id],
                                'EPV7PTSID': cookie_dict[self.EPV7PTSID],
                                '@EP_LOGINID': cookie_dict[self.EP_LOGINID], tlogin: cookie_dict[tlogin],
                                'EPV7EMSID': cookie_dict[self.EPV7EMSID],
                                'EPV7APSID': cookie_dict[self.EPV7APSID],
                                'CLLIVESESSIONID': cookie_dict[self.CLLIVESESSIONID],
                                'EPV7MLSID': cookie_dict[self.EPV7MLSID],
                                'EP6_UTOKEN': cookie_dict[self.EP6_UTOKEN],
                                'EP_BROWSERID': cookie_dict[self.EP_BROWSERID], 'EP_LOGINID': cookie_dict[self.EP_LOGINID]}
            elif 'PLM' in src:
                self.cookies = {'WMONID': cookie_dict[self.WMONID],
                                'logimage.index': '0',
                                'saveLanguage': 'en_US.EUC-KR',
                                'EPV7PTSID': cookie_dict[self.EPV7PTSID],
                                '@EP_LOGINID': cookie_dict[self.EP_LOGINID],
                                'EPV7APSID': cookie_dict[self.EPV7APSID],
                                'CLLIVESESSIONID': cookie_dict[self.CLLIVESESSIONID],
                                'EPV7MLSID': cookie_dict[self.EPV7MLSID],
                                'EP6_UTOKEN': cookie_dict[self.EP6_UTOKEN],
                                'EP_BROWSERID': cookie_dict[self.EP_BROWSERID], 'EP_LOGINID': cookie_dict[self.EP_LOGINID],
                                'eBS': cookie_dict[self.eBS], 'PORTALSESSIONID' : cookie_dict[self.PORTALSESSIONID],
                                'WLSESSIONID' : cookie_dict[self.WLSESSIONID], 'WLPBOMSESSID' : cookie_dict[self.WLPBOMSESSID]}                
        else:
            self.cookies = {}
            print('Cookie %s is invalid' % str(cookie_dict))
    
    def get_cookie(self):
        return self.cookies