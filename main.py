import requests
import json
from wireshark import WireShark 
from knoxportal import KnoxPortal
from plm import PLM
import time
import sys
import string
from dbhelper import SqlLiteHelper
from init import Init


if __name__ == "__main__":
    print("Start...................")

    # Init
    Init().initdb()

    # Pass authetication by cookie
    cookie = {}
    retry = 0
    while not 'EP_LOGINID' in cookie and retry < 4:
        retry += 1
        WireShark().execute()
        time.sleep(10)
        cookie = WireShark().get_cookie_raw()

    if not 'EP_LOGINID' in cookie:
        print("Please log in mysingle first and try again")
        sys.exit(1)

    # Get employee information
    alphabet = string.ascii_lowercase  #"abcdefghijklmnopqrstuvwxyz"
    reversed_alphabet = alphabet[::-1]
    helper = SqlLiteHelper()
    
    for c in list(reversed_alphabet):
      print(c)
      data = KnoxPortal(cookie).get_member_by_c(c)
      helper.insert(data)


    # plm system
    # result = PLM(cookie).get_my_issues()
    #print(result)


    print("End...................")
