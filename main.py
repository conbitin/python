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
    cookies = WireShark().get_cookie_raw('Portal')
    if not KnoxPortal(cookies).is_Portal_logged_in():
        print("Please log in mysingle first and try again")
        sys.exit(1)

    # Get employee information
    # alphabet = string.ascii_lowercase  #"abcdefghijklmnopqrstuvwxyz"
    alphabet = "abcdefghijkl"
    reversed_alphabet = alphabet[::-1]
    helper = SqlLiteHelper()
    
    for c in list(reversed_alphabet):
      print(c)
      data = KnoxPortal(cookies).get_member_by_c(c)
      helper.insert(data)

    helper.make_db_for_idm()

    # plm system
    # result = PLM(cookie).get_my_issues()
    #print(result)


    print("End...................")
