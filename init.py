import sys
from dbhelper import SqlLiteHelper

class Init:
    def __init__(self):
        self.helper = SqlLiteHelper()
    
    def initdb(self):
        self.helper.create()