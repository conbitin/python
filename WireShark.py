from scapy.all import *
import pythoncom
from win32com.client import Dispatch
import time

class PLMWireShark:
    PLM_URL = 'http://splm.sec.samsung.net/wl/tqm/statistics/getDefectByUser.do?fromPlmMainMenu=true'
    TARGET_IP_PLM = '10.43.83.21'
    payload = ''
    packet_count = 0
    cookie_dict = {}

    class IeThread(threading.Thread):
        def run(self):
            pythoncom.CoInitialize()
            ie = Dispatch("InternetExplorer.Application")
            ie.Visible = 0
            ie.Navigate(PLMWireShark.PLM_URL)

    def get_cookie_raw(self):
        return self.cookie_dict

    def parse_payload(self, http_payload):
        #print(http_payload)
        cookie_raw = http_payload[http_payload.find('Cookie')+8:http_payload.find(r"\r\n\r\n")]
        #print(cookie_raw)
        cookie_data = cookie_raw.split('; ')
        #print(cookie_data)
        for item in cookie_data:
            equal = item.find('=')
            self.cookie_dict[item[:equal]] = item[equal+1:]
        print(self.cookie_dict)

        
    def packet_callback(self, packet):
        print("packet_callback")
        self.packet_count += 1
        if str(packet[IP].dst) == self.TARGET_IP_PLM and packet[TCP].payload:
            raw_payload_binary = str(packet[TCP].payload)
            self.payload = "".join((self.payload, raw_payload_binary[2:-1]))

        if self.packet_count == 5:
            self.parse_payload(self.payload)
        
    def execute(self):
        PLMWireShark.IeThread().start()
        time.sleep(1)
        p = sniff(filter='tcp and port 80', timeout=10, count = 5, prn = self.packet_callback, store=0)