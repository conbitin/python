from scapy.all import *
import pythoncom
from comtypes.client import CreateObject, GetActiveObject
import time

class WireShark:
    PORTAL = ['http://samsung.net', '112.107.220.15']
    PLM = ['http://splm.sec.samsung.net/wl/tqm/statistics/getDefectByUser.do?fromPlmMainMenu=true', '10.43.83.21']
    payload = ''
    packet_count = 0
    cookie_dict = {}
    target = []

    class IeThread(threading.Thread):
        def __init__(self, url):
            threading.Thread.__init__(self)
            self.url = url
        def run(self):
            pythoncom.CoInitialize()
            try:
                ie = CreateObject("InternetExplorer.Application")
            except:
                ie = GetActiveObject("InternetExplorer.Application")
            ie.Visible = 0
            ie.Navigate(self.url)

    def get_cookie_raw(self, target):
        if 'Portal' in target:
            self.target = WireShark.PORTAL
        elif 'PLM' in target:
            self.target = WireShark.PLM

        self.cookie_dict = {}

        retry = 0
        while not 'WMONID' in self.cookie_dict and retry < 5:
            retry += 1
            self.packet_count = 0
            self.execute(self.target[0])
            time.sleep(10)
            
        return self.cookie_dict

    def parse_payload(self, http_payload):
        cookie_raw = http_payload[http_payload.find('Cookie')+8:http_payload.find(r"\r\n\r\n")]
        cookie_data = cookie_raw.split('; ')
        for item in cookie_data:
            equal = item.find('=')
            self.cookie_dict[item[:equal]] = item[equal+1:]
        print(self.cookie_dict)

        
    def packet_callback(self, packet):
        print("Sniffing network packet ...")
        self.packet_count += 1
        if str(packet[IP].dst) == self.target[1] and packet[TCP].payload:
            raw_payload_binary = str(packet[TCP].payload)
            self.payload = "".join((self.payload, raw_payload_binary[2:-1]))

        if self.packet_count == 5:
            self.parse_payload(self.payload)
        
    def execute(self, target_url):
        WireShark.IeThread(target_url).start()
        time.sleep(1)
        p = sniff(filter='tcp and port 80', timeout=10, count = 5, prn = self.packet_callback, store=0)

    