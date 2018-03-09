from P4 import P4, P4Exception
import sys, os, time


P4_PORT = "************"
P4_USER = "******"
P4_PWD = "******"
P4_CLIENT = "***************"

class DownloadThread():
    def __init__(self, port, user, pwd, client):
        self.p4 = P4()
        self.p4.port = port
        self.p4.user = user
        self.p4.password = pwd
        self.p4.client = client
        self.ROOT = ""
    
    def download_files(self, project_path, target_count):
        try:
            #print(project_path)
            project_path_full = '//' + project_path + '/...'
            project_depotFiles = self.p4.run('files', '-e', project_path_full)
            for prj in project_depotFiles:
                file = prj['depotFile']
                if not '/DEV/' in file and not '/REL/' in file and not 'dummy.txt' in file:
                    print(file.encode('utf-8'))
                    self.p4.run('sync', '-f', file)

        except P4Exception:
            print("Download Files Exception ...")
            for e in self.p4.errors:
                print(e)
                
    def run(self):
        print('run')
        try:
            self.p4.connect()
            self.p4.run_login()
            client = self.p4.fetch_client()
            self.ROOT = client['Root']
            projects = self.p4.run('files', '-e', '//ADMIN/AutoProtect/Source/PACKAGES/...')
            count = 1
            for prj in projects:
                # prj['depotFile'] = //ADMIN/AutoProtect/Source/PACKAGES/apps/LiveBroadcast.csv
                prj_admin_link = prj['depotFile']
                prj_path = prj_admin_link[len('//ADMIN/AutoProtect/Source/'):len(prj_admin_link)-4]
                self.download_files(prj_path, count)
                count += 1
            self.p4.disconnect()
        except P4Exception:
            print("Run Exception ...")
            for e in self.p4.errors:
                print(e)


if __name__ == "__main__":
    DownloadThread(P4_PORT, P4_USER, P4_PWD, P4_CLIENT).run()
    sys.exit(1)