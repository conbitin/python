import requests, json
import os, sys
import time

user_name = "*******"
http_password = "********"
gerrit_http_url = "http://165.213.202.48/package/"
gerrit_clone_url = "165.213.202.48:2727"
request_header = {
    'Accept': 'application/json'
    , 'Accept-Encoding': 'gzip'
}

output_dir_prefix = "android"

if __name__ == "__main__":
    print("Start ..................")
    request_url = gerrit_http_url + "a/projects/"
    result = requests.get(request_url, headers = request_header, auth=(user_name, http_password))
    result_json = json.loads(result.text[4:])
    clone_url_base = 'ssh://' + user_name + '@' + gerrit_clone_url
    for project in result_json.keys():
        if not "Experiment/" in project:
            clone_url_project = clone_url_base + '/' + project
            output_dir = output_dir_prefix + '/' + project
            os.system('git clone %s "%s"' % (clone_url_project, output_dir))

    print("End ...................")