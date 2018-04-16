import requests, json
import os, sys
import time

user_name = "duc.quynh"
http_password = "THANKyou89"
base_url = "confluence_wiki_url"
request_header = {
    'Content-Type': 'application/json'
    ,'Content-Encoding': 'deflate'
}
backup_folder = 'D:/wiki/'
forbidden_name = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']

def mk_dir(dest):
    if not os.path.exists(dest):
        os.makedirs(dest)
    return dest

def get_space(rest_link):
    request_url = base_url + rest_link
    response = requests.get(request_url, headers=request_header, auth=(user_name, http_password))
    response_json = json.loads(response.content.decode('utf-8'))
    results = response_json['results']
    for space in results:
        space_type = space['type']
        space_name = space['name']
        space_name = remove_forbidden(space_name)
        new_folder = backup_folder + space_type + '/' + space_name
        mk_dir(new_folder)
        
        if 'homepage' in space['_expandable']:
            space_homepage = space['_expandable']['homepage']
            get_page(space_homepage, new_folder)
    
    links = response_json['_links']
    if 'next' in links.keys():
        get_space(links['next'])

def get_page(page_link, dest):
    # get & save content
    request_url = base_url + page_link + '?expand=body.view'
    response = requests.get(request_url, headers=request_header, auth=(user_name, http_password))
    response_json = json.loads(response.content.decode('utf-8'))
    title = response_json['title']
    value = response_json['body']['view']['value']
    if bool(value):
        title = remove_forbidden(title)
        file_path = dest + '/' + title + '.html'
        save_file(value, file_path, True)
    
    rest_link = page_link + '/child/page?start=0&limit=500&expand=body.view'
    get_child_page(rest_link, dest)

    rest_link = page_link + '/child/attachment?start=0&limit=500&expand=body.view'
    get_child_attachment(rest_link, dest)

def get_child_page(page_link, dest):
    request_url = base_url + page_link
    response = requests.get(request_url, headers=request_header, auth=(user_name, http_password))
    response_json = json.loads(response.content.decode('utf-8'))
    results = response_json['results']
    for page in results:
        page_title = page['title']
        page_value = page['body']['view']['value']
        page_title = remove_forbidden(page_title)
        file_path = dest + '/' + page_title + '.html'
        if bool(page_value):
            save_file(page_value, file_path, True)
        if 'children' in page['_expandable']:
            child_link = page['_expandable']['children']

            child_rest_link  = child_link + '/page?start=0&limit=500&expand=body.view'
            get_child_page(child_rest_link, dest)

            child_rest_link  = child_link + '/attachment?start=0&limit=500&expand=body.view'
            get_child_attachment(child_rest_link, dest)
    
    links = response_json['_links']
    if 'next' in links.keys():
        get_child_page(links['next'], dest)

def get_child_attachment(page_link, dest):
    request_url = base_url + page_link
    response = requests.get(request_url, headers=request_header, auth=(user_name, http_password))
    response_json = json.loads(response.content.decode('utf-8'))
    results = response_json['results']
    for file in results:
        download_file_title = file['title']
        download_link = file['_links']['download']
        request_url = base_url + download_link
        print(request_url)
        data_response = requests.get(request_url, headers=request_header, auth=(user_name, http_password))
        download_file_title = remove_forbidden(download_file_title)
        file_path = dest + '/' + download_file_title
        save_file(data_response.content, file_path, None)
        
    links = response_json['_links']
    if 'next' in links.keys():
        get_child_attachment(links['next'], dest)

def save_file(data, file_path, encoded):
    if encoded:
        new_file = open(file_path, "w", encoding='utf-8')
    else:
        new_file = open(file_path, "wb")        
    new_file.write(data)
    new_file.close()

def remove_forbidden(name):
    for ch in forbidden_name:
        name = name.replace(ch, '')
    return name

if __name__ == "__main__":
    print("Start ..................")

    space_type = ['global', 'personal']
    for type in space_type:
        mk_dir(backup_folder + type)
    start = "/rest/api/space?start=0&limit=500&expand=body.view"
    get_space(start)

    print("End ...................")