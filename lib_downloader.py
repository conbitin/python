import requests
import sys
from bs4 import BeautifulSoup
import os

url = 'http://10.252.250.53:8081/lm/content/groups/public/'
url2 = 'https://repo.maven.apache.org/maven2/'
command_prefix = 'wget --recursive --directory-prefix="D:/repo/" --level 0 --no-parent http://10.252.250.53:8081/lm/content/groups/public/'

def get_link_from_url(url_link):
    r = requests.get(url_link, allow_redirects=True, verify = False)
    soup = BeautifulSoup(r.content, 'html.parser')
    links = []
    for link in soup.findAll('a', href = True):
        url = link.get('href')
        if url.endswith('/'):
            split_url = url.rsplit('/', 2)
            real_link = split_url[len(split_url)-2]
            links.append(real_link)

    return links

if __name__ == "__main__":
    link1 = get_link_from_url(url)
    link2 = get_link_from_url(url2)
    for link in link1:
        if not link in link2:
            command = command_prefix + link + '/'
            print(command)
            os.system(command)
    sys.exit(1)