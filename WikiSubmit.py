from confluence import Api
import utils
import requests
import json
import datetime

wiki_url = "http://mobilerndhub.sec.samsung.net/wiki/"  # Samsung DPI Wiki entry url
space = "SVMC"
parentTitle = "BasicApp P - Auto Report"
pageTitle = "BA - Daily Report Test"
# pageUrgentPrjTitle = "BA - Urgent projects"
pageUrgentPrjTitle = "Test"
user = utils.open_file(".wiki")[0]
pw = utils.open_file(".wiki")[1]
api = Api(wiki_url, user, pw)
PRJ_MACRO_TAG = '</ac:structured-macro>'
FIRST_PRJ_TAG = '<ac:parameter ac:name="title">'
LAST_PRJ_TAG = '</ac:parameter>'


def updateWiki(pageTitle, htmlCode):
    """ Updates Wiki page with this new content """

    pageHeader = ""
    pageContent = pageHeader + htmlCode

    try:
        currentPage = api.getpage(name=pageTitle, spacekey=space)
        print("Updating extisting page.")
        page = api.updatepage(name=pageTitle, spacekey=space, content=pageContent,
                              page_id=currentPage['id'], label="tg")
        print("View page at %s" % page['url'])
    except Exception:
        print("Creating new page.")
        page = api.addpage(name=pageTitle, spacekey=space, content=pageContent,
                           parentpage=parentTitle)
        print("View page at %s" % page.page_url)


def get_urgent_project():
    """ get list urgent project """
    list = []
    try:
        content = api.getpagecontent(name=pageUrgentPrjTitle, spacekey=space)
        lines = str(content).split(PRJ_MACRO_TAG)
        for line in lines:
            # print(line)
            first = line.find(FIRST_PRJ_TAG)
            end = line.rfind(LAST_PRJ_TAG)
            if first > -1 and end > -1:
                list.append(line[first + len(FIRST_PRJ_TAG):end])
    except Exception:
        print("Cannot get page %s content" % pageUrgentPrjTitle)

    return list


def getListSingleID(data):
    """
    :param data:  table data
    :return: list mysingle to chart group
    """
    list_id = []
    index = data[0].index('Owner')
    for i in data:
        list_id.append(i[index])
    del (list_id[0])

    return list_id


def create_table(data):
    """ create table data """
    newline = "\n"
    empty_table = True
    html = """<table class="table table-striped table-bordered" style="width:100%">"""
    html += newline

    # Header
    html += "<thead><tr>"
    html += newline

    if type(data[0]) == str:
        pass
    else:
        for c in data[0]:
            html += "<th>"
            html += removeSpecialChar(str(c))
            if c == 'Owner':
                list_id = getListSingleID(data)
                html += makeLinkChatGroup(list_id)
            html += "</th>"
            html += newline

    html += "</tr></thead>"
    html += newline
    html += "<tbody>"
    # End Header

    if type(data[0]) != str:
        for r in data[1:]:
            html += "<tr>"
            html += newline
            for i in range(len(r)):
                c = r[i]
                if c != '0':
                    empty_table = False
                html += "<td>"
                if data[0][i] == 'Case code':
                    html += makeLinkPLM(c)
                elif data[0][i] == 'Owner':
                    html += makeLinkChat(c)
                else:
                    html += removeSpecialChar(str(c))
                html += "</td>"
                html += newline

            html += "</tr>"
            html += newline
    html += "</tbody></table>"
    return html, empty_table


def createHtmlChart(data, title, type_chart, width=500, height=360, is3d='False',
                    orientation='vertical', dataOri='horizontal', rangeAxis=50):
    htmltable, empty_table = create_table(data)
    minimum_range = 0
    if empty_table:
        rangeAxis = 10
        minimum_range = 1

    """ Generates an HTML table from analyzed data """
    html = """
       <ac:structured-macro ac:name="chart">
           <ac:parameter ac:name="title">{}</ac:parameter>
           <ac:parameter ac:name="type">{}</ac:parameter>
           <ac:parameter ac:name="width">{}</ac:parameter>
           <ac:parameter ac:name="height">{}</ac:parameter>
           <ac:parameter ac:name="3d">{}</ac:parameter>
           <ac:parameter ac:name="orientation">{}</ac:parameter>
           <ac:parameter ac:name="dataOrientation">{}</ac:parameter>
           <ac:parameter ac:name="rangeAxisTickUnit">{}</ac:parameter>
           <ac:parameter ac:name="rangeAxisLowerBound">{}</ac:parameter>
        <ac:rich-text-body>
  """.format(title, type_chart, width, height, is3d, orientation, dataOri, rangeAxis, minimum_range)
    html += htmltable
    html += r"""
            </ac:rich-text-body>
        </ac:structured-macro>
"""

    return html


def createLayoutWeekly(chart_week, chart_month, tg_table, tg_chart, table):
    html = """<head>
        </head>
        <body>
            <div  style="margin-top:50px;">
                <div class ="columnLayout two-equal" data-layout="two-equal">
                    <div class ="cell normal" data-type ="normal">
                        <p>{}</p>
                    </div>
                    <div class ="cell normal" data-type ="normal">
                        <p>{}</p>
                    </div>
                </div>
            </div>
            <div  style="margin-top:40px;">
                <hr></hr>
                <h2 style="margin-top:40px;">TG report</h2>
                <p style="color:gray;"> Last 7 days' Issue statistics </p>
                <div class ="columnLayout two-equal" data-layout="two-equal">
                    <div class ="cell normal" data-type ="normal">
                        <p>{}</p>
                    </div>
                    <div class ="cell normal" data-type ="normal">
                        <p>{}</p>
                    </div>
                </div>
            </div>
            <div>
                <h2 style="margin-top:50px;">Summary Issue</h2>
                <p style="color:gray;">This table include all projects (exclude MR project) and number of issues which we worked on last week </p>
                <p>{}</p>
            </div>
            <hr></hr>
        </body>
        """.format(chart_week, chart_month, tg_table, tg_chart, table)

    return html


def createLayoutMonthly(last_month_team, team_chart_line, tg_table, tg_pie):
    html = """<head>
        </head>
        <body>
            <div  style="margin-top:50px;">
                <div class ="columnLayout two-equal" data-layout="two-equal">
                    <div class ="cell normal" data-type ="normal">
                        <p>{}</p>
                    </div>
                    <div class ="cell normal" data-type ="normal">
                        <p>{}</p>
                    </div>
                </div>
            </div>
            <div  style="margin-top:40px;">
                <hr></hr>
                <h2 style="margin-top:40px;">TG report</h2>
                <p style="color:gray;"> Last monthly issue statistics </p>
                <div class ="columnLayout two-equal" data-layout="two-equal">
                    <div class ="cell normal" data-type ="normal">
                        <p>{}</p>
                    </div>
                    <div class ="cell normal" data-type ="normal">
                        <p>{}</p>
                    </div>
                </div>
            </div>
        </body>
        """.format(last_month_team, team_chart_line, tg_table, tg_pie)

    return html


def makeLinkChat(mySingleId):
    """Returns <a> tag with href from single ID"""
    info_link = "mysingleim://%s"
    return r"<a target='_blank'  href='%s'>%s</a>" % (info_link % mySingleId, mySingleId)


def makeLinkChatGroup(listID):
    """Returns <a> tag with href from single ID"""
    strListID = ""
    for i in range(0, len(listID)):
        strListID += str(listID[i]) + ';'
    info_link = "mysingleim://%s"
    return r"<a style='font-size: 12px; font-style: normal;' target='_blank'  href='%s'>%s</a>" % (
    info_link % strListID, "<br />Chat")


def makeLinkPLM(PLMCaseCode):
    """Returns <a> tag with href from mysingleID"""
    PLMLink = r"http://splm.sec.samsung.net/wl/tqm/defect/defectreg/getDefectCodeSearch.do?defectCode=%s"
    return r"<a href='%s'>%s</a>" % (PLMLink % PLMCaseCode, PLMCaseCode)


def make_link_chat(single_id, text):
    """Returns <a> tag with href from single ID"""
    info_link = "mysingleim://%s"
    return r"<a target='_blank'  href='%s'>%s</a>" % (info_link % single_id, text)


def make_link_jira(jira_key):
    jira_link = r"http://mobilerndhub.sec.samsung.net/its/browse/%s"
    return r"<a href='%s'>%s</a>" % (jira_link % jira_key, jira_key)


def make_link_jira_with_summary(jira_key, text):
    jira_link = r"http://mobilerndhub.sec.samsung.net/its/browse/%s"
    return r"<a href='%s'>%s</a>" % (jira_link % jira_key, text)


def make_img_jira(link):
    return r"<img src='%s' class='icon'>" % link


def make_status_jira(text):
    if text.lower() == 'new':
        return r"<span class='aui-lozenge aui-lozenge-subtle aui-lozenge-complete'>%s</span>" % text
    else:
        return r"<span class='aui-lozenge aui-lozenge-subtle aui-lozenge-current'>%s</span>" % text


def removeSpecialChar(string):
    special_char = {r'–': '&ndash;', r'—': '&mdash;', r'¡': '&iexcl;', r'¿': '&iquest;', r'"': '&quot;',
                    r'“': '&ldquo;',
                    r'”': '&rdquo;', r'‘': '&lsquo;', r'’': '&rsquo;', r'«': '&laquo;', r'»': '&raquo;', r'¢': '&cent;',
                    r'©': '&copy;', r'÷': '&divide', r'>': '&gt;', r'<': '&lt;', r'µ': '&micro;', r'·': '&middot',
                    r'¶': '&para;', r'±': '&plusmn', r'€': '&euro;', r'£': '&pound;', r'®': '&reg;', r'§': '&sect;',
                    r'™': '&trade;', r'¥': '&yen;', r'°': '&deg', r'á': '&aacute;', r'Á': '&Aacute;', r'à': '&agrave;',
                    r'À': '&Agrave;', r'â': '&acirc;', r'Â': '&Acirc;', r'å': '&aring;', r'Å': '&Aring;'}

    string = string.replace('&', '&amp;')
    for i in special_char:
        string = string.replace(i, special_char[i])

    return string


def create_isssue_owner(owner_list):
    html = "<head> \n </head> \n <body> \n <div> \n <p>"

    for i in owner_list:
        key = get_user_key(i)
        html += '<ac:link><ri:user ri:userkey="%s" /></ac:link>' % key
        html += ", "

    html += "</p> \n </div> \n </body>"
    return html


def get_user_key(user_name):
    request_data = requests.get("http://mobilerndhub.sec.samsung.net/wiki/rest/api/user?username=%s" % user_name,
                                auth=(user, pw))
    return request_data.json()['userKey']


def get_all_data_jira_task_list(project_key):
    # Query data with in 3 month
    jql_query = "project = %s and status not in (resolved, cancelled) and created > startOfMonth(-2) order by " \
                "created desc" % project_key
    max_result = 1000

    params = {
        "jql": jql_query,
        "startAt": 0,
        "maxResults": max_result,
        "fields": [
            "key",
            "summary",
            "issuetype",
            "created",
            "duedate",
            "assignee",
            "priority",
            "status"
        ]
    }

    url_query = 'http://mobilerndhub.sec.samsung.net/its/rest/api/2/search'
    data_task_list_json = requests.get(url_query, params=params, auth=(user, pw))

    list_all_task = json.loads(data_task_list_json.text)
    return list_all_task['issues']


def convert_date_time(date_time):
    date_time = datetime.datetime.strptime(date_time, "%Y-%m-%d").strftime("%b %d, %Y")
    return date_time


def get_data_jira_task_list_by_team(all_data_jira_task_list, member_id_list):
    num_of_jira_task_by_team = {}
    info_detail_jira_task = []
    number_of_jira_task_by_member = {}
    data_jira_task_for_pie_chart = [["", 'Jira Tasks'], ['Done', 0], ['NEW',  0], ["In Progress", 0]]

    for key, value in member_id_list.items():
        num_of_jira_task_by_team[key] = [0, 0]  # [open, in progress]

    for task_info in all_data_jira_task_list:
        summary = task_info['fields']['summary']

        if not summary.startswith('[Automatic]'):
            due_date = task_info['fields']['duedate']
            created = task_info['fields']['created'][:10]
            # created = convert_date_time(task_info['fields']['created'][:10])

            if due_date is None:
                due_date = ''
            # else:
            #     due_date = convert_date_time(due_date)

            single_id = task_info['fields']['assignee']['key']
            team = ""

            try:
                number_of_jira_task_by_member[single_id] = number_of_jira_task_by_member[single_id] + 1
            except KeyError:
                number_of_jira_task_by_member[single_id] = 1

            status_jira = task_info['fields']['status']['name'].lower()

            if status_jira == 'in progress':
                data_jira_task_for_pie_chart[3][1] += 1
            elif status_jira == 'new':
                data_jira_task_for_pie_chart[2][1] += 1
            else:
                data_jira_task_for_pie_chart[1][1] += 1

            if status_jira == 'in progress' or status_jira == 'new':
                for key, value in member_id_list.items():
                    if single_id in value:
                        team = key
                        if status_jira == 'in progress':
                            num_of_jira_task_by_team[key][1] = num_of_jira_task_by_team[key][1] + 1
                        elif status_jira == 'new':
                            num_of_jira_task_by_team[key][0] = num_of_jira_task_by_team[key][0] + 1
                        break

                if team.lower() == 'smart_manager':
                    team = 'Device Maintenance'
                elif team.lower() == 'myfile_voicer':
                    team = 'MyFile & Recorder'


                info = [
                    make_link_jira(task_info['key']),
                    summary,
                    make_img_jira(task_info['fields']['issuetype']['iconUrl']),
                    created,
                    due_date,
                    make_link_chat(single_id, task_info['fields']['assignee']['displayName']),
                    team,
                    make_img_jira(task_info['fields']['priority']['iconUrl']),
                    make_status_jira(task_info['fields']['status']['name'])
                ]

                info_detail_jira_task.append(info)

    data_chart_pie_jira = 'var dataChartPieJira = ' + str(data_jira_task_for_pie_chart) + '; \n'

    return num_of_jira_task_by_team, info_detail_jira_task, number_of_jira_task_by_member, data_chart_pie_jira
