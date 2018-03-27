import xlwings as xw
import pandas as pd
import re
import utils
import subprocess
import pythoncom
from datetime import date, datetime, timedelta
from confluence import Api
import WikiSubmit as Wiki

LIST_MEMBER = r'sample input\\Member_Organization.xlsx'
LIST_ISSUE = r'sample input\\DEFECT_LIST_Today_Basic.xls'

wiki_url = "http://mobilerndhub.sec.samsung.net/wiki/"  # Samsung DPI Wiki entry url
space = "SVMC"
pageUrgentPrjTitle = "BA - Urgent projects"
user = utils.open_file(".wiki")[0]
pw = utils.open_file(".wiki")[1]
api = Api(wiki_url, user, pw)
PRJ_MACRO_TAG = '</ac:structured-macro>'
FIRST_PRJ_TAG = '<ac:parameter ac:name="title">'
LAST_PRJ_TAG = '</ac:parameter>'

list_team = []
list_tg = []
list_team_of_tg = {}
list_id_and_name_member = {}
urgent_issue_prj_count = {}
dfExcel = pd.DataFrame()

owner_issue_urgent = []
owner_issue_pending = []


def get_urgent_project():
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


def team_member(id_member, member_id_by_team):
    for key, value in member_id_by_team.items():
        if id_member in value:
            return key


def is_pending_issue(ppt, time_pending_standard):
    """ check is long pending issue """
    int_list = [int(s) for s in re.findall('\\d+', ppt)]
    if int_list[0] > time_pending_standard:
        return True
    elif int_list[0] == time_pending_standard and (int_list[1] > 0 or int_list[2] > 0):
        return True
    else:
        return False


def ppt_min(issue):
    """ get time pending issue """
    ppt_index = 2
    int_list = [int(s) for s in re.findall('\\d+', issue[ppt_index])]
    return int_list[0] * 24 * 60 + int_list[1] * 60 + int_list[2]

def get_list_single_id_member():
    pythoncom.CoInitialize()
    try:
        memberBook = xw.Book(LIST_MEMBER)
        memberSheet = memberBook.sheets("My Organization Member")
        tableValue = memberSheet.range('A2').expand('table').value
        length = len(memberSheet.range('A2').expand('right').value)

        dataConvert = pd.DataFrame(tableValue, columns=memberSheet.range('A1').expand('right').value[:length])
        list_single_id_member = dataConvert["mySingle"].tolist()
        list_single_id_member = [single_id.lower() for single_id in list_single_id_member]
        memberBook.close()
        return list_single_id_member

    except Exception as e:
        print(e)

def get_team_info():
    """ get total info team of part """
    pythoncom.CoInitialize()
    try:
        memberBook = xw.Book(LIST_MEMBER)
        memberSheet = memberBook.sheets("My Organization Member")
        tableValue = memberSheet.range('A2').expand('table').value
        length = len(memberSheet.range('A2').expand('right').value)

        dataConvert = pd.DataFrame(tableValue, columns=memberSheet.range('A1').expand('right').value[:length])
        list_tg = list(set(dataConvert["SVMC TG"].tolist()))
        list_team = list(set(dataConvert["SVMC Sub TG"].tolist()))

        list_single_id_member = dataConvert["mySingle"].tolist()
        list_single_id_member = [single_id.lower() for single_id in list_single_id_member]
        list_name_member = dataConvert["Full Name"].tolist()

        # remove all space start-end element
        list_tg = list(map(str.strip, list_tg))
        list_team = list(map(str.strip, list_team))
        list_team_name = []

        all_member_id_part = dataConvert["mySingle"].tolist()
        numOfMember = len(all_member_id_part)

        list_team_of_tg = {}
        member_name_by_team = {}
        member_id_by_team = {}
        list_id_and_name_member = dict(zip(list_single_id_member, list_name_member))
        # create list team of TG
        for tg in list_tg:
            dataTG = dataConvert.loc[dataConvert["SVMC TG"] == tg]
            teamList = dataTG["SVMC Sub TG"].tolist()
            teamList = [team_name.replace(' (clock,memo,cal)', '').replace(' &amp; ', '_').replace(' ', '_')
                        for team_name in list(set(teamList))]
            list_team_of_tg[tg] = teamList

        # create list member Name/ID of Team
        for team in list_team:
            team_name = team.replace(' (clock,memo,cal)', '').replace(' &amp; ', '_').replace(' ', '_')
            list_team_name.append(team_name)
            dataTeam = dataConvert.loc[dataConvert["SVMC Sub TG"] == team]

            list_member_name = dataTeam["Full Name"].tolist()
            member_of_team = [name.title() for name in
                              list_member_name]  # name = proper case (John Smith) , id = lower case
            member_name_by_team[team_name] = member_of_team
            team_member_id = dataTeam["mySingle"].tolist()
            team_member_id = [id.lower() for id in team_member_id]
            member_id_by_team[team_name] = team_member_id

        memberBook.close()
        return list_tg, list_team_of_tg, list_team_name, numOfMember, \
               member_name_by_team, member_id_by_team, all_member_id_part, list_id_and_name_member
    except Exception as e:
        print(e)


def is_urgent_issue(name, urgent_project):
    """ check is urgent issue """
    for prj in urgent_project:
        if prj in name:
            return prj
    return ''


def folder_issue(data, index):
    """ return to folder of issue """
    if data['Occurr. Stg.'][index] in str(['DI', 'PI', 'DV', 'PV', 'PR', 'SR', 'MS']):
        return 'Main'
    elif 'MR' in data['Pjt. Name'][index]:
        return 'MR'
    else:
        return 'Sub'


def summary_analysis(input_data):
    """ convert data to summary report """
    chart_summary = {}
    for team in input_data:
        open_issue = input_data[team][0]  # just get open issue
        chart_summary[team] = open_issue

    return chart_summary


def summary_analysis_active(input_data, type_data):
    """ create data include open-active """
    issue_open = {'type_chart': type_data, 'type_data': 'open'}
    issue_active = {'type_chart': type_data, 'type_data': 'active'}
    for team in input_data:
        active_issue = input_data[team][0] + input_data[team][2]
        issue_open[team] = input_data[team][0]
        issue_active[team] = active_issue

    summary_open_active = [issue_open, issue_active]
    return summary_open_active


def read_excel_issue(file_name):
    """ read issue from excel to dataFrame """
    issueBook = xw.Book(file_name)
    issueSheet = issueBook.sheets("DEFECT")
    pd_header = issueSheet.range('A3').expand('right').value
    num_row = len(issueSheet.range('A3').expand('down').value) + 2
    print("number of row: ", num_row)
    if num_row < 1000:
        tableValue = issueSheet.range('A4').expand('table').value
        data = pd.DataFrame(tableValue, columns=pd_header)
    else:
        num_column = len(pd_header)
        range_table1 = 'A4:'
        range_table2 = 'A'
        last_column_name = chr(64 + num_column)

        split_point = int(num_row / 2)
        range_table1 += last_column_name + str(split_point)
        range_table2 += str(split_point + 1) + ':' + last_column_name + str(num_row)

        table_value1 = issueSheet.range(range_table1).value
        table_value2 = issueSheet.range(range_table2).value

        df1 = pd.DataFrame(table_value1, columns=pd_header)
        df2 = pd.DataFrame(table_value2, columns=pd_header)

        data = df1.append(df2, ignore_index=True)

    length_frame = len(data)
    data = data.drop_duplicates(subset=['Case Code'], keep='first')
    issueBook.close()
    return data, length_frame


def urgent_count_prj(prj):
    """ get total count detail urgent project """
    tempdict = urgent_issue_prj_count.get(prj, {})
    return sum(tempdict.values())


def urgent_count_all_prj():
    return urgent_issue_prj_count


def get_urgent_owner():
    """ return owner of issue urgent """
    return owner_issue_urgent


def get_data_frame():
    """ return to dataFrame read from excel """
    return dfExcel


def get_list_team_of_tg():
    return list_team_of_tg


def get_list_team():
    """ return to team name of part """
    return list_team


def get_list_tg():
    """ return to team name of part """
    return list_tg


def get_list_id_and_name_member():
    """ return to single_id and name member of part """
    return list_id_and_name_member


def check_daily_comment(name_owner, comments, reg_date):
    """ check urgent issue comment yesterday """
    register_date = datetime.strptime(reg_date, '%Y-%m-%d').date()
    today = date.today()
    yesterday = today - timedelta(1)
    title_owner = name_owner + '/SEV-Basic App. P:'

    if register_date == today:
        # if register today
        return "New"
    else:
        pos_cmt_latest = comments.find(title_owner)
        if pos_cmt_latest == -1:
            # issue has no comment
            return "No"
        else:
            # check date of comment
            position_start = pos_cmt_latest + len(title_owner)
            time_cmt_latest = comments[position_start: position_start + 10]
            date_cmt_latest = datetime.strptime(time_cmt_latest, '%Y.%m.%d').date()
            if date_cmt_latest == yesterday:
                return "Yes"
    return "No"


def num_date_comment(name_owner, comments, reg_date):
    """ get num of date no comment issue """
    register_date = datetime.strptime(reg_date, '%Y-%m-%d').date()
    today = date.today()
    title_owner = name_owner + '/SEV-Basic App. P:'
    if register_date == today:
        # if register today
        return "New"
    else:
        if not comments:
            # issue has no comment
            return "No Comment"
        pos_cmt_latest = comments.find(title_owner)
        if pos_cmt_latest == -1:
            # issue has no comment
            return "No Comment"
        else:
            # check date of comment
            position_start = pos_cmt_latest + len(title_owner)
            time_cmt_latest = comments[position_start: position_start + 10]
            date_cmt_latest = datetime.strptime(time_cmt_latest, '%Y.%m.%d').date()

            days_no_cmt = (today - date_cmt_latest).days
            if days_no_cmt == 0:
                return "Today"
            elif days_no_cmt == 1:
                return "1 Day"
            else:
                return str(days_no_cmt) + " Days"


def convert_dict_to_arr(dict_data):
    """ convert dict data to arr """
    out_data = [["", ""]]
    for key, value in dict_data.items():
        out_data.append([key, value])
    return out_data


def issue_analysis():
    """ analysis total issue report """
    long_pending_main = []
    long_pending_sub_mr = []

    pending_day = {}
    main_folder_issue = {}
    sub_mr_folder_issue = {}
    dict_num_of_issues_open_by_team = {}
    dict_num_of_issues_open_by_tg = {}
    dict_info_issues_open_by_team = {}
    dict_info_issues_open_by_tg = {}
    urgent_by_prj = {}
    pending_issue_amount_5 = {}
    pending_issue_amount_7 = {}
    pending_sub_mr_by_team = {}

    list_urgent = get_urgent_project()
    urgent_project = [item.upper() for item in list_urgent]
    list_key_prj = [item.lower() for item in list_urgent]
    for prj in list_key_prj:
        urgent_by_prj[prj] = []

    global urgent_issue_prj_count
    global owner_issue_urgent
    global owner_issue_pending
    global list_team
    global dfExcel
    global list_team_of_tg
    global list_team
    global list_tg
    global list_id_and_name_member

    list_tg, list_team_of_tg, list_team, numOfMember, \
    member_name_by_team, member_id_by_team, all_member_id_part, list_id_and_name_member = get_team_info()
    num_issue_of_member = dict.fromkeys(all_member_id_part, 0)

    for team in list_team:
        pending_day[team] = [0, 0, 0]
        pending_issue_amount_5[team] = 0
        pending_issue_amount_7[team] = 0
        pending_sub_mr_by_team[team] = 0
        main_folder_issue[team] = [0, 0, 0]  # [open, close, resolved]
        sub_mr_folder_issue[team] = [0, 0, 0]  # [open, close, resolved]
        dict_num_of_issues_open_by_team[team] = {}
        dict_info_issues_open_by_team[team] = []

    for tg in list_tg:
        dict_info_issues_open_by_tg[tg] = []
        dict_num_of_issues_open_by_tg[tg] = {}

    dfExcel, length_frame = read_excel_issue(LIST_ISSUE)

    pending_time_5 = 5
    pending_time_7 = 7
    for i in range(0, length_frame):
        try:
            id_member = dfExcel['Manager ID'][i]
        except:
            continue
        team = team_member(str(id_member), member_id_by_team)
        tg = team_member(str(team), list_team_of_tg)

        if team is not None:
            # count issue open
            if dfExcel['PPT'][i] != '-':
                issue = [dfExcel['Case Code'][i], dfExcel['Pjt. Name'][i],
                         dfExcel['PPT'][i], dfExcel['Title'][i],
                         dfExcel['Manager ID'][i], team, tg, dfExcel['Priority'][i]]

                issue[0] = Wiki.makeLinkPLM(issue[0])
                issue[4] = Wiki.makeLinkChat(issue[4])

                num_issue_of_member[id_member] += 1
                is_folder_issue = folder_issue(dfExcel, i)

                # - START: data chart pie for team, tg
                name_model = dfExcel['Dev. Mdl. Name/Item Name'][i]

                if 'SM-' not in name_model or is_folder_issue != 'Main':
                    name_model = 'Sub & MR'
                else:
                    pos = name_model.index('SM-')
                    name_model = name_model[pos + 3:]
                    name_model = name_model.split('_')[0]

                try:
                    dict_num_of_issues_open_by_team[team][name_model] = \
                        dict_num_of_issues_open_by_team[team][name_model] + 1
                except KeyError:
                    dict_num_of_issues_open_by_team[team][name_model] = 1

                try:
                    dict_num_of_issues_open_by_tg[tg][name_model] = \
                        dict_num_of_issues_open_by_tg[tg][name_model] + 1
                except KeyError:
                    dict_num_of_issues_open_by_tg[tg][name_model] = 1

                # get comment Of issue
                comment_issue = dfExcel['Comment'][i]
                name_of_owner = list_id_and_name_member[dfExcel['Manager ID'][i]]
                register_date = str(dfExcel['Registered Date'][i])[:10]
                team_issue_detail = issue + [num_date_comment(name_of_owner, comment_issue, register_date)]

                dict_info_issues_open_by_team[team].append(team_issue_detail)
                dict_info_issues_open_by_tg[tg].append(team_issue_detail)
                # - END: data chart pie for team, tg

                if is_folder_issue == 'Main':
                    main_folder_issue[team][0] += 1  # count open issue Main Folder
                    checkUrgent = is_urgent_issue(dfExcel['Dev. Mdl. Name/Item Name'][i], urgent_project).lower()
                    if checkUrgent != '':
                        if checkUrgent in urgent_issue_prj_count:
                            if team in urgent_issue_prj_count[checkUrgent]:
                                urgent_issue_prj_count[checkUrgent][team] = urgent_issue_prj_count[checkUrgent][
                                                                                team] + 1
                            else:
                                urgent_issue_prj_count[checkUrgent][team] = 1
                        else:
                            urgent_issue_prj_count[checkUrgent] = {}
                            urgent_issue_prj_count[checkUrgent][team] = 1

                        # get issue detail by project add Column Daily comment
                        urgent_issue = issue + [check_daily_comment(name_of_owner, comment_issue, register_date)]
                        urgent_by_prj[checkUrgent].append(urgent_issue)

                        if id_member not in owner_issue_urgent:
                            owner_issue_urgent.append(id_member)
                    # count issue long pending
                    pday = [int(s) for s in re.findall('\\d+', dfExcel['PPT'][i])]
                    for idx in range(3):
                        pending_day[team][idx] += pday[idx]
                    if is_pending_issue(dfExcel['PPT'][i], pending_time_5):
                        pending_issue_amount_5[team] += 1
                        long_pending_main.append(issue)
                        if id_member not in owner_issue_pending:
                            owner_issue_pending.append(id_member)
                    if is_pending_issue(dfExcel['PPT'][i], pending_time_7):
                        pending_issue_amount_7[team] += 1
                else:
                    # count OPEN issue MR and Sub folder
                    sub_mr_folder_issue[team][0] += 1
                    if is_pending_issue(dfExcel['PPT'][i], pending_time_7):
                        long_pending_sub_mr.append(issue)
                        pending_sub_mr_by_team[team] += 1
            else:
                # count sub and MR issue resolved / close
                if str(dfExcel['Resloved Confirm Date'][i]) == 'NaT':
                    if folder_issue(dfExcel, i) == 'Main':
                        main_folder_issue[team][2] += 1  # count issue RESOLVED not close in Main Folder
                    else:
                        sub_mr_folder_issue[team][2] += 1
                else:
                    # count issue CLOSE in Main Folder
                    if folder_issue(dfExcel, i) == 'Main':
                        main_folder_issue[team][1] += 1
                    else:
                        sub_mr_folder_issue[team][1] += 1

    long_pending_sub_mr[1:] = sorted(long_pending_sub_mr[1:], key=lambda x: -ppt_min(x))
    subprocess.call("taskkill /IM excel.exe")

    pending_result = {}

    for team in list_team:
        pending_result[team] = pending_day[team][0] + int(pending_day[team][1] / 24) + int(
            pending_day[team][2] / (24 * 60))

    # create data chart long pending Main/Sub in PLM Issue TAB
    data_pending_main_sub = {}
    data_pending_main_sub['main_5days'] = convert_dict_to_arr(pending_issue_amount_5)
    data_pending_main_sub['sub_7days'] = convert_dict_to_arr(pending_sub_mr_by_team)
    # create data table long pending Main 5days / Sub 7days
    data_table_long_pending = {}
    data_table_long_pending['main_5days'] = long_pending_main
    data_table_long_pending['sub_7days'] = long_pending_sub_mr

    pending_issue_amount_5['type_data'] = "fiveDays"
    pending_issue_amount_7['type_data'] = "sevenDays"
    num_pending_57days = [pending_issue_amount_5, pending_issue_amount_7]

    main_plm_open = summary_analysis(main_folder_issue)
    main_summary_bar = summary_analysis_active(main_folder_issue, 'main')
    sub_summary_bar = summary_analysis_active(sub_mr_folder_issue, 'sub')
    summary_data = main_summary_bar + sub_summary_bar

    data_detail_issue_open_by_team = {}
    dict_num_of_issues_open_by_team.update(dict_num_of_issues_open_by_tg)
    for key, value in dict_num_of_issues_open_by_team.items():
        info = []
        for k, v in value.items():
            info.append([k, v])
        data_detail_issue_open_by_team[key] = [['', 'Number issue']] + info

    dict_info_issues_open_by_team.update(dict_info_issues_open_by_tg)

    # add new item detail all urgent projects
    all_urgent_prj = []
    for key, values in urgent_by_prj.items():
        all_urgent_prj += values
    urgent_by_prj['detail_all_urgent'] = all_urgent_prj

    return main_plm_open, summary_data, data_table_long_pending, \
           num_pending_57days, data_pending_main_sub,\
           urgent_by_prj, list_team, all_member_id_part, member_id_by_team, \
           data_detail_issue_open_by_team, dict_info_issues_open_by_team, num_issue_of_member
