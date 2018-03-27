from apscheduler.schedulers.blocking import BlockingScheduler
import AutoDownloadPLM as AutoPLM
import WeeklyReport as Report
from datetime import date, datetime
import WikiSubmit as Wiki
import V2_UrgentTable as Urgent
import V2_Chart as Chart
import ReadDataFromExcel as DataEXL
import utils


def create_urgent_owner():
    issue_owner = Wiki.create_isssue_owner(DataEXL.get_urgent_owner())
    html = """<div style="margin-top:40px;">
                <h2 >Issue Owners</h2>
                <p>{}</p>
            </div>""".format(issue_owner)
    return html


def create_data_jira_task_chart(data):
    list_jira = []
    for team, value in data.items():
        list_jira.append([team] + value)

    data_out = "var jira_chart_input = " + str([['', 'Jira task open', 'Jira task in progress']]
                                               + list_jira) + "; \n"
    return data_out


def create_data_work_chart(amount_work_of_member, list_member_id):
    mem_no_issue_tg = dict.fromkeys(DataEXL.get_list_tg(), 0)
    mem_no_issue_team = dict.fromkeys(DataEXL.get_list_team(), 0)

    for team, members in list_member_id.items():
        for member in members:
            if amount_work_of_member[member] == 0:
                mem_no_issue_team[team] += 1

    for tg, teams in DataEXL.get_list_team_of_tg().items():
        for team in teams:
            mem_no_issue_tg[tg] += mem_no_issue_team[team]

    mem_no_issue_team['type_data'] = 'viewTeam'
    mem_no_issue_tg['type_data'] = 'viewTG'
    chart_pie = [mem_no_issue_team, mem_no_issue_tg]
    data_chart = 'var data_pie_chart = ' + str(chart_pie) + '; \n'
    return data_chart


def create_data_top_10_member_chart(num_issue_member, num_task_member):
    list_data = []
    list_id_and_name_member = DataEXL.get_list_id_and_name_member()
    for id_member, num_issue in num_issue_member.items():
        name_member = list_id_and_name_member[id_member]
        try:
            list_data.append([name_member, num_issue, num_task_member[id_member]])
        except KeyError:
            list_data.append([name_member, num_issue, 0])

    list_data.sort(key=lambda x: x[1] + x[2], reverse=True)
    data_chart = 'var data_top_10_member_chart = ' \
                 + str([['Member', 'PLM issue', 'Jira task']] + list_data[:10]) + '; \n'
    return data_chart


def analysis_and_submit_wiki():
    """ Main function submit to wiki """
    pageTitle = "BA - Daily Report Test"
    main_plm_open, main_sub_chart, data_table_long_pending, \
    data_pending_chart, data_pending_main_sub, \
    urgent_by_prj, team_list, all_member_id, member_id_list,\
    data_detail_issue_open_by_team, dict_info_issues_open_by_team, num_issue_of_member = DataEXL.issue_analysis()

    # get base template file
    base_html = open('v2_template/home.html', 'r').read()
    menu_html = open('v2_template/menu.html', 'r').read()
    style_css = open('v2_template/style.css', 'r').read()
    java_script = open('v2_template/javascript.js', 'r').read()

    project_name_query = "SVMCBAP"
    urgent_html = Urgent.getUrgentTableCode()

    data_issue_team = "var dataTotalTeam = " + str(data_detail_issue_open_by_team) + "; \n"
    data_info_issues_open_by_team = "var dataInfoIssueTeam = " + str(dict_info_issues_open_by_team) + "; \n"

    all_data_jira_task_list = Wiki.get_all_data_jira_task_list(project_name_query)
    num_of_jira_task_by_team, info_detail_jira_task, number_of_jira_task_by_member, data_chart_pie_jira \
        = Wiki.get_data_jira_task_list_by_team(all_data_jira_task_list, member_id_list)

    for single_id, num_task in number_of_jira_task_by_member.items():
        num_issue_of_member[single_id] += num_task

    data_chart_pie = create_data_work_chart(num_issue_of_member, member_id_list)
    data_top_10_member_chart = create_data_top_10_member_chart(num_issue_of_member, number_of_jira_task_by_member)

    # create data for table and chart
    data_chart_urgent = Chart.create_data_urgent()
    data_table_urgent = "var dataTableUrgent = " + str(urgent_by_prj) + "; \n"
    data_table_pending_main = "var dataTablePending = " + str(data_table_long_pending) + "; \n"

    data_table_jira = data_out = "var dataJiraTask = " + str(info_detail_jira_task) + "; \n"
    data_jira_task = create_data_jira_task_chart(num_of_jira_task_by_team)
    data_summary_plm = Chart.create_data_plm_jira_issue(main_plm_open, num_of_jira_task_by_team)

    # create data for chart change data
    data_main_sub_summary = "var tab_pending_main_sub = " + str(main_sub_chart) + '; \n'
    data_pending_57days = "var tab_pending_57days = " + str(data_pending_chart) + '; \n'
    data_chart_lp_main_sub = "var chartMainSubLongPending = " + str(data_pending_main_sub) + '; \n'
    data_for_two_chart = data_main_sub_summary + data_pending_57days

    data_for_js = data_chart_urgent + data_table_pending_main + data_chart_lp_main_sub + data_summary_plm +\
                  data_jira_task + data_table_jira + data_table_urgent + data_info_issues_open_by_team + data_issue_team

    java_script = java_script.replace('<!--position_data_chart-->', data_for_js)
    java_script = java_script.replace('<!--position_data_two_chart-->', data_for_two_chart)
    java_script = java_script.replace('<!--data_chart_pie_member_no_issue-->', data_chart_pie)
    java_script = java_script.replace('<!--data_top_10_member_chart-->', data_top_10_member_chart)
    java_script = java_script.replace('<!--data_chart_pie_jira_task-->', data_chart_pie_jira)

    base_html = base_html.replace('<!--position_insert_menu-->', menu_html)
    base_html = base_html.replace('<!--position_urgent_table-->', urgent_html)
    base_html += style_css
    base_html += java_script

    html = '''<ac:structured-macro ac:name="html" ac:schema-version="1" ac:macro-id="16042b08-1210-47fd-9606-052492f11da4"><ac:plain-text-body><![CDATA['''
    html += base_html + ''']]></ac:plain-text-body></ac:structured-macro>'''
    Wiki.updateWiki(pageTitle, html)


def run_daily_report():
    """ run download excel and submit daily report """
    if check_time_report():
        user_name = utils.open_file(".plm")[0]
        password = utils.open_file(".plm")[1]

        AutoPLM.AutoDownloadPLM(user_name, password)
        analysis_and_submit_wiki()

def check_time_report():
    """ check time to report during 8h30 - 18h """
    time_now = datetime.now().time()
    time_start = time_now.replace(hour=7, minute=30)
    time_end = time_now.replace(hour=22, minute=30)
    if time_start < time_now < time_end:
        return True
    else:
        return False

if __name__ == "__main__":
    run_daily_report()
    today = date.today()
    # today = datetime.strptime('2018-01-01', '%Y-%m-%d').date()
    weekDay = today.strftime("%A")
    if weekDay == 'Monday':
        print("start submit weekly report")
        Report.analysis_week_month(today, "Weekly")

    today_date = today.strftime("%d")
    if today_date == '01':
        print("start submit monthly report")
        Report.analysis_week_month(today, "Monthly")

    print("start schedule task")
    scheduler = BlockingScheduler()
    scheduler.add_job(run_daily_report, 'interval', minutes=30)
    scheduler.start()