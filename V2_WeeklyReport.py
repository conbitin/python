from datetime import datetime, timedelta
import pandas as pd
import numpy as np
import collections
import WikiSubmit as Wiki
import ReadDataFromExcel as ReadExcel
from DatabaseHelper import DatabaseHelper


def analysis_week_month(today_date, type_report):
    """ analysis data for weekly/monthly report """
    # create layout weekly report
    base_html = open('v2_template/home_weekly.html', 'r').read()
    style_css = open('v2_template/style.css', 'r').read()
    java_script = open('v2_template/js_weekly.js', 'r').read()

    # start analysis data
    df = ReadExcel.get_data_frame()
    table_name_weekly = "WeeklyReport"
    table_name_monthly = "MonthlyReport"

    list_team = ReadExcel.get_list_team()
    list_team_of_tg = ReadExcel.get_list_team_of_tg()
    member_id_by_team = ReadExcel.get_member_id_of_team()

    # delete some column no need
    column_delete = ['Case Code', 'Title', 'Problem Type', 'Occurr. Freq.(Detail)', 'PPT (Hour)',
                     'PPT', 'Cause', 'Countermeasure', '# of Reject', 'Resolve confirmer', 'Reproduction Route',
                     'Problem', 'Resolve Charge', 'Occurr. Freq.', 'Occurr. Stg.', 'Resolve incharge ID', 'Priority']

    db = DatabaseHelper(list_team)
    for column in column_delete:
        if column in df.columns:
            df = df.drop(column, axis=1)

    # convert Resolve Date to string and substring get format yyyy-mm-dd
    register_date = df['Registered Date'].values
    resolve_date = df['Resolve Date'].values
    resolve_confirm_date = df['Resloved Confirm Date'].values

    list_register_date = [str(item)[:10] for item in register_date]
    list_resolve_date = [str(item)[:10] for item in resolve_date]
    list_confirm_date = [str(item)[:10] for item in resolve_confirm_date]
    df_date_header = ['Date']

    df_register_date = pd.DataFrame(list_register_date, columns=df_date_header)
    df_resolve_date = pd.DataFrame(list_resolve_date, columns=df_date_header)
    df_confirm_date = pd.DataFrame(list_confirm_date, columns=df_date_header)
    df['Registered Date'] = df_register_date['Date'].values
    df['Resolve Date'] = df_resolve_date['Date'].values
    df['Resloved Confirm Date'] = df_confirm_date['Date'].values

    if type_report == "Weekly":
        # get dataFrame issue by week
        days_last_week = []
        for i in range(1, 8):
            day = today_date - timedelta(i)
            days_last_week.append(day.strftime('%Y-%m-%d'))
        df_register_lw = df[df['Registered Date'].isin(days_last_week)]

        weeklyIssue = {}
        weeklyProjectTG = list(set(df_register_lw["Pjt. Name"].tolist()))
        df_report = pd.DataFrame(weeklyIssue, index=weeklyProjectTG, columns=list_team)

        # get list project weekly by team
        for team, member in member_id_by_team.items():
            df_team_issue_lw = df_register_lw[df_register_lw['Manager ID'].isin(member)]
            listProject = df_team_issue_lw["Pjt. Name"].tolist()  # get list project of team by weekly
            counterProject = collections.Counter(listProject)
            team = team.replace(' (clock,memo,cal)', '').replace(' &amp; ', '_').replace(' ', '_')
            for project, numIssue in counterProject.items():
                df_report.at[project, team] = numIssue

        current_week = today_date.isocalendar()[1]
        last_week_number = current_week - 1
        if last_week_number == 0:
            last_week_number = 52
        df_report['Total'] = df_report.sum(axis=1)
        df_report['Week'] = last_week_number  # save number week to report

        # save dataFrame to database
        db.delete_table_last_week()  # delete all old data after insert new
        db.insert_dataFrame(df_report)

        # get data issue last week from dataBase
        data_last_week = db.get_data_last_week(last_week_number)
        df_last_week = pd.DataFrame(data_last_week, dtype=int)
        df_last_week = df_last_week.fillna(0)

        title_chart_stack = ['Project'] + list_team
        data_last_week = df_last_week.values.tolist()
        data_last_week.sort(key=lambda x: sum(item for item in x[1:]), reverse=True)
        data_last_week.insert(0, title_chart_stack)

        # insert summary last week to table WeeklyReport
        summary_urgent, summary_issue_lw,\
        total_issue_lw, data_chart_bar_lw = create_data_latest(last_week_number, df_register_lw, member_id_by_team)

        if db.checkWeekExists(last_week_number):
            db.updateDataWeekly(total_issue_lw)
        else:
            db.insertDataWeekly(total_issue_lw)
        # get data from database
        line_chart_weekly, list_team_name = getDataLineChart(db, table_name_weekly)
        tg_report_table_week, tg_chart_pie_week = data_latest_tg(list_team_of_tg, summary_issue_lw, summary_urgent)

        line_chart_data = "var dataChartLine = " + str(line_chart_weekly) + "; \n"
        line_chart_title = "var titleChartLine = " + str(list_team_name) + "; \n"
        line_chart = line_chart_title + line_chart_data

        type_report = "var type_report = 'weeklyReport'; \n"
        data_chart_bar = "var dataChartBar = " + str(data_chart_bar_lw) + "; \n"
        data_tg_report = "var dataTgReport = " + str(tg_report_table_week) + "; \n"
        data_chart_pie = "var dataChartPie = " + str(tg_chart_pie_week) + "; \n"
        data_chart_stack_lw = "var dataChartStack = " + str(data_last_week) + "; \n"

        data_for_js = type_report + data_chart_bar + data_tg_report + data_chart_pie
        java_script = java_script.replace('<!--position_data_chart-->', data_for_js)
        java_script = java_script.replace('<!--data_for_line_chart_weekly-->', line_chart)
        java_script = java_script.replace('<!--data_issue_team_by_prj-->', data_chart_stack_lw)

        base_html = base_html.replace('<!--number-of-week-->', str(last_week_number))
        base_html += style_css
        base_html += java_script
        html = '''<ac:structured-macro ac:name="html" ac:schema-version="1" ac:macro-id="16042b08-1210-47fd-9606-052492f11da4"><ac:plain-text-body><![CDATA['''
        html += base_html + ''']]></ac:plain-text-body></ac:structured-macro>'''
        pageTitle = "BA - Weekly Report"
        Wiki.updateWiki(pageTitle, html)

    elif type_report == "Monthly":
        # ************* submit monthly report *******************
        this_month = today_date.replace(day=1)
        last_month = this_month - timedelta(days=1)
        last_month_str = last_month.strftime('%m%Y')
        last_month_int = int(last_month_str)
        summary_month, summary_total_month, chart_month = create_data_latest(last_month_int, df, member_id_by_team,
                                                                             type_data='Month')
        if db.checkMonthExists(last_month_int):
            db.updateDataMonthly(summary_total_month)
        else:
            db.insertDataMonthly(summary_total_month)
        line_chart_monthly, list_team_name = getDataLineChart(db, table_name_monthly)
        tg_report_table_month, tg_chart_pie_month = data_latest_tg(list_team_of_tg, summary_month, type_data='Month')

        line_chart_data = "var dataChartLine = " + str(line_chart_monthly) + "; \n"
        line_chart_title = "var titleChartLine = " + str(list_team_name) + "; \n"
        line_chart = line_chart_title + line_chart_data

        type_report = "var type_report = 'monthlyReport'; \n"
        num_of_month_rp = "var MonthReport = '" + str(last_month.strftime('%b-%y')) + "'; \n"
        data_chart_bar = "var dataChartBar = " + str(chart_month) + "; \n"
        data_tg_report = "var dataTgReport = " + str(tg_report_table_month) + "; \n"
        data_chart_pie = "var dataChartPie = " + str(tg_chart_pie_month) + "; \n"

        data_for_js = type_report + data_chart_bar + data_tg_report + data_chart_pie + num_of_month_rp
        java_script = java_script.replace('<!--position_data_chart-->', data_for_js)
        java_script = java_script.replace('<!--data_for_line_chart_monthly-->', line_chart)

        base_html += style_css
        base_html += java_script
        html = '''<ac:structured-macro ac:name="html" ac:schema-version="1" ac:macro-id="16042b08-1210-47fd-9606-052492f11da4"><ac:plain-text-body><![CDATA['''
        html += base_html + ''']]></ac:plain-text-body></ac:structured-macro>'''
        pageTitle = "BA - Monthly report"
        Wiki.updateWiki(pageTitle, html)


def getDataLineChart(db, table_name):
    """ create data line chart for monthly and weekly """
    data_report = []
    list_column_name = db.getColumnName(table_name)
    total_data = db.getData(table_name)
    new_data = list(zip(*total_data))
    for value in new_data:
        project_data = [data for data in value]
        data_report.append(project_data)

    # get 4 latest week
    i = 0
    number_of_week = 4
    for data in data_report:
        index = len(data) - number_of_week
        del data[0:index]
        data_report[i] = data
        i += 1

    list_team_name = []
    list_week_report = []
    list_team_report = []
    for team, value in zip(list_column_name, data_report):
        if team == 'Week' or team == 'Month':
            list_week_report.append(value)
        else:
            # change all None to 0
            values = np.array(value)
            values[values == None] = 0
            list_team_report.append(values.tolist())
            list_team_name.append(team)

    list_temp_data = list_week_report + list_team_report
    data_line_chart = str(np.transpose(list_temp_data).tolist())
    return data_line_chart, list_team_name


def create_data_latest(time_report, data_frame, member_id_by_team, type_data='Week'):
    """create data for last week / last month report"""
    summary_issue = {}
    summary_urgent = {}
    if type_data == 'Week':
        summary_total_issue = [time_report]
    else:
        this_year = datetime.now().year
        summary_total_issue = [time_report * 10000 + this_year]
    data_chart_bar = [['', 'Open Issue', 'Resolve Issue']]
    for team, member in member_id_by_team.items():
        df_total_issue_team = data_frame[data_frame['Manager ID'].isin(member)]
        df_team_resolved = df_total_issue_team[~df_total_issue_team['Resolve Date'].isin(['NaT'])]
        df_team_close = df_total_issue_team[~df_total_issue_team['Resloved Confirm Date'].isin(['NaT'])]

        total_issue_team = len(df_total_issue_team.index)
        total_resolve_team = len(df_team_resolved.index)
        issue_close = len(df_team_close.index)
        issue_open = total_issue_team - total_resolve_team
        issue_resolved = total_resolve_team - issue_close
        team_name = team.replace(' (clock,memo,cal)', '').replace(' &amp; ', '_').replace(' ', '_')
        summary_issue[team_name] = [issue_open, issue_resolved, total_issue_team]
        summary_total_issue.append(total_issue_team)

        data_chart_bar.append([team, issue_open, issue_resolved])
        summary_urgent[team_name] = count_urgent_lw(df_total_issue_team)

    if type_data == 'Week':
        return summary_urgent, summary_issue, summary_total_issue, data_chart_bar
    else:
        return summary_issue, summary_total_issue, data_chart_bar


def count_urgent_lw(data_frame):
    """ count num issue urgent last week """
    list_urgent = ReadExcel.get_list_urgent_prj()
    Urgent_project = [item.upper() for item in list_urgent]
    dict_row_value = data_frame.to_dict(orient='records')
    close_urgent = 0
    urgent_team = [0, 0, 0]  # [open, resolve, total]
    for index in range(0, len(dict_row_value)):
        project_name = dict_row_value[index]['Dev. Mdl. Name/Item Name']
        if ReadExcel.is_urgent_issue(project_name, Urgent_project) != '':
            urgent_team[2] += 1
            if dict_row_value[index]['Resolve Date'] == 'NaT':
                urgent_team[0] += 1
            elif dict_row_value[index]['Resloved Confirm Date'] != 'NaT':
                close_urgent += 1

    urgent_team[1] = urgent_team[2] - (urgent_team[0] + close_urgent)
    return urgent_team


def data_latest_tg(list_team_of_tg, summary_issue_lw, summary_urgent={}, type_data='Week'):
    """create data weekly report for TG """
    tg_report_table = []
    weekly_tg_chart_pie = [['', 'issue by TG']]
    team_of_tg = covert_team_tg(list_team_of_tg)
    for tg_name in team_of_tg:
        tg_report_lw = []
        tg_issue = [0, 0, 0]  # tg_issue[open, resolve, total]
        tg_urgent = [0, 0, 0]  # tg_urgent[open, resolve, total]
        for team in team_of_tg[tg_name]:
            for i in range(0, 3):
                tg_issue[i] += summary_issue_lw[team][i]
                if type_data == 'Week':
                    tg_urgent[i] += summary_urgent[team][i]
        tg_report_lw.append(tg_name)
        tg_report_lw.append(str(int(tg_issue[2])) + '/' + str(int(tg_issue[1])) + '/' + str(int(tg_issue[0])))
        tg_report_table.append(tg_report_lw)
        # create data for pie chart
        if type_data == 'Week':
            weekly_tg_chart_pie.append([tg_name, tg_urgent[2]])
            tg_report_lw.append(str(int(tg_urgent[2])) + '/' + str(int(tg_urgent[1])) + '/' + str(int(tg_urgent[0])))
        else:
            weekly_tg_chart_pie.append([tg_name, tg_issue[2]])

    return tg_report_table, weekly_tg_chart_pie


def covert_team_tg(list_team_tg):
    """get list team in TG; remove duplicate team"""
    for tg in list_team_tg:
        list_team_tg[tg] = list(set(list_team_tg[tg]))
        list_team_tg[tg] = [team_name.replace(' (clock,memo,cal)', '').replace(' &amp; ', '_').replace(' ', '_') for
                            team_name in list_team_tg[tg]]
    return list_team_tg

