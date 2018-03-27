from datetime import datetime, timedelta
import pandas as pd
import collections
import subprocess

import WikiSubmit as Wiki
import ReadDataFromExcel as ReadExcel
from DatabaseHelper import DatabaseHelper


def analysis_week_month(today_date, type_report):
    df = ReadExcel.get_data_frame()
    table_name_weekly = "WeeklyReport"
    table_name_monthly = "MonthlyReport"

    listTG, list_team_of_tg, listTeam, numOfMember, list_memberName, list_memberID, all_member_id_list = ReadExcel.get_team_info()
    list_team = [team_name.replace(' (clock,memo,cal)', '').replace(' &amp; ', '_').replace(' ', '_') for team_name in
                 listTeam]
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
    # print(df)

    subprocess.call("taskkill /IM excel.exe")

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
        for team, member in list_memberID.items():
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
        # print(df_last_week)

        header_table = ['Project'] + list_team + ['Total']
        data_last_week = df_last_week.values.tolist()
        lw_report_table = [header_table] + data_last_week

        # insert summary last week to table WeeklyReport
        summary_urgent, summary_issue_lw, total_issue_lw, data_chart_bar_lw = create_data_latest(last_week_number,
                                                                                                 df_register_lw,
                                                                                                 list_memberID)
        if db.checkWeekExists(last_week_number):
            db.updateDataWeekly(total_issue_lw)
        else:
            db.insertDataWeekly(total_issue_lw)

        # get data from database
        line_chart_weekly = getDataLineChart(db, table_name_weekly)
        lw_report_table = convert_data_format(lw_report_table)
        tg_report_table_week, tg_chart_pie_week = data_latest_tg(list_team_of_tg, summary_issue_lw, summary_urgent)
        submit_weekly_rp(last_week_number, data_chart_bar_lw, lw_report_table, tg_report_table_week, tg_chart_pie_week,
                         line_chart_weekly)

    elif type_report == "Monthly":
        # ************* submit monthly report *******************
        today_str = today_date.strftime('%Y-%m-%d')
        month_num = int(today_str[5:7])
        if month_num == 1:
            month_num = 12
        else:
            month_num = month_num - 1

        summary_month, summary_total_month, chart_month = create_data_latest(month_num, df, list_memberID,
                                                                             type_data='Month')
        if db.checkMonthExists(month_num):
            db.updateDataMonthly(summary_total_month)
        else:
            db.insertDataMonthly(summary_total_month)

        line_chart_monthly = getDataLineChart(db, table_name_monthly)
        tg_report_table_month, tg_chart_pie_month = data_latest_tg(list_team_of_tg, summary_month, type_data='Month')
        submit_monthly_rp(month_num, chart_month, tg_report_table_month, tg_chart_pie_month, line_chart_monthly)


def getDataLineChart(db, table_name):
    """ create data line chart for monthly and weekly """
    data_report = []
    data_line_chart = []
    list_column_name = db.getColumnName(table_name)
    total_data = db.getData(table_name)
    new_data = list(zip(*total_data))
    for value in new_data:
        project_data = [data for data in value]
        data_report.append(project_data)

    # get 4 latest week
    i = 0
    for data in data_report:
        index = len(data) - 4
        del data[0:index]
        data_report[i] = data
        i += 1
    for team, value in zip(list_column_name, data_report):
        if team != 'Week' or team != 'Month':
            team = team + '(' + str(value[0]) + '-' + str(value[1]) + '-' + str(value[2]) + '-' + str(value[3]) + ')'
        data_line_chart.append([team] + value)

    return data_line_chart


def create_data_latest(last_week_number, data_frame, list_memberID, type_data='Week'):
    """create data for last week / last month report"""
    summary_issue = {}
    summary_urgent = {}
    summary_total_issue = [last_week_number]
    data_chart_bar = [['Team'], ['Open'], ['Resolve']]
    for team, member in list_memberID.items():
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

        team_title = team_name + ' (' + str(issue_open) + '-' + str(issue_resolved) + ')'
        data_chart_bar[0].append(team_title)
        data_chart_bar[1].append(issue_open)
        data_chart_bar[2].append(issue_resolved)

        summary_urgent[team_name] = count_urgent_lw(df_total_issue_team)

    if type_data == 'Week':
        return summary_urgent, summary_issue, summary_total_issue, data_chart_bar
    else:
        return summary_issue, summary_total_issue, data_chart_bar


def count_urgent_lw(data_frame):
    """ count num issue urgent last week """
    list_urgent = ReadExcel.get_urgent_project()
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
    if type_data == 'Week':
        tg_report_table = [["TG", "Total \n (Total/Resolved/Open)", "Urgent \n (Total/Resolved/Open)"]]
    else:
        tg_report_table = [["TG", "Total \n (Total/Resolved/Open)"]]
    team_of_tg = covert_team_tg(list_team_of_tg)
    tg_chart_pie_header = ['TG']
    tg_chart_pie_data = ['Total']
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
        tg_chart_pie_header.append(tg_name)
        if type_data == 'Week':
            tg_chart_pie_data.append(tg_urgent[2])
            tg_report_lw.append(str(int(tg_urgent[2])) + '/' + str(int(tg_urgent[1])) + '/' + str(int(tg_urgent[0])))
        else:
            tg_chart_pie_data.append(tg_issue[2])

    tg_chart_pie = [tg_chart_pie_header, tg_chart_pie_data]

    return tg_report_table, tg_chart_pie


def covert_team_tg(list_team_tg):
    """get list team in TG; remove duplicate team"""
    for tg in list_team_tg:
        list_team_tg[tg] = list(set(list_team_tg[tg]))
        list_team_tg[tg] = [team_name.replace(' (clock,memo,cal)', '').replace(' &amp; ', '_').replace(' ', '_') for
                            team_name in list_team_tg[tg]]
    return list_team_tg


def convert_data_format(table_weekly):
    arr_data = []
    for project in table_weekly:
        arr_temp = []
        latest_item = len(project) - 1
        for item in project:
            try:
                if item == 0:
                    arr_temp.append('-')
                else:
                    arr_temp.append(int(item))
            except:
                arr_temp.append(item)

        if arr_temp[0].find('MR') == -1:
            if type(arr_temp[latest_item]) is str or arr_temp[latest_item] >= 5:
                arr_data.append(arr_temp)

    return arr_data


def submit_weekly_rp(num_week, chart_last_week, table_last_week, table_tg, chart_pie_tg, line_chart):
    """ create layout weekly report update Wiki """
    pageTitle = "BA - Weekly Report"
    title_chart_weekly = 'Summary issue week ' + str(num_week)

    htmltable, empty_table = Wiki.create_table(table_last_week)
    htmlTableTG, empty_table = Wiki.create_table(table_tg)
    htmlChartLastWeek = Wiki.createHtmlChart(chart_last_week, title_chart_weekly, 'bar', orientation='horizontal',
                                             rangeAxis=5)

    htmlChartWeekly = Wiki.createHtmlChart(line_chart, 'Weekly total issue', 'line', orientation='vertical',
                                           rangeAxis=5)
    htmlChartTG = Wiki.createHtmlChart(chart_pie_tg, 'Total urgent issue', 'pie')

    html = Wiki.createLayoutWeekly(htmlChartLastWeek, htmlChartWeekly, htmlTableTG, htmlChartTG, htmltable)
    Wiki.updateWiki(pageTitle, html)


def submit_monthly_rp(num_month, last_month_team, tg_table, tg_chart_pie, line_chart):
    """ create layout weekly report update Wiki """
    pageTitle = "BA - Monthly report"
    title_chart_weekly = 'Summary month ' + str(num_month)
    htmlChartLastMonth = Wiki.createHtmlChart(last_month_team, title_chart_weekly, 'bar', orientation='horizontal',
                                              rangeAxis=10)
    htmlChartLine = Wiki.createHtmlChart(line_chart, 'Monthly total issue', 'line', orientation='vertical',
                                           rangeAxis=50)
    htmlTable, empty_table = Wiki.create_table(tg_table)
    htmlChartPie = Wiki.createHtmlChart(tg_chart_pie, 'Total issue', 'pie')

    html = Wiki.createLayoutMonthly(htmlChartLastMonth, htmlChartLine, htmlTable, htmlChartPie)
    Wiki.updateWiki(pageTitle, html)

"""
###### to test post weekly report  ######
if __name__ == "__main__":
    today = datetime.strptime('2018-01-01', '%Y-%m-%d').date()
    analysis_week_month(today, "Monthly")
"""