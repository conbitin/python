import ReadDataFromExcel as DataEXL


def create_data_urgent(type_data="URGENT"):
    """ create variable for chart urgent """
    input_data = DataEXL.urgent_count_all_prj()
    list_team = DataEXL.get_list_team()
    total_data = dict((team, 0) for team in list_team)

    for prj in input_data:
        for team in input_data[prj]:
            total_data[team] += input_data[prj][team]
    out_total_data = "var all_models = " + str(convert_data(total_data)) + ";\n"

    prj_detail = {}
    for prj in input_data:
        prj_detail[prj] = convert_data(input_data[prj])
    out_detail_data = "var prj_urgent_detail = " + str(prj_detail) + '; \n'
    data_chart = out_total_data + out_detail_data

    if type_data == "URGENT":
        return data_chart
    else:
        return list_team, total_data


def create_data_plm_jira_issue(main_summary_data, num_of_jira_task_by_team):
    list_summary_plm = []

    for key, value in num_of_jira_task_by_team.items():
        num_of_task_jira = value[0] + value[1]
        num_of_issue_plm = main_summary_data[key]

        if num_of_task_jira + num_of_issue_plm > 0:
            plm_issue = [key, num_of_issue_plm, num_of_task_jira]
            list_summary_plm.append(plm_issue)

    data_summary_plm = "var summary_data_chart = " + str([['', 'PLM Issue', 'Jira Task']] + list_summary_plm) + "; \n"
    return data_summary_plm


def convert_data(dict_data):
    """ convert dict data to chart """
    out_data = [["", ""]]
    for team in dict_data:
        if dict_data[team] > 0:
            out_data.append([team, dict_data[team]])
    return out_data

# if __name__ == "__main__":
#     DataEXL.issue_analysis()
#     print(create_data_urgent())
