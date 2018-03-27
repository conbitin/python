import traceback
import WikiSubmit as Wiki
import ReadDataFromExcel as DataEXL
import pandas as pd

base_table = open('v2_template/table.html','r').read()
prj_has_issue = []


def getIssueCount(prj):
    return DataEXL.urgent_count_prj(prj.lower())


def insert_row(prj, dvr, pvr, pra, sra, issues):
    global base_table
    global prj_has_issue
    position = base_table.find('<!--insert_row-->')
    if issues != 0:
        prj_type = 'table-danger'
        prj_has_issue.append(prj.upper())
    else:
        prj_type = 'table-primary'

    row = '''<tr class="{}">
                            <td class="prj-name">{}</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td class="num-issue">{}</td>
                        </tr>'''.format(prj_type, prj, dvr, pvr, pra, sra, issues)
    base_table = base_table[:position] + row + base_table[position:]


def fill_urgent_project_detail():
    try:
        content = Wiki.api.getpagecontent(name=Wiki.pageUrgentPrjTitle, spacekey=Wiki.space)
        table = pd.read_html(content)
        table = table[0]
        for index, row in table[1:].iterrows():
            prj = row[0]

            dvr = row[1]
            if str(dvr).lower() == 'nan':
                dvr = '-'

            pvr = row[2]
            if str(pvr).lower() == 'nan':
                pvr = '-'

            pra = row[3]
            if str(pra).lower() == 'nan':
                pra = '-'

            sra = row[4]
            if str(sra).lower() == 'nan':
                sra = '-'

            insert_row(prj, dvr, pvr, pra, sra, getIssueCount(prj))

    except Exception as e:
        print(e)
        traceback.print_exc()
        print("Cannot get page %s content" % Wiki.pageUrgentPrjTitle)

    return base_table


def getUrgentTableCode():
    fill_urgent_project_detail()
    return base_table





