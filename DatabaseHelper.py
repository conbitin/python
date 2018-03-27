import sqlite3


class DatabaseHelper:
    def __init__(self, list_team):
        self.conn = sqlite3.connect('SmartManagement.db')
        self.list_team = list_team
        self.create_tb_lw_report()
        self.create_tb_weekly()
        self.create_tb_monthly()

    def create_tb_lw_report(self):
        """ create database to save data weekly """
        cursor = self.conn.cursor()
        create_command = "CREATE TABLE IF NOT EXISTS LastWeekReport(`index` TEXT, "
        for team in self.list_team:
            create_command += team + " INTEGER, "
        create_command += "Total INTEGER, Week INTEGER)"
        #print(create_command)
        cursor.execute(create_command)
        cursor.close()

    def create_tb_weekly(self):
        """ create table monthly report """
        cursor = self.conn.cursor()
        create_command = "CREATE TABLE IF NOT EXISTS WeeklyReport(Week INTEGER, "
        for team in self.list_team:
            create_command += team + " INTEGER, "
        create_command = create_command[:-2] + ")"
        # print(create_command)
        cursor.execute(create_command)
        cursor.close()

    def create_tb_monthly(self):
        """ create table monthly report """
        cursor = self.conn.cursor()
        create_command = "CREATE TABLE IF NOT EXISTS MonthlyReport(Month INTEGER, "
        for team in self.list_team:
            create_command += team + " INTEGER, "
        create_command = create_command[:-2] + ")"
        cursor.execute(create_command)
        cursor.close()

    def insert_dataFrame(self, data_frame):
        """ insert weekly to database """
        cursor = self.conn.cursor()
        data_frame.to_sql("LastWeekReport", self.conn, if_exists='append')
        self.conn.commit()
        cursor.close()

    def get_data_last_week(self, week_num):
        """ get data by week """
        cursor = self.conn.cursor()
        query_command = "SELECT `index`, "
        for team in self.list_team:
            query_command += team + ","
        query_command += " Total FROM LastWeekReport WHERE Week=?"

        cursor.execute(query_command, (week_num,))
        result = cursor.fetchall()
        cursor.close()
        return result

    def deleteByWeek(self, week_num, table_name):
        """ delete row by week to update new data """
        cursor = self.conn.cursor()
        delete_command = "DELETE FROM " + table_name + " WHERE Week=?"
        cursor.execute(delete_command, (week_num,))
        self.conn.commit()
        cursor.close()

    def delete_table_last_week(self):
        """ delete all row table LastWeekReport"""
        cursor = self.conn.cursor()
        delete_command = "DELETE FROM LastWeekReport"
        cursor.execute(delete_command)
        self.conn.commit()
        cursor.close()

    def checkWeekExists(self, week):
        """ check week has been update """
        cursor = self.conn.cursor()
        sql_command = "SELECT Week FROM WeeklyReport"
        cursor.execute(sql_command)
        result_set = cursor.fetchall()
        arr_result = [result[0] for result in result_set]
        return week in arr_result
        cursor.close()

    def insertDataWeekly(self, data_week):
        """ insert total issue weekly to monthly report """
        cursor = self.conn.cursor()
        value_command = " VALUES ("
        sql_command = "INSERT INTO WeeklyReport(Week,"
        for team in self.list_team:
            sql_command += team + ","
            value_command += " ?,"
        sql_command = sql_command[:-1] + ")"
        value_command += "? )"
        sql_command += value_command
        cursor.execute(sql_command, tuple(data_week))
        self.conn.commit()
        cursor.close()

    def updateDataWeekly(self, data_week):
        """ insert total issue weekly to monthly report """
        cursor = self.conn.cursor()
        sql_command = "UPDATE WeeklyReport SET Week=?,"
        for team in self.list_team:
            sql_command += team + "=?,"
        sql_command = sql_command[:-1] + "WHERE Week=" +str(data_week[0])
        cursor.execute(sql_command, tuple(data_week))
        self.conn.commit()
        cursor.close()

    def checkMonthExists(self, month):
        """ check week has been update """
        cursor = self.conn.cursor()
        sql_command = "SELECT Month FROM MonthlyReport"
        cursor.execute(sql_command)
        result_set = cursor.fetchall()
        arr_result = [result[0] for result in result_set]
        return month in arr_result
        cursor.close()

    def insertDataMonthly(self, data_month):
        """ insert total issue to monthly report """
        cursor = self.conn.cursor()
        value_command = " VALUES ("
        sql_command = "INSERT INTO MonthlyReport(Month,"
        for team in self.list_team:
            sql_command += team + ","
            value_command += " ?,"
        sql_command = sql_command[:-1] + ")"
        value_command += "? )"
        sql_command += value_command
        cursor.execute(sql_command, tuple(data_month))
        self.conn.commit()
        cursor.close()

    def updateDataMonthly(self, data_month):
        """ insert total issue to monthly report """
        cursor = self.conn.cursor()
        sql_command = "UPDATE MonthlyReport SET Month=?,"
        for team in self.list_team:
            sql_command += team + "=?,"
        sql_command = sql_command[:-1] + "WHERE Month=" +str(data_month[0])
        cursor.execute(sql_command, tuple(data_month))
        self.conn.commit()
        cursor.close()

    def getColumnName(self, table_name):
        """ get list column name of table """
        cursor = self.conn.cursor()
        command = "SELECT * FROM " +table_name
        cursor.execute(command)
        col_name_list = [tuple[0] for tuple in cursor.description]
        return col_name_list

    def getData(self, table_name):
        cursor = self.conn.cursor()
        query_command = "SELECT * FROM " +table_name
        cursor.execute(query_command)
        result = cursor.fetchall()
        cursor.close()
        return result