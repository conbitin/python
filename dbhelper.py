import sqlite3 as sql

class SqlLiteHelper:
    def __init__(self):
        self.conn = sql.connect("mydatabase.db")
    
    def create(self):
        c = self.conn.cursor()
        c.execute("CREATE TABLE IF NOT EXISTS Employees(compTelNo text, deptNm text, emailAddr text, empNo text, enChagBizCn text, "
                  "enCompNm text, enDeptNm text, enFnm text, enJobgrdNm text, enJobplAddr text, enJobpoNm text, "
                  "mphoNo text, userId text primary key)")
        self.conn.commit()
        c.close()

    def insert(self, employee_list):
        c = self.conn.cursor()
        for employee in employee_list:
            print(employee[2])
            c.execute("INSERT OR IGNORE INTO Employees VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", employee)
        self.conn.commit()
        c.close()

    def query(self, em_id):
        c = self.conn.cursor()
        c.execute("SELECT * FROM Employees WHERE userId=?", (em_id,))
        rows = c.fetchall()
        self.conn.commit()
        c.close()
        return rows