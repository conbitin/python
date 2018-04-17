import sqlite3 as sql

class SqlLiteHelper:
    def __init__(self):
        self.conn = sql.connect("mydatabase.db")
    
    def create(self):
        c = self.conn.cursor()
        c.execute("CREATE TABLE IF NOT EXISTS Employees(compTelNo text, deptNm text, emailAddr text, empNo text, enChagBizCn text, "
                  "enCompNm text, enDeptNm text, enFnm text, enJobgrdNm text, enJobplAddr text, enJobpoNm text, "
                  "mphoNo text, userId text primary key, execYn text, rlnmYn text, dispJobpoNm text, dispEnJobgrdNm text, "
                  "dispJobgrdNm text, dispJobgrdJobpoIndiCd text)")
        self.conn.commit()
        c.close()

    def insert(self, employee_list):
        c = self.conn.cursor()
        for employee in employee_list:
            print(employee[2])
            c.execute("INSERT OR IGNORE INTO Employees VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", employee)
        self.conn.commit()
        c.close()

    def query(self, em_id):
        c = self.conn.cursor()
        c.execute("SELECT * FROM Employees WHERE userId=?", (em_id,))
        rows = c.fetchall()
        self.conn.commit()
        c.close()
        return rows

    def make_db_for_idm(self):
        self.conn.create_function("isDigit", 1, self.is_digit)
        c = self.conn.cursor()
        c.execute("SELECT userId, empNo, rlnmYn FROM Employees WHERE rlnmYn!='V' AND isDigit(empNo)")
        rows = c.fetchall()
        self.conn.commit()
        c.close()

        self.conn = sql.connect("myidmdatabase.db")
        c = self.conn.cursor()
        c.execute("DROP TABLE IF EXISTS Employees")
        c.execute("CREATE TABLE Employees(userId text primary key, empNo text, rlnmYn text)")
        for row in rows:
            c.execute("INSERT OR IGNORE INTO Employees VALUES(?, ?, ?)", row)
        self.conn.commit()
        c.close()

    def is_digit(self, s):
        if s:
            return s.isdigit()
        return False