# -*- coding: utf8 -*-
import psycopg2
import xlrd
import sys
sys.getdefaultencoding()

def read_excel():


    fname = "jingji.xlsx"
    bk = xlrd.open_workbook(fname)
    shxrange = range(bk.nsheets)
    try:
        sh = bk.sheet_by_name("x1433311643")
    except:
        print "no sheet in %s named Sheet1" % fname
    # 获取行数
    nrows = sh.nrows
    # 获取列数
    ncols = sh.ncols
    print "nrows %d, ncols %d" % (nrows, ncols)
    # 获取第一行第一列数据
    cell_value = sh.cell_value(1, 1)

    row_list = []
    col_list = []

    a={'as':'必修课'}

    col2_list=[]
    for i in xrange(5,nrows-4):
        for j in xrange(0,ncols-1,5):
            cell_va=sh.cell_value(i, j)
            if cell_va:
                cell_bixiu = sh.cell_value(i, j+4)
                if cell_bixiu.encode('utf-8')==a.values()[0]:
                    col2_list.append(cell_va)


    # 获取各行数据
    for i in range(1, ncols):
        cell_values = sh.col_values(i)

        col_list.append(cell_values)


    return col2_list

def creat_table():
    conn = psycopg2.connect(database="student", user="hanqing", password="hanqing", host="127.0.0.1", port="5432")

    print "Opened database successfully"

    cur = conn.cursor()
    cur.execute('''CREATE TABLE student_jingji
           (ID INT PRIMARY KEY     NOT NULL,
           NAME           TEXT    NOT NULL );''')
    print "Table created successfully"

    conn.commit()
    conn.close()


def insert_db(i,value):
    conn = psycopg2.connect(database="student", user="hanqing", password="hanqing", host="127.0.0.1", port="5432")

    print "Opened database successfully"

    cur = conn.cursor()

    cur.execute("""INSERT INTO student_jingji (ID,NAME) \
          VALUES (%s, %s)""",(i,value));

    conn.commit()
    conn.close()

    print "insert student success %s",i


if __name__=='__main__':
    conn = psycopg2.connect(database="student", user="hanqing", password="hanqing", host="127.0.0.1", port="5432")

    print "Opened database successfully"

    # excel_value=read_excel()
    # for i in xrange(len((excel_value[7]))):
    #     print  excel_value[7][i]
    creat_table()
    excel_value = read_excel()
    for i in xrange(len(excel_value)):
        print  excel_value[i]

        # insert_db(i,excel_value[7][i])
        cur = conn.cursor()

        cur.execute("""INSERT INTO student_jingji (ID,NAME) \
                  VALUES (%s,%s)""", (i+1,excel_value[i]));

    conn.commit()
    conn.close()

    # conn.close()



    print excel_value


