# -*- coding: utf8 -*-
import psycopg2
import xlrd
import os




def read_excel(fname):

    # fname = "/home/hanqing/django_workspace/1.xlsx"
    bk = xlrd.open_workbook(fname)

    print "文件名为;",fname

    # global jiankong
    # jiankong=1
    # if jiankong is not 0:

        # shxrange = range(bk.nsheets)
        # global number
        # number = len(shxrange) - 1
        # jiankong=0

        # sh = bk.sheet_by_index(number)


    # else:
    try:
        shxrange = range(bk.nsheets)

        # sh = bk.sheet_by_name("x1203330854")
        number=len(shxrange)-1
        sh = bk.sheet_by_index(number)

        number = number - 1

        sh = bk.sheet_by_index(number)


        print "sheet in %s " % sh.name


    except:
        print "no sheet in %s named Sheet1" % fname

    # 获取行数
    nrows = sh.nrows
    # 获取列数
    ncols = sh.ncols
    print "nrows %d, ncols %d" % (nrows, ncols)
    # 获取第一行第一列数据
    # cell_value = sh.cell_value(1, 1)

    row_list = []
    col_list = []

    a = {'as': '必修课'}

    col2_list = []
    for i in xrange(5, nrows - 4):
        for j in xrange(0, ncols - 1, 5):
            cell_va = sh.cell_value(i, j)
            if cell_va:
                cell_bixiu = sh.cell_value(i, j + 4)
                if cell_bixiu.encode('utf-8') == a.values()[0]:
                    col2_list.append(cell_va)

    # 获取各行数据
    for i in range(1, ncols):
        cell_values = sh.col_values(i)

        col_list.append(cell_values)


    return col2_list


if __name__=='__main__':
    conn = psycopg2.connect(database="student", user="hanqing", password="hanqing", host="127.0.0.1", port="5432")

    print "Opened database successfully"

    # global number

    filepathc="/home/hanqing/django_workspace/14wenhua"
    a={}
    pathdir = os.listdir(filepathc)
    sorted_pre=[]

    for line in pathdir:
        if line[0] is not '.':
            if line[10]=='.':
                sorted_pre.append(int(line[9]))
                a[int(line[9])]=line

            else:
                sorted_pre.append(int(line[9:11]))
            # pathdir.index(line)
                a[(line[9:11])]=line
    # b=sorted(a.keys)
    a=sorted(a.iteritems(),key=lambda d:d[0])

    global jiankong
    jiankong=1

    for line in a:
        alldir=line[1]

    # for alldir in pathdir:
        childdir = os.path.join('%s/%s' % (filepathc, alldir))

        excel_value=read_excel(childdir,jiankong)

        excel_vals=[]
        for i in xrange(len(excel_value)):
            excel_vals.append(excel_value[i].encode('utf-8'))


        cur = conn.cursor()

        cur.execute("""select name from  student_wenhua""");

        pre_read=[]

        values=cur.fetchall()
        for i in xrange(len(values)):

            pre_read.append(values[i][0])

        # print excel_value
        # print pre_read
        ret=[]
        if pre_read == excel_vals:
            print "the same excel"
        else:
            for i in pre_read:
                if i not in excel_vals:
                    ret.append(i)

        print "缺少的课程数: ",(len(ret))

        for line in ret:

            print line

    conn.close()


