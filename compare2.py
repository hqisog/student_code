# -*- coding: utf8 -*-
import psycopg2
import xlrd
import os




# def read_excel(fname):
#     bk = xlrd.open_workbook(fname)
#     print "文件名为;",fname
#     try:
#         shxrange = range(bk.nsheets)
#         # sh = bk.sheet_by_name("x1203330854")
#         number=len(shxrange)-1
#         sh = bk.sheet_by_index(number)
#         number = number - 1
#         sh = bk.sheet_by_index(number)
#         print "sheet in %s " % sh.name
#     except:
#         print "no sheet in %s named Sheet1" % fname
#     # 获取行数
#     nrows = sh.nrows
#     # 获取列数
#     ncols = sh.ncols
#     print "nrows %d, ncols %d" % (nrows, ncols)
#     # 获取第一行第一列数据
#     # cell_value = sh.cell_value(1, 1)
#     a = {'as': '必修课'}
#     col2_list = []
#     for i in xrange(5, nrows - 4):
#         for j in xrange(0, ncols - 1, 5):
#             cell_va = sh.cell_value(i, j)
#             if cell_va:
#                 cell_bixiu = sh.cell_value(i, j + 4)
#                 if cell_bixiu.encode('utf-8') == a.values()[0]:
#                     col2_list.append(cell_va)
#     return col2_list


if __name__=='__main__':
    conn = psycopg2.connect(database="student", user="hanqing", password="hanqing", host="127.0.0.1", port="5432")

    print "Opened database successfully"

#文件目录 需要手动修改
    filepathc="/home/hanqing/django_workspace/14jingji2"
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
                a[(line[9:11])]=line
    a=sorted(a.iteritems(),key=lambda d:d[0])

    n=0
    for line in a:
        alldir=line[1]
        n=n+1

    # for alldir in pathdir:
        childdir = os.path.join('%s/%s' % (filepathc, alldir))

        # excel_value=read_excel(childdir)

        #将excel中的数据读出来，按必修课读出用 +++分割
        # +++++++++


        bk = xlrd.open_workbook(childdir)
        print "文件名为;", childdir
        if  n==1 :
            # sh = bk.sheet_by_index(number)
            shxrange = range(bk.nsheets)
            number = len(shxrange) - 1
            sh = bk.sheet_by_index(number)

        else:



            try:
                # sh = bk.sheet_by_name("x1203330854")
                # sh = bk.sheet_by_index(number)\
                #注意这里number有时需要手动改变-1或 -2  todo
                #number的作用 是为 检索
                number = number - 1
                sh = bk.sheet_by_index(number)
                print "sheet in %s " % sh.name
            except:
                print "no sheet in %s named Sheet1" % childdir
        # 获取行数
        nrows = sh.nrows
        # 获取列数
        ncols = sh.ncols
        print "nrows %d, ncols %d" % (nrows, ncols)
        # 获取第一行第一列数据
        # cell_value = sh.cell_value(1, 1)

        #转换为字典 是为 解决编码问题
        a = {'as': '必修课'}
        excel_value = []
        for i in xrange(5, nrows - 4):
            for j in xrange(0, ncols - 1, 5):
                cell_va = sh.cell_value(i, j)
                if cell_va:
                    cell_bixiu = sh.cell_value(i, j + 4)
                    if cell_bixiu.encode('utf-8') == a.values()[0]:
                        excel_value.append(cell_va)

        # ++++++


        excel_vals=[]
        for i in xrange(len(excel_value)):
            excel_vals.append(excel_value[i].encode('utf-8'))


        cur = conn.cursor()

        cur.execute("""select name from  student_jingji""");

        pre_read=[]

        values=cur.fetchall()
        for i in xrange(len(values)):

            pre_read.append(values[i][0])

        # 对比 基础数据 和 读入数据的不同
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


