# -*- coding: utf8 -*-
import xlrd
import redis
import json
import sys
sys.getdefaultencoding()

pool = redis.ConnectionPool(host='127.0.0.1', port=6379)


def save_redis(value):
    r = redis.Redis(connection_pool=pool)
    key='student_key'
    r.set(key,json.dumps(value),ex=7*24*60*60)
    # r.set(key,value,ex=7*24*60*60)


def get_redis():
    r = redis.Redis(connection_pool=pool)
    redis_value=r.mget('student_key')
    return redis_value



def read_excel():


    fname = "reflect.xls"
    bk = xlrd.open_workbook(fname)
    shxrange = range(bk.nsheets)
    try:
        sh = bk.sheet_by_name("guide.terminal(1)")
    except:
        print "no sheet in %s named Sheet1" % fname
    # 获取行数
    nrows = sh.nrows
    # 获取列数
    ncols = sh.ncols
    print "nrows %d, ncols %d" % (nrows, ncols)
    # 获取第一行第一列数据
    cell_value = sh.cell_value(1, 1)
    # print cell_value

    row_list = []
    col_list = []
    # 获取各行数据
    for i in range(1, ncols):

        cell_values = sh.col_values(i)

    # row_data = sh.row_values(i)
        col_list.append(cell_values)

    for i in xrange(len(col_list[7])-1):
        save_redis(col_list[7][i])




    return col_list



if __name__=='__main__':
    read_excel()
    # a={}
    excel_value=get_redis()
    print excel_value
    # for i in xrange(len(excel_value)-1):
    # b=excel_value[0].decode("utf-8")
    # print b
    #     d={str(i):excel_value[i]}
    #     a.append(d)
    # a=excel_value[0].decode("utf-8")
    # c=json.dumps(a, encoding='UTF-8', ensure_ascii=False)

    # print  a
    # r.set('name', 'zhangsan',ex=7*24*60*60)
    # print r.get('name')
