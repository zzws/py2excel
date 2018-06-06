# -*- coding: utf-8 -*-

import xlwt,xlrd
import cx_Oracle
import datetime
import os,sys

def properties(i):

    '''设置单元格背景色
　　0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue,
　　5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon,
　　17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow ,
　　20 = Dark Magenta, 21 = Teal,
　　22 = Light Gray, 23 = Dark Gray
　　'''

    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = i #黄色

    '''设置位置'''
    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER   #水平居中
    alignment.vert = xlwt.Alignment.VERT_CENTER    #垂直居中

    '''字体和大小'''
    font = xlwt.Font()
    font._weight =20
    # font.name ='SimSun'

    '''设置所有框线'''
    borders = xlwt.Borders()
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1

    style = xlwt.XFStyle() # Create Style

    style.alignment = alignment # Add Alignment to Style
    style.pattern = pattern
    style.font = font
    style.borders = borders

    return  style

def get_father_file_path():
    #获取主程序所在路径

    path=sys.path[0]
    if os.path.isdir(path):
        # current_pwd
        return os.path.abspath(os.path.join(os.path.dirname(__file__)))

        # father pwd
        # return os.path.abspath(os.path.join(os.path.dirname(__file__),os.pardir))

        # grandpar pwd
        # return os.path.abspath(os.path.join(os.path.dirname(__file__),os.pardir,os.pardir))
    elif os.path.isfile(path):
        return os.path.dirname(path)

def get_sql_scripts (father_file_path):
    for root, dirs, files in os.walk(father_file_path+'/output'):
        tmp_file_list = []
        for i in files:
            tmp_file_list.append(i.split('.')[0])

        return tmp_file_list

def col_length(table):
    for i in range (12):
        col_length = table.col(i)
        if i <> 8:
            col_length.width = 256 * 15
        else:
            col_length.width = 256 * 45


def create_excel_head(file,table):
    i = 0

    list  = [['日期', '数据库ip', '应用系统', '开发中心', '开发组', '接口人', '执行用户', 'SQL_ID', 'SQL文本', '问题描述', '逻辑读', '耗时', 'COSTS'],
            ['优化方式', '优化方案', '开发负责人', '处理人', '处理时间', '预计投产时间', '实际投产时间', '备注'],
            ['确认人', '确认时间', '确认结果', '首次通过', '确认意见', '备注'], ['优化前', '优化后', '提升(倍)'], ['优化前', '优化后', '提升(倍)'],
            ['优化前', '优化后', '提升(倍)'], ['更新时间']]
    list2 = [[0, 0, 0, 12, 'DBA'], [0, 0, 13, 20, '开发同事'], [0, 0, 21, 26, 'DBA组员'], [0, 0, 27, 29, '逻辑读'],
             [0, 0, 30, 32, '耗时'], [0, 0, 33, 35, 'COSTS']]

    for m in range(6):
        style = properties(m+2)
        table.write_merge(list2[m][0],list2[m][1],list2[m][2],list2[m][3],list2[m][4],style)  #合并单元格，构造第一行
        for j in list[m]:

            if i <= list2[m][3]:
                table.write(1, i, j, style)
                i=i+1
            else:
                break

        table.write_merge(list2[m][0],list2[m][1],list2[m][2],list2[m][3],list2[m][4],style) # 构造第二行

def dblink(type, tns=''):
    if type =="drop":
        try:
            target_conn = cx_Oracle.connect('%s'%tns)
            conn = target_conn.cursor()
            sql_result = conn.execute("select count(1) from user_db_links where db_link='TO_5754'")
            is_db_link = sql_result.fetchall()[0][0]
            if is_db_link > 0:
                conn.execute("drop database link to_5754")
            else:
                pass
        except Exception,dberror:
            print dberror
            return "DROP_LINK_FAILED"
    elif type =="check":
        try:
            target_conn = cx_Oracle.connect('%s' % tns)
            conn = target_conn.cursor()
            sql_result = conn.execute("select count(1) from user_role_privs where granted_role='DBA'")
            is_dba = sql_result.fetchall()[0][0]

            sql_result = conn.execute("select count(1) from user_sys_privs where privilege = 'CREATE DATABASE LINK'")
            is_db_link = sql_result.fetchall()[0][0]

            if is_dba == 0 and is_db_link == 0:
                print (tns + "需要创建database link 的权限。").encode('gb2312')
                return "NO_DB_LINK_PRIV"
        except Exception,dberror:
            print dberror
            return "CHECK_LINK_FAILED"



def conn_target_db(type,sql='',tns='',ip=''):
    # 连接到目标数据库，抓取数据

    try:
        target_conn=cx_Oracle.connect('%s'%tns)
        conn = target_conn.cursor()
        # print type
        if type == 'py2excel':
            conn.execute("create database link to_5754 connect to zc identified by Enmotech using '192.168.56.102/sam'")
            return conn.execute(sql)
        else:
            conn.execute(sql)
            conn.execute('commit')


    except Exception,dberror:
        print dberror
        # print '%s,connect fiailed.'%ip
        return "CONNECT_FAILED" # this return used for the next if statement. if has not this return , this program will return errors

def get_sql_scripts (father_file_path,type=''):

    '''py2excel 用于读取sql 脚本目录，
       ex2oracle 用于读取excel 目录'''
    if type == "py2excel":
        tmp_path = '/scripts'
    elif type =="db_info":
        tmp_path = '/conf'
    else:
        tmp_path ='/output'

    for root, dirs, files in os.walk(father_file_path + tmp_path):
        tmp_file_list = []
        for i in files:
            tmp_file_list.append(i.split('.')[0])

        return tmp_file_list

def open_files(father_file_path,middle_path,file_name,check='',ip='',owner='',input_days=3):
    #打开文件

    file = open(father_file_path+'\\'+middle_path+'\\'+file_name)
    if middle_path == 'conf':
        lines = file.readlines()
        return lines
    else:
        #拼接  sql
        tmp_sql = file.read()
        full_sql = tmp_sql%(check,owner,input_days,ip)
        return full_sql
    file.close()

def write_line(target_conn,line_list,ip,table,m,check):
    '''写查询结果到excel'''
    sql_result = target_conn.fetchall()
    # m = m + len(sql_result)
    if len(sql_result) > 0:

        for each_result in sql_result:
            # print type(each_result[2])
            new_list = []
            for i in range(len(each_result)):
                # print type(each_result[i])
                if i == 2:
                    '''第3个元素是 sql_fulltext ,clob 类型的'''
                    # print i, type(each_result[i]),each_result[2].read(),type(each_result[2].read())
                    new_list.append( each_result[2].read())
                else:
                    new_list.append(each_result[i])

            new_sql_result = line_list + new_list

            for i in range(len(new_sql_result)):
                style = properties(22)
                table.write(m, i, new_sql_result[i], style)
                style = properties(22)
                table.write(m, i, new_sql_result[i], style)
                style = properties(5)
                table.write(m,27,new_sql_result[i-2],style) # 优化前逻辑读
                style = properties(6)
                table.write(m, 30, new_sql_result[i - 1], style)#优化前耗时
                style = properties(7)
                table.write(m, 33, new_sql_result[i], style) #优化前cost
            m = m + 1
        print (ip+" done !!")
        return m
    else:
        print (ip + " " + check + " has no results !!")
        return m

def del_move_excel(father_file_path,type=''):
    # 删掉没有top sql的excel。 获取需要遍历的excel
    sql_list = get_sql_scripts(father_file_path, 'ex2oracle')
    for excel in sql_list:
        excel = unicode(excel, 'gb2312')
        # excel = unicode(excel, 'utf8') #linux 运行需要
        # print excel
        full_excel = father_file_path + '\output\\' + excel + '.xls'
        # full_excel = father_file_path + '/output/' + excel + '.xls' # for linux
        # data = xlrd.open_workbook(unicode(full_excel,'gb2312'))
        data = xlrd.open_workbook(full_excel)
        table = data.sheet_by_index(0)
        nrows = table.nrows #获取行数，用于删除没有结果集的excel

        if nrows <=2:
            os.remove(full_excel)
        else:
            pass

        


