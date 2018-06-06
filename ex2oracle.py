# -*- coding: utf-8 -*-

import xlwt
import sys,os
import datetime
import xlrd
import main_functions as mf

def ex2oracle():
    reload(sys)
    sys.setdefaultencoding('utf-8')
    os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'

    # 获取父文件夹路径
    father_file_path = mf.get_father_file_path()

    #获取需要遍历的excel
    sql_list = mf.get_sql_scripts(father_file_path,'ex2oracle')
    # print father_file_path+'\output'
    # print sql_list

    tns = "zc/Enmotech@192.168.56.102/sam"

    if len(sql_list) > 0 :
        #文件夹不为空，再继续执行

        '''由于每次开发反馈的可能有重复sql，这里和下面的sql3, 4,5 做重复忽略sql 的排除'''
        sql = "truncate table zc.t_ingore_sql "
        mf.conn_target_db('ex2oracle', sql, tns)

        for excel in sql_list:
            excel = unicode(excel,'gb2312')
            print excel
            full_excel = father_file_path+'\output\\'+excel+'.xls'
            # data = xlrd.open_workbook(unicode(full_excel,'gb2312'))
            data = xlrd.open_workbook(full_excel)
            table = data.sheet_by_index(0)

            for m in range(2,table.nrows):
                is_comm = table.cell(m,13).value.encode('utf8') #为了处理单元格内的中文

                if is_comm == "忽略":
                    # 将被忽略的sql id 放到数据库，用于下次执行时，过滤掉这部分sql
                    print m,table.cell(m,0),table.cell(m,1),table.cell(m,7),is_comm
                    datatime = table.cell(m,0).value
                    ip = table.cell(m,1).value
                    sql_id = table.cell(m,7).value
                    comments = table.cell(m,20).value
                    # sql = "insert into zc.t_ingore_sql_tmp values( '"+ datatime  +"','"+ sql_id +"','"+ ip +"')"
                    sql = "insert into zc.t_ingore_sql_tmp values( '"+ datatime  +"','"+ sql_id +"','"+ ip +"','"+ comments +"')"
                    print sql

                    mf.conn_target_db('ex2oracle',sql, tns)
                else :
                    pass


        sql3 = "insert into zc.t_ingore_sql (sql_id,ip,comments) select distinct sql_id,ip,comments from zc.t_ingore_sql_tmp "
        sql4 = "truncate table zc.t_ingore_sql_tmp"
        sql5 = "insert into zc.t_ingore_sql_tmp (sql_id,ip,comments) select  sql_id,ip,comments from zc.t_ingore_sql "

        for i in range(3,6):
            sql = 'sql' + str(i)
            sql = locals()[sql]

            mf.conn_target_db('ex2oracle', sql, tns)


            # mf.del_move_excel(father_file_path,'move')





ex2oracle()