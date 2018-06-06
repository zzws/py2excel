# -*- coding: utf-8 -*-

import main_functions as mf
import xlwt
import sys,os
import datetime
# import ex2oracle as ex2


reload(sys)
sys.setdefaultencoding('utf-8')
os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'

input_days =14
date1 = datetime.datetime.now().strftime("%Y-%m-%d")


# 获取父文件夹路径
father_file_path = mf.get_father_file_path()

# 获取需要执行的 sql 脚本
# sql_list = mf.get_sql_scripts(father_file_path,'py2excel')
sql_list =['time_consume']
# 打开db_info文件
db_infos = mf.open_files(father_file_path, 'conf', 'db_info.txt')

file = xlwt.Workbook(encoding = 'utf-8',style_compression=2)
""" style_compression=2 不加这个会导致,程序出现报错：xlwt set style making error: More than 4094 XFs (styles)。
 或者可以将 style 放到循环的外层。"""


table = file.add_sheet('SQL 审核', cell_overwrite_ok = True)

'''从第二行，第八列冻结单元格'''
table.panes_frozen = True
table.horz_split_pos= 2
table.vert_split_pos= 8

# 定义单元格长度
mf.col_length(table)

# 创建 excel 头
mf.create_excel_head(file,table)

m=2  #起始填充数据标志位

#按照组名和时间 构建excel 名
group_name = db_infos[0].split(',')[3]
file_name = 'SQL审核12_' + group_name + '_' + date1 + ".xls"
file_name = father_file_path+'/output/'+file_name.encode('GB2312')
# print file_name


for line in db_infos:
    # 遍历所有目标数据库，开始执行sql
    tns = line.split(',')[-1]
    ip=tns.split("/")[1].split("@")[1].split(':')[0]

    line_list = line.split(',')
    new_line = line_list.pop()

    # print sql_list

    for check in sql_list:
        #填充时间到行头
        line_list.insert(0, date1)
        # print ("len(line_list)= " + str(len(line_list)))

        # 获取需要执行的完整sql
        sql_file = check + ".sql"
        sql = mf.open_files(father_file_path, 'scripts', sql_file,check, ip, input_days)
        # print sql

        #检查是否已经存在 to_5754 的dblink
        mf.dblink("drop",tns)

        # 检查是否有 db link 的权限
        is_check = mf.dblink("check",tns)

        if is_check=="NO_DB_LINK_PRIV":
            continue #跳过没有dblink 权限的目标库，打印日志，并继续下一个库。
        else:
            # 连接到target db，并执行sql
            # print sql
            target_conn = mf.conn_target_db('py2excel',sql, tns, ip)

            if target_conn == 'CONNECT_FAILED':
                print (ip + " can cannot be connected.")
                continue;
            else:
                '''为了迭代 m ，在write_line 函数中返回了 m '''
                m = mf.write_line(target_conn,line_list,ip,table,m)

        mf.dblink("drop",tns)

file.save(file_name)

mf.del_null_excel(father_file_path)