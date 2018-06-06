下一步计划：


更新历史：

2017-12-05
1，将忽略原因加载到数据库

2017-12-04
1，新增按照owner 过滤sql 的功能。

2017-11-29
1，修复调用py2excel 和ex2oracle 时访问路径错乱的bug
2，修复当上一个库没有 top sql，导致下一个库的top sql无法写入 excel 的bug


2017-12-27
1，实现抓取clob 字段 sql_fulltext
2，cx_Oracle 5.下的版本有可能会出现ProgrammingError: LOB variable no longer valid after subsequent fetch 的报错，升级到 6以上即可。
3，windows 安装 whl文件 python -m pip install xxx.whl


2017-12-19
1，加载忽略sql 到数据库时，新增2个字段，group name 和 audit time
2017-12-07
1，迁移项目至linux平台
2，解决拼接sql时，由于在linux平台下回车符异常导致的sql 换行问题

2017-12-04
1，增加单元格冻结功能，方便处理excel
2，增加根据应用用户名，抓取top sql的功能

2017-11-29
1，修复调用py2excel 和ex2oracle 时访问路径错乱的bug
2，通过返回 m，修复当上一个库没有top sql，导致下一个库的top sql 无法写入excel 的bug，
3，调整excel 生成算法。无top sql时，不产生excel

2017-11-28
1,  file_name = file_name.encode('GB2312')
    reload(sys)
    sys.setdefaultencoding('utf-8'
    通过以上代码实现中文文件名不乱码。
2，解决循环超出后，导致程序报错：valueError more than 4094 XFs(styles)
3，加载被忽略的sql到数据库，用于下次执行程序时可以排除掉这部分sql

2017-11-27
-- 11-27 09：00
1，通过设置os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'， 解决top sql text 中存在乱码导致可能出现的程序异常
2，通过保存忽略sql，在下次查询top sql时用dblink 排查已经被忽略的sql
2017-11-24
1，实现从目标库抓取top sql 并写入文件头下方
2017-11-23
1，实现写入excel 文件头
