select t.* from (
select a.PARSING_SCHEMA_NAME username,
       sql_id,
       sql_text,
       '%s',
       round(a.BUFFER_GETS /executions,2) ,
       round(elapsed_time/1000000/executions,2) elap_per_exec,
       optimizer_cost
  from v$sql a
 where a.PARSING_SCHEMA_NAME not in
       ('SYS', 'MZ3_MON', 'ORACLE_OCM', 'DBSNMP')
   and sql_text not like '/*%%'
   and sql_text not like '%%/*+NESTED_TABLE_GET_REFS+*/%%'
   and sql_text not like '%%/* OPT_DYN_SAMP */%%'
   and sql_text not like '%%/* SQL Analyze%%'
   and sql_text not like '%%/* ds_svc */%%'
   --and upper(module) not like '%%PL/SQL%%'
   --and upper(module) not like '%%PLSQL%%'
   and upper(ltrim(sql_text)) not like 'CALL%%'
   and upper(ltrim(sql_text)) not like 'EXECUTE%%'
   and upper(ltrim(sql_text)) not like 'DECLARE%%'
   and upper(sql_text) not like '%%DBMS_STATS%%'
   and upper(module) not like '%%SQL/PLUS%%'
   and a.LAST_ACTIVE_TIME between sysdate - '%s' and sysdate
   --and round(a.ELAPSED_TIME / 100000 / executions, 2) > 10
   and a.EXECUTIONS >0
   and a.sql_id not in (select sql_id from t_ingore_sql@to_5754 where ip ='%s')
 order by round(a.ELAPSED_TIME / 100000 / executions, 2) desc) t
 where rownum <=3
