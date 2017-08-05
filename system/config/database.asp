<%
on error resume next
dim conn,connstr,connmdb

'****  Access数据库连接字符串
    connmdb = "system/data/KM_09_12_20.mdb" '数据库文件目录
    connstr = "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & server.mappath(connmdb)
'	connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & server.mappath(connmdb)
'****  SQL数据库连接字符串
'	connstr = "driver={SQL Server};server=192.168.0.1,7788;database=KM_09_12_20;UID=sa;PWD=sa"
'*****************************************

set conn = Server.CreateObject("ADODB.Connection") 
    conn.Open connstr
	if err then
	   err.Clear
	   set conn = nothing
	   response.Write "数据库连接有误!请检查相应文件!" & err.description
	   response.end
	end If
%>