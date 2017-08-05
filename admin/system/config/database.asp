<%
'###### 连接后台数据库###### 
'on error resume next
set conn = dataConn("",dbpath)
set sysconn = dataConn("",sysdbpath)
	
'###### 读取配置信息 ###### 
set config=server.createobject("adodb.recordset") 
    exec="select * from sys_config" 
    config.open exec,sysconn,1,1
	if not config.eof then
	   c_title      =config("title")
	   c_url        =config("url")
	   c_contact    =config("contact")
	   c_tel        =config("tel")
	   c_fax        =config("fax")
	   c_mobile     =config("mobile")
	   c_email      =config("email")
	   c_qq         =config("qq")
	   c_msn        =config("msn")
	   c_address    =config("address")
	   c_keywords   =config("keywords")
	   c_description=config("description")
	   c_copyright  =config("copyright")
	end if
	config.close
set config=nothing
%>