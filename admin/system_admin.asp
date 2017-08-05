<!--#include file="system/core/initialize_system.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=c_title%> | <%=Sysname%>v<%=versions%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
</head>
<frameset rows="80,*" cols="*" frameborder="no" border="0" framespacing="0" >
<frame src="system/models_system/top.asp" name="topFrame" id="topFrame" scrolling="no" noresize>
<frameset rows="15,*,15" cols="*" frameborder="no" border="0" framespacing="0" >
<frame src="system/models_system/top.asp?page=top" name="topFrame" id="topFrame" scrolling="no" noresize>
<frameset cols="172,*" name="bodyFrame" id="bodyFrame" frameborder="no" border="0" framespacing="0"  >
<frame src="system/models_system/meun.asp" name="menu" id="menu" scrolling="auto" noresize>
<frame src="system/plugins/admin_aspcheck.asp" name="main" id="main" scrolling="auto" noresize></frameset>
<frame src="system/models_system/top.asp?page=bottom" name="topFrame" id="topFrame" scrolling="no" noresize></frameset>
</frameset>
<noframes><body>你的浏览器不支持框架！</body></noframes>
</html>