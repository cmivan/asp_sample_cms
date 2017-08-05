<!--#include file="system/core/initialize.asp"-->
<!--#include file="system/helpers/md5_helper.asp"-->
<%
username = request.form("username")
password = md5(request.form("password"))
verifycode = request.form("verifycode") 
action = request.form("action")

if action="admin.login" then

	if username = "" then call alert.reFresh("用户名不能为空!","?")
	if password = "" then call alert.reFresh("密码不能为空!","?")
	if verifycode= "" then call alert.reFresh("验证码不能为空!","?") 
	if cstr(session("cmcms.code.num"))<>cstr(verifycode) then call alert.reFresh("验证码错误!","?")

	set rs = sysconn.execute("select top 1 * from sys_admin where username='"&username&"' and password='"&password&"'") 
		if rs.eof or rs.bof then
		   call alert.reFresh("帐号或者密码错误，请重新输入!","?")
		else
		   '#### 二重验证 ####-- 保证正常登陆
		   if password<>rs("password") then
			  call alert.reFresh("帐号或者密码错误，请重新输入!","?")
		   else
			  session("cmcms.admin") = username
			  session("user_id")     = rs("id")
			  session("adminname")   = rs("truename")
			  session("user_power")  = rs("power")
			  session("super")       = rs("super")
			  session("user_editbox")= rs("editbox")
			  response.redirect "system_admin.asp"
			  response.End()
		   end if
		end if
	set rs=nothing
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=c_title%> | <%=Sysname%>v<%=versions%></title>
<meta http-equiv="Content-Type" content="text/html;charset=utf-8" />
<script language="javascript" src="public/style/jquery.1.4.js"></script>
<script language="javascript" src="public/style/jquery.edit.style.js"></script>
<link rel="stylesheet" type="text/css" href="../public/assets/css/bootstrap.ie6.helper.css"/>
<link rel="stylesheet" type="text/css" href="../public/assets/css/bootstrap.css"/>
<style type="text/css">
.main{margin:100px auto 0px;background:#FFF;padding: 12px;padding-top:5px;width: 260px;}
.copyright{text-align:right;margin:10px auto;font-size:10px;color:#999999;}
.copyright a{font-weight:bold;color:#F63;text-decoration:none;}
.copyright a:hover{color:#000;}
.versions{font-weight:lighter;color:#000;font-family:Verdana, Geneva, sans-serif;}
.conutBox{display:none;}
</style>
</head><body>

<div class="main">

<form name="log_in" id="log_in" method="post" class="well form-inline" action="">
<label>帐号(Id)</label><br />
<input name="username" id="username" tabindex="1" type="text"/><br /><br />
<label>密码(Password)</label>
<input name="password" id="password" tabindex="2" type="password"/><br /><br />
<label>验证码(Code)</label><br />
<div style="position:relative; width:220px;">
<div style="position:absolute; right:4px; top:4px;">
<a href="javascript:void(0);" class="btn" id="verifycodeImg"><img src="public/valid/code.asp" /></a>
</div>
<input name="verifycode" id="verifycode" tabindex="3" type="text" maxlength="4"/>
</div>
<br /><br />

<button type="reset" class="btn">重置</button>&nbsp;
<button type="submit" name="submit" class="btn btn-success">
<span class="icon icon-user icon-white">&nbsp;</span> 我要登录(Login)
</button>
<input type="hidden" name="action" value="admin.login" />
</form>

<div class="copyright">
Powered by <a href="<%=Official%>" target="_blank"><%=Sysname%>&nbsp;<span class="versions">v<%=versions%></span></a>
<span class="conutBox">&nbsp;|&nbsp;<script src="http://s11.cnzz.com/stat.php?id=2058869&web_id=2058869" language="JavaScript"></script></span>
</div>

</div>


<script type="text/javascript">
var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));</script>
<script type="text/javascript">
try {var pageTracker = _gat._getTracker("UA-15415383-2");pageTracker._setDomainName("none");pageTracker._setAllowLinker(true);pageTracker._trackPageview();} catch(err) {}</script>

</body>
</html>