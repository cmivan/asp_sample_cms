<!--#include file="email_config.asp"-->
<br />

<table border="0" align=center cellpadding=0 cellspacing=10 class="forum1">
<tr><td>
<!--#include file="email_top.asp"-->
<form name="adduser" method="post" action="EMAIL_AddUser.asp">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="forum2">
<tr class="forumRaw">
<td width="100" align="center" class="forumRow"><p><font color="#000000">E-mail地址</font></p>                  </td>
<td valign="top" class="forumRow"> <input name="email" type="text" class="bk" style="width:100%" maxlength="50"></td>
<td width="100" align="center" valign="top" class="forumRow"><input type="submit" name="Submit" value="提交保存" class="Tips_bo" />
<input type="hidden" name="action" value="adduser" /></td>
</tr>
</table> 
</form>
</td></tr>
</table>

<%
if request("action")<>"adduser" then
response.end
else
email=request("email")
mailnow=now()
conn.execute "insert into email (dateandtime,email) values ('"&mailnow&"','"&email&"')"
response.write"<SCRIPT language=JavaScript>alert('帐号添加成功，你可以继续添加！');"
response.write "</SCRIPT>"
end if
%>