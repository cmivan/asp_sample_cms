<!--#include file="email_config.asp"-->

<% 
set rs=server.createobject("adodb.recordset") 
sql="select * from Email_default order by id desc" 
rs.open sql,conn,1,1   
%>
<br> 
<%
'添加功能和修改功能

if request("action")<>"frommail" and request("action")<>"mailsubject" and request("action")<>"mailbody" then 
   'response.end
else

if request("action")="frommail" then
frommail=request("frommail")
conn.execute "update Email_default set frommail='"&frommail&"'"
response.write"<SCRIPT language=JavaScript>alert('发信人地址更改成功！');"
response.write "</SCRIPT>"
end if

if request("action")="mailsubject" then
mailsubject=request("mailsubject")
conn.execute "update Email_default set mailsubject='"&mailsubject&"'"
response.write"<SCRIPT language=JavaScript>alert('邮件默认标题更改成功！');"
response.write "</SCRIPT>"
end if

if request("action")="mailbody" then
mailbody=request("mailbody")
conn.execute "update Email_default set mailbody='"&mailbody&"'"
response.write"<SCRIPT language=JavaScript>alert('邮件默认内容更改成功！');"
response.write "</SCRIPT>"
end if

end if
%>
<table border="0" align=center cellpadding=0 cellspacing=10 class="forum1">
<tr><td>
<!--#include file="email_top.asp"-->

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="forum2">
<tr class="forumRaw">
<td align="center" valign="top" class="forumRaw"><b>邮 件</b> <strong>列 表 设 置</strong></td>
</tr>

<tr class="forumRow">
<td align="left" valign="top">
<table width="60%" border="0" cellpadding="0" cellspacing="2">
<form name="addcat" method="post" action="EMAIL_Mail_default.asp">
<tr>
<td align="center"><input name="frommail" style="width:100%;"  type="text" value="<%=rs("frommail")%>" size="25" maxlength="50" /></td>
<td width="140" align="left"><input style="width:100%;"  type="submit" name="Submit" value="保存默认发信人" class="Tips_bo" />
<font color="#000000"><b>
<input name="action" type="hidden" id="action" value="frommail" />
</b></font></td>
</tr>
</form>
</table></td>
</tr>

<tr class="forumRow">
<td align="left" valign="top"><table width="60%" border="0" cellpadding="0" cellspacing="2">
<form name="form1" method="post" action="EMAIL_Mail_default.asp">
<tr><td align="center"><input name="mailsubject" style="width:100%;"  type="text" value="<%=rs("mailsubject")%>" size="25" maxlength="50" /></td>
<td width="140" align="left"><input style="width:100%;"  type="submit" name="Submit22" value="保存默认标题" class="Tips_bo" />
<input name="action" type="hidden" id="action" value="mailsubject" /></td>
</tr>
</form>
</table>
</td></tr>

<tr class="forumRaw">
<td valign="top">
<form name="form1" method="post" action="EMAIL_Mail_default.asp">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="forum2">
<tr class="forumRaw"><td><p><font color="#000000">设置邮件默认发送内容</font>
<input name="action" type="hidden" id="action" value="mailbody" />
</p></td><td width="60">&nbsp;</td>
</tr>

<tr>
<td align="center"><textarea name="mailbody" style="width:100%;height:65px;" cols="46" rows="5"><%=rs("mailbody")%></textarea></td>
<td width="60" align="center"><input type="submit" style="height:63px;" name="Submit2" value="提交保存" class="Tips_bo" /></td>
</tr></table>
</form>
</td></tr>
</table>
</td></tr>
</table>
