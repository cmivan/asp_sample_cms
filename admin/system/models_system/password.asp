<!--#include file="../../system/core/initialize_system.asp"-->
<!--#include file="../libraries/articlesHtml.asp"-->
<%
dim UserEdit
    UserEdit=request.QueryString("UserEdit")
'=========== 删除账号 ===============
dim delID
    delID=request.QueryString("delID")
 if delID<>"" and isnumeric(delID) then
    sysconn.execute("delete * from sys_admin where power<>1 and id="&delID)
 end if

'=========== 修改账号信息 ===============
if session("user_id")<>"" and request.Form("edit")<>"" then

    editbox  =request.form("editbox")
    username =request.form("username")
    username =replace(username,"'","")
    truename =request.form("truename")
    password =request.form("password")
    if editbox="" or isnumeric(editbox)=false then editbox=0

set rs=server.CreateObject("adodb.recordset")
    if request.Form("edit")="add" then
	   if session("user_power")=1 then
		  sql="select * from sys_admin where username='"&username&"'"
		  rs.open sql,sysconn,1,3
		  if not rs.eof then call alert.reFresh("该账号已存在!","?UserEdit=add")
		  rs.addnew
	   else
	      call alert.reFresh("权限不足!","?")
	   end if
    else
       sql="select * from sys_admin where id="&session("user_id")
       rs.open sql,sysconn,1,3
       if rs.eof then call alert.reFresh("账号信息不存在!","?UserEdit=add")
    end if

    if username<>"" then rs("username")=username
	if truename<>"" then rs("truename")=truename
	if password<>""  then rs("password")=md5(password)
	rs("editbox")=editbox
	session("user_editbox")= editbox

    rs.update 
    rs.close 
set rs=nothing

if err=0 then call alert.reFresh("账号信息保存成功,请牢记!","?")

end if
%>

<script type="text/javascript">
function check(){
   if (document.password.username.value.length=="")
   {
		 alert("用户名不能为空!");
		 document.password.username.focus();
		 return false;
   }
   if (document.password.password.value!=document.password.password_again.value)
   {
		 alert("两次输入密码不一致，请重新输入！");
		 document.password.password.focus();
		 return false;
   }
   return true;
}
</script>
<body>
<table border="0" align="center" cellpadding="0" cellspacing="10" class="forum1">
<tr><td valign="top">

<%if session("user_power")=1 then%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop table">
<tr class="forumRaw"><td align="center">
<a class="btn btn-mini" href="?" <%if UserEdit<>"add" then response.Write("class=onstyle")%>>
<span class="ico icon-plus">&nbsp;</span>&nbsp;编辑我的账号</a>
<a class="btn btn-mini" href="?UserEdit=add" <%if UserEdit="add" then response.Write("class=onstyle")%>>
<span class="ico icon-list">&nbsp;</span>&nbsp;添加管理员</a>
</td></tr></table>
<%end if%>

<%if UserEdit="add" then%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop table">
 <form name="password" method="post" id="password" action="" onSubmit="return check()">
<tr class="forumRow">
<td width="35%" align="right">账号：</td>
<td width="65%"><input name="username" type="text" size="30"/></td>
</tr>

<tr class="forumRow">
<td align="right">称呼：</td>
<td><input name="truename" type="text" size="30" /></td>
</tr>		  

<tr class="forumRow">
<td align="right">密码：</td>
<td><input name="password" type="password" id="password" size="30" /></td>
</tr>

<tr class="forumRow">
<td align="right">确认密码：</td>
<td><input name="password_again" type="password" id="password_again" size="30" /></td>
</tr>

<tr class="forumRow" style="display:none">
<td align="right">编辑器：</td>
<td>

<select name="editbox" id="editbox">
<%
set rs=sysconn.execute("select * from sys_ebox")
    do while not rs.eof
%>
<option value="<%=rs("id")%>"><%=rs("title")%></option>
<%
    rs.movenext
    loop
set rs=nothing
%>
<!--<option value="0">Ewebeditor&nbsp;-&nbsp;(常用)</option>
<option value="1">Fckeditor&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;(支持IE8下运行)</option>
<option value="2">Ewebeditor&nbsp;-&nbsp;(支持Word图片粘贴)</option>
-->
</select>
</td></tr>

<tr class="forumRaw">
<td>&nbsp;<input name="edit" type="hidden" id="edit" value="add" /></td>
<td align="left"><%=articlesHtml.submitBtn("","保存信息")%></td>
</tr>
</form>
</table>
<%else%>

<% 
set rs=sysconn.execute("select * from sys_admin where id="&session("user_id"))
    rs.open exec,sysconn,1,1 
	if not rs.eof then
	
	db_title = "帐号管理"
%>

<!--#include file="../models_system/articles/articles_edit_head.asp"-->

<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop table">
<form name="password" method="post" id="password" action="" onSubmit="return check()">

<tr class="forumRow">
<td width="35%" align="right">账号：</td>
<td width="65%"><input name="username" type="text" value="<%=rs("username")%>" size="30" disabled="disabled"/></td></tr>

<tr class="forumRow">
<td align="right">称呼：</td>
<td><input name="truename" type="text" value="<%=rs("truename")%>" size="30" /></td>
</tr>		  

<tr class="forumRow">
<td align="right">密码：</td>
<td><input name="password" type="password" id="password" size="30" /></td>
</tr>

<tr class="forumRow">
<td align="right" >确认密码：</td>
<td ><input name="password_again" type="password" id="password_again" size="30" /></td>
</tr>

<tr class="forumRow" style="display:none">
<td align="right">编辑器：</td>
<td>
<%
dim editbox
editbox=rs("editbox")
if text.isNum(editbox) = false then editbox=0
%>
<select name="editbox" id="editbox">
<%
set rs=sysconn.execute("select * from sys_ebox")
    do while not rs.eof
%>
<option value="<%=rs("id")%>"><%=rs("title")%></option>
<%
    rs.movenext
    loop
set rs=nothing
%>
</select>
</td>
          </tr>

<tr class="forumRaw">
<td>&nbsp;<input name="edit" type="hidden" id="edit" value="edit" /></td>
<td align="left"><%=articlesHtml.submitBtn("","保存信息")%></td>
</tr>
</form>
</table>
<%
    end if
    rs.close
set rs=nothing
%>

<%end if%>



<%if session("user_power")=1 then%>
<table width="100%" border="0" align=center cellpadding=4 cellspacing=1 class="forum2">
<tr class="forumRaw">
<td width="50" align="center">ID</td>
<td width="150" align="left">&nbsp;管理员账号</td>
<td align="left">&nbsp;称呼</td>
<td width="50" align="center">操作</td>
</tr>
<% 
set rs=sysconn.execute("select * from sys_admin where power<>1 and super<>1")
    do while not rs.eof
%>
<tr class="forumRow">
<td width="50" align="center"><%=rs("id")%></td>
<td width="150" align="left" >&nbsp;<%=rs("username")%></td>
<td align="left" >&nbsp;<%=rs("truename")%></td>
<td align="center" ><a href="?delID=<%=rs("id")%>">删除</a></td>
</tr>
<%
    rs.movenext
    loop
set rs=nothing
%>
</table>
<%end if%>
</td></tr>
</table>


</body>