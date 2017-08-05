<!--配置文件-->
<!--#include file="article_config.asp"-->
<%
'<><><><><>处理提交数据部分<><><><><>
dim edit_id
   edit_id=request("id")
if isnum(edit_id) then
   editStr=sys_str_edit
else
   editStr=sys_str_add
end if
   
if request.Form("edit")="ok" then
  {接收数据}
  {数据判断}

'///////////  写入数据部分 //////////
set rs=server.createobject("adodb.recordset") 
    if isnum(edit_id) then
	   exec="select * from "&db_table&" where id="&edit_id  '判断，修改数据
       rs.open exec,conn,1,3
    else
       exec="select * from "&db_table                       '判断，添加数据
	   rs.open exec,conn,1,3
       rs.addnew
    end if
    if isnum(edit_id) and rs.eof then
	   response.Write(sys_tip_none)
    else
       {写入数据}
	   
	   '<><>判断是否正确写入<><>
	   if err<>0 then call backPage(editStr&sys_tip_false,"article_edit.asp?id="&edit_id,0)
	end if
	rs.update
	rs.close
set rs=nothing
    call backPage(editStr&sys_tip_ok,"article_manage.asp?typdB_id="&typdB_id&"&typdS_id="&typdS_id,0)
	
end if

'<><><><><><>读取数据部分<><><><><><>
id=request.QueryString("id") 
if isnum(id) then
   set rs=conn.execute("select * from "&db_table&" where id="&id) 
	   if not rs.eof then
{读取数据}
	   end if
	   rs.close
   set rs=nothing  
end if
%>


<script>
function CheckForm()
{
 {表单判断}
}
///返回提示框
function Tips(ThisF){
  ThisF.style.background="#FF6600";
}
</script>
<body>
<TABLE border="0" align="center" cellpadding="0" cellspacing="10" class="forum1">
<TR><td>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="forum2 forumtop">
<tr class="forumRaw forumtitle">
<td align="center"><strong>&nbsp;&nbsp;<%=db_title%> <%=editStr%></strong></td>
</tr></table>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="forum2">
<form name="FormE" method="post" action="" onSubmit="return CheckForm(this);">
{表单数据}
<tr class="forumRaw">
<td width="88" align="right" valign="top">
<input name="id" type="hidden" value="<%=id%>" />
<input name="edit" type="hidden" value="ok" />
</td>
<td><input name="submit" type="submit" value="确认<%=editStr%>" class="button" /></td>
</tr>
</form>
</table></td>
</tr></table>
</body>
</html>