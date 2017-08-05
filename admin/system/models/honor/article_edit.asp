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
  tf_title=request.form("tf_title")
tf_content=request.form("tf_content")
tf_note=request.form("tf_note")
tf_add_data=request.form("tf_add_data")
if isdate(tf_add_data)=false then tf_add_data=now()
tf_order_id=request.form("tf_order_id")

  

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
       rs("title")=tf_title
rs("content")=tf_content
rs("note")=tf_note
rs("add_data")=tf_add_data
if isnum(tf_order_id) then rs("order_id")=tf_order_id

	   
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
tf_title=rs("title")
tf_content=rs("content")
tf_note=rs("note")
tf_add_data=rs("add_data")
tf_order_id=rs("order_id")
if isnum(tf_order_id)=false then tf_order_id=0

	   end if
	   rs.close
   set rs=nothing  
end if
%>


<script>
function CheckForm()
{
	return true;
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
<tr class="forumRaw"><td align="center"><strong>&nbsp;&nbsp;<%=db_title%> <%=editStr%></strong></td>
</tr></table>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="forum2">
<form name="FormE" method="post" action="" onSubmit="return CheckForm(this);">
<tr class="forumRow"><td align="right">标题：</td><td><input name="tf_title" type="text" id="tf_title" value="<%=tf_title%>" style="width:400px;" />

</td></tr>
<tr class="forumRow"><td align="right">内容：</td><td>
<textarea id="tf_content" name="tf_content" style="display:none" ><%=tf_content%></textarea>
<textarea id="tf_content___Config" style="display:none" ><%=tf_content%></textarea>
<iframe id="tf_content___Frame" src="../../public/editbox/fckeditor/editor/fckeditor.html?InstanceName=tf_content&amp;Toolbar=Default" width="100%" height="400" frameborder="0" scrolling="no"></iframe>
</td></tr>

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