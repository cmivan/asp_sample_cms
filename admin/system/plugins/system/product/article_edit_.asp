﻿<!--配置文件-->
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
type_id=request.form("type_id")
'//系判断是否为数组，再判断是否为数字(否则失败)
type_ids=split(type_id,",")
if ubound(type_ids)=1 then
   if isnum(type_ids(0)) and isnum(type_ids(1)) then
	  typeB_id=type_ids(0)
	  typeS_id=type_ids(1)
	  session("type_id")=type_id   '记录本次操作分类
   else
	  call errtips(sys_tip_errtype)
   end if
else
   if isnum(type_id) then
	  typeB_id=type_id
	  typeS_id=null
	  session("type_id")=type_id   '记录本次操作分类
   else
	  call errtips(sys_tip_errtype)
   end if
end if
  

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
       articles.rs("title")=tf_title
articles.rs("typeS_id") =typeS_id
	   
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
typeS_id  =rs("typeS_id")
'记录分类,用于分类下拉
if isnum(typeS_id) then
   session("type_id")=typeB_id
else
   session("type_id")=typeB_id&","&typeS_id
end if
if txt.isNum(tf_order_id)=false then tf_order_id=0
	   end if
	   rs.close
   set rs=nothing  
end if
%>


<script>
function CheckForm()
{
 
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
<tr class="forumRow"><td align="right">title：</td><td><input type="text" name="tf_title" id="tf_title" value="<%=tf_title%>" />

<IFRAME ID="eWebEditor1" SRC="../../public/editBox/ewebeditor/ewebeditor.asp?id=tf_content&style=s_light" FRAMEBORDER="0" SCROLLING="no" WIDTH="100%" HEIGHT="400"></IFRAME>




















<input type=button value="浏览图片" onClick="showUploadDialog('image', 'FormE.tf_big_pic', '')" class="button"/>




<input type=button value="浏览图片" onClick="showUploadDialog('image', 'FormE.tf_small_pic', '')" class="button"/>














































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