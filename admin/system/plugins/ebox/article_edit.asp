<!--#include file="article_config.asp"-->
<%
'//创建文章对象，并绑定数据库
set articles = New articlesClass
    articles.thisConn = conn
    articles.table = db_table
    articles.title = db_title
%>
<%
'<><><><><>处理提交数据部分<><><><><>
dim edit_id
   edit_id=request("id")
if text.isNum(edit_id) then
   editStr=sys_str_edit
else
   editStr=sys_str_add
end if

   
if request.Form("edit")="ok" then
   tf_title=request.form("tf_title")
   tf_box  =request.form("tf_box")
   tf_box_1=request.form("tf_box_1")
   tf_add_data=request.form("tf_add_data")
   if isdate(tf_add_data)=false then tf_add_data=now()
   tf_order_id=request.form("tf_order_id")

'///////////  写入数据部分 //////////
set rs=server.createobject("adodb.recordset") 
    if text.isNum(edit_id) then
	   exec="select * from "&db_table&" where id=" & edit_id  '判断，修改数据
       rs.open exec,conn,1,3
    else
       exec="select * from "&db_table                       '判断，添加数据
	   rs.open exec,conn,1,3
       rs.addnew
    end if
    if text.isNum(edit_id) and rs.eof then
	   response.Write(sys_tip_none)
    else
       rs("title")=tf_title
	   rs("box")  =tf_box
	   rs("box_1")=tf_box_1
	   rs("add_data")=tf_add_data
	   if text.isNum(tf_order_id) then rs("order_id")=tf_order_id
	   '<><>判断是否正确写入<><>
	   if err<>0 then call backPage(editStr&sys_tip_false&err.description,"article_edit.asp?id="&edit_id,0)
	end if
	rs.update
	rs.close
set rs=nothing
    call backPage(editStr&sys_tip_ok,"article_manage.asp",0)
	
end if

'<><><><><><>读取数据部分<><><><><><>
id=request.QueryString("id") 
if text.isNum(id) then
   set rs=sysconn.execute("select * from "&db_table&" where id="&id) 
	   if not rs.eof then
tf_title=rs("title")
tf_box  =rs("box")
tf_box_1=rs("box_1")
tf_order_id=rs("order_id")
if text.isNum(tf_order_id)=false then tf_order_id=0
	   end if
	   rs.close
   set rs=nothing  
end if
%>


<body>
<TABLE border="0" align="center" cellpadding="0" cellspacing="10" class="forum1">
<TR><td>

<table border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop table">
<tr class="forumRaw"><TD>

<a class="btn btn-mini" href="article_edit.asp">
<span class="icon icon-plus">&nbsp;</span>&nbsp;<%=db_title%>添加</a>
<a class="btn btn-mini" href="article_manage.asp">
<span class="icon icon-list">&nbsp;</span>&nbsp;<%=db_title%>管理</a>

&nbsp;<%=db_title%>&nbsp;<%=editStr%>

</td></TR>
</table>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="forum2 table">
<form name="FormE" method="post" action="">
<tr class="forumRow"><td align="right">标题：</td><td><input name="tf_title" type="text" id="tf_title" value="<%=tf_title%>" size="60" />

</td></tr>
<tr class="forumRow">
<td align="right">简单版：</td>
<td><%=articlesHtml.editBox("tf_box_1",tf_box_1,"","","")%></td>
</tr>
<tr class="forumRow">
<td align="right">普通版：</td><td><%=articlesHtml.editBox("tf_box",tf_box,"","","")%></td></tr>
<tr class="forumRow"><td align="right">录入时间：</td><td><input type="text" name="tf_add_data" id="tf_add_data" value="<%=tf_add_data%>" />

</td></tr>
<tr class="forumRow"><td align="right">排序：</td><td><input type="text" name="tf_order_id" id="tf_order_id" value="<%=tf_order_id%>" />


</td></tr>

<tr class="forumRaw">
<td width="88" align="right" valign="top">
<input name="id" type="hidden" value="<%=id%>" />
<input name="edit" type="hidden" value="ok" />
</td>
<td><%=articlesHtml.submitBtn("",editStr)%></td>
</tr>
</form>
</table></td>
</tr></table>
</body>
</html>