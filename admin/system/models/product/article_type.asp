<!--#include file="article_config.asp"-->
<body>
<%
'*****创建文章对象，并绑定数据库*****
articles.table = db_table
articles.title = db_title

'*****删除分类*****
dim class_b_del,class_s_del
class_b_del = request.QueryString("class_b_del")
class_s_del = request.QueryString("class_s_del")

if text.isNum(class_b_del) then call articles.delTypeB(class_b_del)
if text.isNum(class_s_del) then call articles.delTypeS(class_s_del)

'*****处理删除排序问题*****
dim add_id_b
    add_id_b = request.QueryString("add_id_b")
if text.isNum(add_id_b) then session("type_id") = add_id_b

'处理分类排序问题
dim order,id_b,id_s,row_now_id,at_order_id
order = request.QueryString("order")
id_b  = request.QueryString("id_b")
id_s  = request.QueryString("id_s")

'*****执行重新排序*****
if order<>"" and text.isNum(id_b) then
   row_now_id = id_b
   if text.isNum(id_s) then row_now_id = id_s
   set row_now=conn.execute("select * from "&db_table&"_type where id="&row_now_id)
	 if not row_now.eof then
		at_order_id= row_now("order_id")
		row_now_id = row_now("id")
		where = array( array("type_id",0) ) '重新排序
		if text.isNum(id_s) then where = array( array("type_id",id_b) )
		call articles.listReOrder("id","order_id",row_now_id,at_order_id,order,where,1)
	 else
		call alert.reFresh("参数有误!","?")
	 end if
end if


'*****处理添加分类问题*****
dim id,edit,title,order_id,pic,type_id
id   = request("id")
edit = request("edit") 
title   = request.form("title")
order_id= request.form("order_id")
pic     = request.form("pic")
type_id = request.form("type_id")

if text.isNum(order_id) = false then order_id = 0
if text.isNum(type_id) = false then type_id = 0

if edit<>"" then
	if edit="update" and text.isNum(id) then
	   editStr = sys_str_edit
	else
	   editStr = sys_str_add
	end if
	if title="" then call alert.reFresh("分类名称不能为空!","?")

	articles.table = db_table & "_type"
	articles.rsUpdate(id)
		'写入数据
		articles.rs("title") = title
		articles.rs("order_id") = order_id
		articles.rs("pic") = pic
		if edit="add" then articles.rs("type_id") = type_id
		session("type_id") = type_id
	articles.rsUpdateClose()
	
    call alert.reFresh("保存成功!","?")
end if
%>
<table border="0" align="center" cellpadding="0" cellspacing="10" class="forum1">
<TR><td>
<%if db_typeShow<>1 then%>
<table border="0" cellpadding="5" cellspacing="1" class="forum2 forumtop">
<form name="mclass_type_add" method="post" action="?edit=add">
<tr class="forumRaw">
<td width="70" align="center"><span class="mainTitle">添加</span></td>
<td align="center" style="padding:4px;">
<select name="type_id">
<option value="0">大类</option>
<%
set r_type=conn.execute("select * from "&db_table&"_type where type_id=0 order by order_id asc")
	while not r_type.eof
	  if cstr(session("type_id"))=cstr(r_type("id")) then
	     response.Write("<option value=""" & r_type("id") & """ selected>[" & r_type("title") & "] 小类</option>")
	  else
	     response.Write("<option value=""" & r_type("id") & """>[" & r_type("title") & "] 小类</option>")
	  end if
	r_type.movenext
	wend
	r_type.close
set r_type=nothing
%>
    </select>
<input name="title" type="text"/>
<button name="submit" type="submit" class="btn"><span class="icon icon-plus">&nbsp;</span>&nbsp;添加分类</button>
</td></tr>
</form>	
</table>
<%end if%>

<%	
set row_b=server.createobject("adodb.recordset") 
	row_b_sql="select * from "&db_table&"_type where type_id=0 order by order_id asc" 
	row_b.open row_b_sql,conn,1,3
    if row_b.eof then
	   response.Write "&nbsp;暂无"&db_title&"分类！"
    else
%>
<table border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop table">
<tr class="forumRaw forumtitle">
<td class="manage-ids" colspan="2">编号</td>
<td align="left">&nbsp;<%=db_title%>分类名称</td>
<td class="manage-type-order">序号</td>
<td class="manage-type-order">排序</td>
<td class="manage-edits">保存操作</td>
</tr>
<%
    do while not row_b.eof
	'重写排序
	r_b=r_b+1
	row_b("order_id")=r_b
	row_b.update()
%>
<form class="form-search" name="article_type<%=row_b("id")%>" method="post" action="?edit=update&id=<%=row_b("id")%>">  
<%if cstr(request("id"))=cstr(row_b("id")) then%>
<tr class="forumRow" onDblClick="submit();" title="双击可完成编辑!">
<td align="center"><%=row_b("id")%><input name="id" type="hidden" id="id" value="<%=row_b("id")%>" /></td>
<td align="center"><img src="<%=Rpath%>public/images/ico/add_class_s.gif" width="9" height="9" border="0" /></td>
<td align="left">
<input name="title" type="text" class="input2" id="title" value="<%=row_b("title")%>" />
<input name="pic" type="text" id="pic" value="<%=row_b("pic")%>" size="15" />
<%=articlesHtml.upImgBtn("浏览图片","article_type"&row_b("id")&".pic")%>
</td>
<td class="manage-type-order"><input name="order_id" type="text" value="<%=row_b("order_id")%>"></td>
<td align="center"><%=articlesHtml.typeArrow(row_b("id"),"")%></td>
<td align="center">

<div class="btn-group">
<button name="update2" type="submit" class="btn btn-mini button2"><span class="icon icon-ok">&nbsp;</span> 保存</button>

<%=articlesHtml.btnManage( "typeB_id="&row_b("id") , "管理" )%>
</div>

</td></tr>
<%else%>
<tr class="forumRow" onDblClick="getEdit(<%=row_b("id")%>);" title="双击可编辑该分类!">
<td width="40" align="center"><%=row_b("id")%><input name="id" type="hidden" id="id" value="<%=row_b("id")%>" /></td>
<td align="center"><img src="<%=Rpath%>public/images/ico/add_class_s.gif" width="9" height="9" border="0" /></td>
<td align="left">&nbsp;<%=row_b("title")%></td>
<td width="50" align="center" class="forumRow2 red" style="font-weight:bold"><%=row_b("order_id")%></td>
<td align="center"><%=articlesHtml.typeArrow(row_b("id"),"")%></td>
<td align="center">
<div class="btn-group">
<%=articlesHtml.btnDel( "class_b_del="&row_b("id") , "确定删除该分类？" )%>
<%=articlesHtml.btnManage( "typeB_id="&row_b("id") , "管理" )%>
</div>
</td></tr>
<%end if%> 
</form>
<%
set row_s=server.CreateObject("ADODB.Recordset")
    sql1="select * from "&db_table&"_type where type_id="&row_b("id")&" order by order_id asc,id desc"
	row_s.open sql1,conn,1,3
	r_s=0
    do while not row_s.eof
   '#### 重写排序 ####
	r_s=r_s+1
	row_s("order_id")=r_s
	row_s.update()
%>
<form name="article_m_type<%=row_s("id")%>" method="post" action="?edit=update">
<%if cstr(request("id"))=cstr(row_s("id")) then%>
<tr class="forumRow" onDblClick="submit();" title="双击可完成编辑!">
<td width="40" align="center">
<%=row_s("id")%><input name="id" type="hidden" id="id" value="<%=row_s("id")%>" />
</td>
<td width="20" align="center"><img src="<%=Rpath%>public/images/ico/type_ico.gif" /></td>
<td align="left">
<input name="title" type="text" class="input2" id="title" value="<%=row_s("title")%>" />
<input name="pic" type="text" id="pic" value="<%=row_s("pic")%>" size="15" />
<%=articlesHtml.upImgBtn("浏览图片","article_m_type"&row_s("id")&".pic")%>
</td>

<td class="manage-type-order"><input name="order_id" type="text" value="<%=row_s("order_id")%>"></td>
<td align="center"> <%=articlesHtml.typeArrow(row_b("id"),row_s("id"))%> </td>

<td align="center">
<div class="btn-group">
<button name="update2" type="submit" class="btn btn-mini button2"><span class="icon icon-ok">&nbsp;</span> 保存</button>
<a class="btn btn-success btn-mini update" href="article_manage.asp?typeB_id=<%=row_b("id")%>&typeS_id=<%=row_s("id")%>">
<span class="icon icon-edit icon-white">&nbsp;</span> 管理</a>
</div>
</td></tr>
<%else%>
<tr class="forumRow" onDblClick="getEdit(<%=row_s("id")%>);" title="双击可编辑该分类!">
<td width="40" align="center">
<%=row_s("id")%><input name="id" type="hidden" id="id" value="<%=row_s("id")%>" />
</td>
          
<td width="20" align="center"><img src="<%=Rpath%>public/images/ico/type_ico.gif" /></td>
<td align="left">&nbsp;<%=row_s("title")%></td>
<td width="50" align="center"><%=row_s("order_id")%></td>
<td align="center"><%=articlesHtml.typeArrow(row_b("id"),row_s("id"))%> </td>

<td align="center">
<div class="btn-group">
<%=articlesHtml.btnDel( "class_s_del="&row_s("id") , "确定删除该分类？" )%>
<%=articlesHtml.btnManage( "typeB_id="&row_b("id")&"&typeS_id="&row_s("id") , "管理" )%>
</div>
</td></tr>
<%end if%> 
</form>
<%
		row_s.movenext:loop
		row_s.close
	set row_s=nothing
	
	row_b.movenext:loop
	end if
	row_b.close
set row_b=nothing


'//关闭对象
set articles = nothing
%>

</table>
</td></TR></table>
<!--进入编辑状态-->
<script>function getEdit(id){window.location.href='?id='+id;}</script>
</body>
</html>