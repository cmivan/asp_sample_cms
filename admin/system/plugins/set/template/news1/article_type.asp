<!--配置文件-->
<!--#include file="article_config.asp"-->
<!--进入编辑状态-->
<script>function getEdit(id){window.location.href='?id='+id;}</script>

<body>
<%
'################| 处理删除排序问题 |################
   add_id_b=request.QueryString("add_id_b")
if isnum(add_id_b) then session("type_id")=add_id_b
'################| 处理删除排序问题 |################
   class_b_del=request.QueryString("class_b_del")
   class_s_del=request.QueryString("class_s_del")
if isnum(class_b_del) then 
'-===================| 删除大类 |=====================-
  conn.execute("delete from "&db_table&"_type where id="&class_b_del)  
  '删除相应小类
  conn.execute("delete from "&db_table&"_type where type_id="&class_b_del)
  '删除相应文章
  conn.execute("delete from "&db_table&" where typeB_id="&class_b_del)
  del_Info="成功删除相应的大类、小类、文章"
  call backPage(del_info,"?",0)
end if

if isnum(class_s_del) then 
'-===================| 删除小类 |=====================-
  conn.execute("delete from "&db_table&"_type where id="&class_s_del)
  conn.execute("delete from "&db_table&" where typeS_id="&class_s_del)
  del_info="成功删除分类、相应的内容!"
  call backPage(del_info,"?",0)
end if


'################| 处理分类排序问题 |################
  order =request.QueryString("order")
  id_b  =request.QueryString("id_b")
  id_s  =request.QueryString("id_s")
  
  if order<>"" and isnum(id_b) then
     if isnum(id_s) and isnum(id_b) then
	    order_sql="select * from "&db_table&"_type where id="&id_s
	 else
	    order_sql="select * from "&db_table&"_type where id="&id_b
	 end if
	 
	 set row_now=conn.execute(order_sql)
     if not row_now.eof then
        row_now_order_id=row_now("order_id")
		row_now_id      =row_now("id")
	 else
		call backPage("参数有误!","?",0)
	 end if
	 '%%%%%%%%%%%%%%%%%%%%%%%
 	 if isnum(id_s) and isnum(id_b) then
    	orderStr=" and type_id="&id_b
	 else
    	orderStr=" and type_id=0"
 	 end if
	 
  end if


  if order="up" then
    '---------------------------------------------
set row_up=conn.execute("select * from "&db_table&"_type where order_id<"&row_now_order_id&orderStr&" order by order_id desc")
	if not row_up.eof then
       row_up_order_id=row_up("order_id")
	   conn.execute("update "&db_table&"_type set order_id="&row_now_order_id&" where order_id="&row_up_order_id&orderStr)
	   conn.execute("update "&db_table&"_type set order_id="&row_up_order_id&" where id="&row_now_id&orderStr)
    else
	  call backPage("排序已到上限!","?",0)
	end if
set row_up=nothing

  elseif order="down" then
  
set row_down=conn.execute("select * from "&db_table&"_type where order_id>"&row_now_order_id&orderStr&" order by order_id asc")
	if not row_down.eof then
       row_down_order_id=row_down("order_id")
       row_down_query ="update "&db_table&"_type set order_id="&row_now_order_id&" where order_id="&row_down_order_id&orderStr
       row_now_query  ="update "&db_table&"_type set order_id="&row_down_order_id&" where id="&row_now_id&orderStr
       conn.execute(row_down_query)
       conn.execute(row_now_query)
	else
       call backPage("排序已到下限!","?",0)
	end if
	row_down.close
set row_down=nothing
	 
  end if
  

'################| 处理添加分类问题 |################
   id=request("id")
   edit=request("edit")
   
   title   = request.form("title")
   order_id= request.form("order_id")
   pic     = request.form("pic")
   type_id = request.form("type_id")
   
IF isnum(order_id)=false then order_id=0
IF isnum(type_id)=false then type_id=0

if edit<>"" then
   IF title=""  then call errtips("分类名称不能为空!")
set rs=server.createobject("adodb.recordset")
 if edit="update" and isnum(id) then
    editStr=sys_str_edit
    sql="select * from "&db_table&"_type where id="&id 
    rs.open sql,conn,1,3
 else
    editStr=sys_str_add
    sql="select * from "&db_table&"_type"
    rs.open sql,conn,1,3
    rs.addnew
 end if

'######### 写入数据 #############
    rs("title")   = title
    rs("order_id")= order_id
	rs("pic")     = pic
	
 if edit="add" then rs("type_id") = type_id
    session("type_id")=type_id

    rs.update
    rs.close
set rs=nothing
    call backPage("保存成功!","article_type.asp",0)
end if 
%>




<TABLE border="0" align="center" cellpadding="0" cellspacing="10" class="forum1">
<TR><td>

<%if db_typeShow<>1 then%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="forum2 forumtop">
<form name="mclass_type_add" method="post" action="article_type.asp?edit=add">
<tr class="forumRaw">
<td width="70" align="center"><span class="mainTitle">添加</span></td>
<td width="80" align="center">
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
</td>
<td align="center" class="forumRow"><input name="title" type="text" size="20" style="width:100%"></td>
<td width="86" align="center" class="forumRow"><input name="order_id" type="text" value="0" size="5"></td>			
<td width="120" align="center" class="forumRow"><input name="submit" type="submit" class="button" value="添加分类"></td>
</tr>
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

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="forum2">
<tr class="forumRaw forumtitle">
<td colspan="2" align="center" style="width:70px;"><span class="mainTitle">编号</span></td>
<td align="left"><span class="mainTitle">&nbsp;<%=db_title%>分类名称</span></td>
<td colspan="2" align="center" style="width:85px;"><span class="mainTitle">排序</span></td>
<td align="center" style="width:120px;"><span class="mainTitle">保存操作</span></td>
</tr>


<%
    do while not row_b.eof
	'#### 重写排序 ####
	r_b=r_b+1
	row_b("order_id")=r_b
	row_b.update()
%>
<form name="article_type<%=row_b("id")%>" method="post" action="article_type.asp?edit=update&id=<%=row_b("id")%>">
         
<%if cstr(request("id"))=cstr(row_b("id")) then%>
<tr class="forumRow" onDblClick="submit();" title="双击可完成编辑!">
<td align="center"><%=row_b("id")%><input name="id" type="hidden" id="id" value="<%=row_b("id")%>" /></td>
<td align="center"><a href="?add_id_b=<%=row_b("id")%>"><img src="<%=Rpath%>Images/ico/add_class_s.gif" width="9" height="9" border="0" /></a></td>
<td align="left">
<input name="title" type="text" class="input2" id="title" value="<%=row_b("title")%>" />
<input name="pic" type="text" id="pic" value="<%=row_b("pic")%>" size="15" />
<input type=button value="浏览" onClick="showUploadDialog('image', 'article_type<%=row_b("id")%>.pic', '')" class="button" />
</td>
<td align="center" style="width:50px;">
<input style="background-color:#F7F7EE; text-align:center" name="order_id" type="text" value="<%=row_b("order_id")%>" size="4">
</td>
<td align="center" style="width:35px;">
<a href="?order=up&id_b=<%=row_b("id")%>"><img src="<%=Rpath%>Images/ico/up_ico.gif" width="12" height="12" border="0"></a>
<a href="?order=down&id_b=<%=row_b("id")%>"><img src="<%=Rpath%>Images/ico/down_ico.gif" width="12" height="12"></a>
</td>
<td align="center">
<input name="update2" type="submit" class="button2" value="保存" />
<input name="update3" type="button" class="button" id="update" value="管理" onClick="window.location.href='article_manage.asp?typeB_id=<%=row_b("id")%>'" />
</td></tr>
<%else%>

<tr class="forumRow" onDblClick="getEdit(<%=row_b("id")%>);" title="双击可编辑该分类!">
<td width="40" align="center"><%=row_b("id")%><input name="id" type="hidden" id="id" value="<%=row_b("id")%>" /></td>
<td width="20" align="center"><a href="?add_id_b=<%=row_b("id")%>"><img src="<%=Rpath%>Images/ico/add_class_s.gif" width="9" height="9" border="0" /></a></td>
<td align="left">&nbsp;<%=row_b("title")%></td>
<td width="50" align="center" class="forumRow2 red" style="font-weight:bold"><%=row_b("order_id")%></td>
<td align="center">
<a href="?order=up&id_b=<%=row_b("id")%>"><img src="<%=Rpath%>Images/ico/up_ico.gif" width="12" height="12" border="0"></a>
<a href="?order=down&id_b=<%=row_b("id")%>"><img src="<%=Rpath%>Images/ico/down_ico.gif" width="12" height="12"></a>
</td>

<td align="center">
<input name="delete" type="button" class="button" onClick="javascript:if(confirm('确定要删除此导航栏吗？删除后不可恢复!')){window.location.href='?class_b_del=<%=row_b("id")%>';}else{history.go(0);}" value="删除"/>
<input name="update3" type="button" class="button" id="update" value="管理" onClick="window.location.href='article_manage.asp?typeB_id=<%=row_b("id")%>'" /></td>
</tr>
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
<td width="20" align="center"><img src="<%=Rpath%>Images/ico/type_ico.gif" /></td>
<td align="left">
<input name="title" type="text" class="input2" id="title" value="<%=row_s("title")%>" />
<input name="pic" type="text" id="pic" value="<%=row_s("pic")%>" size="15" />
<input type=button value="浏览" onClick="showUploadDialog('image', 'article_m_type<%=row_s("id")%>.pic', '')" class="button" />
</td>

<td width="50" align="center">
<input style="background-color:#FAFAF5;text-align:center" name="order_id" type="text" value="<%=row_s("order_id")%>" size="4">
</td>
<td align="center">
<a href="?order=up&id_b=<%=row_b("id")%>&id_s=<%=row_s("id")%>">
<img src="<%=Rpath%>Images/ico/up_ico.gif" width="12" height="12"></a>
<a href="?order=down&id_b=<%=row_b("id")%>&id_s=<%=row_s("id")%>">
<img src="<%=Rpath%>Images/ico/down_ico.gif" width="12" height="12"></a>
</td>

<td align="center">
<input name="update" type="submit" class="button2" value="保存" />
<input name="update23" type="button" class="button" value="管理" onClick="window.location.href='article_manage.asp?typeB_id=<%=row_b("id")%>&typeS_id=<%=row_s("id")%>'" />
</td>
</tr>
<%else%>
<tr class="forumRow" onDblClick="getEdit(<%=row_s("id")%>);" title="双击可编辑该分类!">
<td width="40" align="center">
<%=row_s("id")%><input name="id" type="hidden" id="id" value="<%=row_s("id")%>" />
</td>
          
<td width="20" align="center"><img src="<%=Rpath%>Images/ico/type_ico.gif" /></td>
<td align="left">&nbsp;<%=row_s("title")%></td>
<td width="50" align="center"><%=row_s("order_id")%></td>
<td align="center">
<a href="?order=up&id_b=<%=row_b("id")%>&id_s=<%=row_s("id")%>">
<img src="<%=Rpath%>Images/ico/up_ico.gif" width="12" height="12"></a>
<a href="?order=down&id_b=<%=row_b("id")%>&id_s=<%=row_s("id")%>">
<img src="<%=Rpath%>Images/ico/down_ico.gif" width="12" height="12"></a>
</td>

<td align="center">
<input name="delete" type="button" class="button" onClick="javascript:if(confirm('确定要删除此导航栏吗？删除后不可恢复!')){window.location.href='?class_s_del=<%=row_s("id")%>';}else{history.go(0);}" value="删除"/>
<input name="update23" type="button" class="button" value="管理" onClick="window.location.href='article_manage.asp?typeB_id=<%=row_b("id")%>&typeS_id=<%=row_s("id")%>'" />
</td>
</tr>
<%end if%> 
</form>

<%
		row_s.movenext
		loop
		row_s.close
	set row_s=nothing
	
	
	row_b.movenext
	loop
	end if
	row_b.close
set row_b=nothing
%>
      </table>


</td></TR></TABLE>
</body>
</html>