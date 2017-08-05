<!--配置文件-->
<!--#include file="article_config.asp"-->
<%
 typeB_id=request.QueryString("typeB_id")
 typeS_id=request.QueryString("typeS_id")
%>
<body>
<TABLE border="0" align="center" cellpadding="0" cellspacing="10" class="forum1">
<tr><td>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="forum2 forumtop">
<form name="search" method="get" action="">
<tr class="forumRaw"><TD width="100%" align="center" class="mainTitle"><%=db_title%> 管理列表</td>
<TD align="center">
<table border="0" cellpadding="0" cellspacing="0" class="forum2">
<tr class="forumRaw2"><td><input name="keyword" type="text" id="keyword" style="font-size: 9pt" size="25" /></td>
<td align="center">
<input name="submit" type="submit" value="<%=db_title%>搜索" align="absMiddle" class="button" />
<input name="typeB_id" type="hidden" id="typeB_id" style="font-size: 9pt" value="<%=typeB_id%>" size="25" />
<input name="typeS_id" type="hidden" id="typeS_id" style="font-size: 9pt" value="<%=typeS_id%>" size="25" />
</td></tr>
</table>
</td></TR>
</FORM>
</table>

	  
<!--#include file="../../public/articles_manage_type.asp"-->
<%
'############### | 批量处理 | ###############
   edits     = request.Form("edits")
   edit_item = request.Form("edit_item")
   type_id   = request.Form("type_id")
   
if edits<>"" then
   edit_items=split(edit_item,",")
if ubound(edit_items)>=0 then
'-=------ 判断，读取到的数据符合
   for i=0 to ubound(edit_items)
   
   if edit_items(i)<>"" and isnumeric(edit_items(i)) then
	      '---------  开始执行 -------------
if edits="del" then
   '__________ 删除
       delexec="delete * from "&db_table&" where id="&edit_items(i)
       conn.execute delexec
   '__________ 审核
elseif edits="check" then
       checkexec="update "&db_table&" set ok=1 where id="&edit_items(i)
       conn.execute checkexec
   '__________ 未审核
elseif edits="not_check" then
       checkexec="update "&db_table&" set ok=0 where id="&edit_items(i)
       conn.execute checkexec
   '__________ 移动
elseif edits="move" then
	   type_ids=split(type_id,",")
	   if ubound(type_ids)=1 then
          if type_ids(0)<>"" and isnumeric(type_ids(0)) and type_ids(1)<>"" and isnumeric(type_ids(1)) then
	         checkexec="update "&db_table&" set typeB_id="&type_ids(0)&",typeS_id="&type_ids(1)&" where id="&edit_items(i)
             conn.execute checkexec
          else
			 errStr="选择分类有误!"
          end if  
       else
	      if type_id<>"" and isnumeric(type_id) then
             checkexec="update "&db_table&" set typeB_id="&type_id&",typeS_id=null where id="&edit_items(i)
             conn.execute checkexec
		  else
             errStr="选择分类有误!"
		  end if
	   end if

else
   errStr="请选择相应的操作!"
end if

end if

next

    '完成批量操作
     errStr="ok"

   else
     errStr="请选择要操作的文章!"
end if

end if



'--------------------------------------------
r_ok=request.QueryString("ok")
r_hot=request.QueryString("hot")
r_id=request.QueryString("r_id")

if r_id<>"" and isnumeric(r_id) then
   set rs=server.createobject("adodb.recordset") 
       exec="select * from "&db_table&" where id="&r_id
       rs.open exec,conn,1,3 
	   if not rs.eof then
	      
		  if r_ok<>"" and int(r_ok)=1 then
		     rs("ok")=0
		  elseif r_ok<>"" and int(r_ok)=0 then
		     rs("ok")=1
		  end if
		  
	      if r_hot<>"" and int(r_hot)=1 then
		     rs("hot")=0
		  elseif r_hot<>"" and int(r_hot)=0 then
		     rs("hot")=1
		  end if

	   rs.update
	   end if
	   rs.close
   set rs=nothing
end if
 
'---------------------------------------------
	
	'########## 处理接收到搜索字符的情况 #############
       keyword=request("keyword")
	if keyword<>"" then
	   keyword_sql =" and title like '%"&request("keyword")&"%'"
	   keyword_sql2=" where title like '%"&request("keyword")&"%'"
	end if
	
	'########## 定义排序字符 #############
	 dim order_sql
	     order_sql=" order by order_id asc,id desc"
    
	if typeS_id<>"" and isnumeric(typeS_id)then
	   exec="select * from "&db_table&" where typeS_id="&typeS_id&keyword_sql&order_sql
	elseif typeB_id<>"" and isnumeric(typeB_id)then 
	   exec="select * from "&db_table&" where typeB_id="&typeB_id&keyword_sql&order_sql
	else
	   exec="select * from "&db_table&keyword_sql2&order_sql
	end if

		

	'### 用于热门，审核 按钮链接,批量移动修改等
	 Dim FullUrl
  	     FullUrl="?typeB_id="&typeB_id&"&typeS_id="&typeS_id&"&page="&page&"&keyword="&keyword&""

	
		
set rs=server.createobject("adodb.recordset")
	rs.open exec,conn,1,1 
	if rs.eof then
%>

<table width="100%" cellpadding="0" cellspacing="1" class="forum2">
<tr class="forumRow"><td class="noInfo">暂无 <%=db_title%> 相应内容!</td></tr>
</table>
	
<%else%>

<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forum2">
<form name="news_category" method="post">
<tr class="forumRaw forumtitle">
<td width="50" align="center" >ID</td>
<td width="20" align="center" >&nbsp;</td>
<td>&nbsp;<%=db_title%>&nbsp;名称\标题</td>
<td width="18" align="center" >热</td>
<td width="18" align="center" >审</td>
<td width="80" align="center" >录入时间</td>
<td width="130" align="center" >修改操作</td>
</tr>



<%
	rs.PageSize =18 '每页记录条数
	iCount=rs.RecordCount '记录总数
	iPageSize=rs.PageSize
	maxpage=rs.PageCount 
	page=request("page")
	
	if not IsNumeric(page) or page="" then
	   page=1
	else
	   page=cint(page)
	end if
	
	if page<1 then
	   page=1
	elseif  page>maxpage then
	   page=maxpage
	end if
	
	rs.AbsolutePage=Page
	if page=maxpage then
	   x=iCount-(maxpage-1)*iPageSize
	else
	   x=iPageSize
	end if	
	
For i=1 To x
%>

<tr class="forumRow">
<td width="50" align="center"><%=rs("id")%></td>
<td width="20" align="center"><input name="edit_item" type="checkbox" class="delitem" value="<%=rs("id")%>" /></td>
<td>&nbsp;
<a href="article_edit.asp?id=<%=rs("id")%>">
<%if keyword<>"" then
 response.Write(replace(rs("title"),keyword,"<span style='color:#FF0000'>"&keyword&"</span>"))
else
 response.Write(rs("title"))
end if
%></a>
</td>
<td width="18" align="center">

<%if rs("hot")=1 then%>
<a href="<%=FullUrl&"&hot=1&r_id="&rs("id")%>" class="yes">√</a>
<%else%>
<a href="<%=FullUrl&"&hot=0&r_id="&rs("id")%>" class="no">×</a>
<%end if%>

</td><td width="18" align="center">

<%if rs("ok")=1 then%>
<a href="<%=FullUrl&"&ok=1&r_id="&rs("id")%>" class="yes">√</a>
<%else%>
<a href="<%=FullUrl&"&ok=0&r_id="&rs("id")%>" class="no">×</a>
<%end if%>

</td><td width="80" align="center">
<a title="<%=rs("add_data")%>"><%=year(rs("add_data"))%>-<%=month(rs("add_data"))%>-<%=day(rs("add_data"))%></a>
</td>
<td align="center">
<input name="delete" type="button" value="删除" onClick="javascript:if(confirm('确定要删除此<%=db_title%>吗？删除后不可恢复!')){window.location.href='article_manage.asp?act=del&id=<%=rs("id")%>';}else{history.go(0);}" class="button" />
<input type="button" name="update" value="修改" onClick="window.location.href='article_edit.asp?id=<%=rs("id")%>' " class="button"/>
</td></tr>
<%
	rs.movenext
	next
%>

<tr class="forumRaw">
<td align="center">&nbsp;</td>
<td align="center"><input type="checkbox" id="delsel" /></td>
<td colspan="5" align="left">

<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr class="forumRaw2"><td>
&nbsp;
<label><input name="edits" type="radio" id="edits" value="check" checked="checked"/> 审核&nbsp;</label>
<label><input name="edits" type="radio" id="radio" value="not_check"/> 未审核&nbsp;</label>
<label><input name="edits" type="radio" id="radio3" value="del"/> 删除&nbsp;</label>
<label><input name="edits" type="radio" id="radio2" value="move"/> 转移&nbsp;</label>

<select name="type_id" id="type_id">
<option>请选择相应的分类</option>
<%
	dim rsc
	'///######### 大类
Set rsc=Conn.Execute("select * from "&db_table&"_type where type_id=0 order by order_id asc")
	while not rsc.eof
	selected=""
	if cstr(type_id)=cstr(rsc("id")) then selected="selected"
	   response.Write("<option value='" & rsc("id") & "' "&selected&">" & rsc("title") & "</option>")
				
	'///######### 小类
	Set rsm=Conn.Execute("select * from "&db_table&"_type where type_id="&rsc("id")&" order by order_id asc")
		do while not rsm.eof
		selected=""
		if typeB_id&","&typeS_id=rsc("id")&","&rsm("id") then selected="selected"
		   response.write "<option style=color:#999999 value='" & rsc("id")&","&rsm("id") & "' "&selected&">  ├ "& rsm("title") &"</option>"
		rsm.movenext
		loop
		rsm.close
    Set rsm=nothing
					   
    rsc.movenext
    wend
	rsc.close
set rsc=nothing
%>
          </select>
      <input type="submit" name="Submit" value="执行批操作" class="button" /></td>
    </tr>
  
</table></td>
</tr>
			
</form>
</table>
				
<%end if%>

<!--#include file="../../public/articles_paging.asp"-->

</td></TR></TABLE>


<%
if errStr="ok" then
   if edits="del" then
      errStr="成功删除!"
   elseif edits="check" then
      errStr="成功审核!"
   elseif edits="not_check" then
      errStr="成功取消审核!"
   elseif edits="move" then
      errStr="成功移动!"
   end if
end if

if errStr<>"" then
   response.Write("<script>alert('"&errStr&"');window.location.href='"&FullUrl&"';</script>")
end if
%>

</body>
</html>
<%
if request("act")="del" then
	set rs=server.createobject("adodb.recordset")
	id=Request.QueryString("id")
	sql="select * from "&db_table&" where id="&id
	rs.open sql,conn,2,3
	rs.delete
	rs.update
	Response.Write "<script>alert('"&db_title&"刪除成功！');window.location.href='"&FullUrl&"';</script>" 
end if
%>