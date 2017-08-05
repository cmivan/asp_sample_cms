<!--#include file="../../system/core/initialize_system.asp"-->
<body>
<TABLE border="0" align="center" cellpadding="0" cellspacing="10" class="forum1">
<TR><td>

<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop table">
<tr class="forumRaw" ><td class="mainTitle"><span class="ico icon-list">&nbsp;</span>&nbsp;系统导航栏目管理</td></tr>
</table>
  
<%	
	set row_b=server.createobject("adodb.recordset") 
	    exec="select * from sys_nav where type_id=0 order by order_id asc" 
	    row_b.open exec,sysconn,1,1 
	 if row_b.eof then
		response.Write "&nbsp;暂无"&db_title&"分导航！"
	 else
%>

<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop table">
<form name="mclass_type_add" method="post" action="?edit=add">
<tr class="forumRow">
<td colspan="2" align="center" class="forumRow">&nbsp;</td>
<td class="forumRow"><select name="type_id" style="width:99%;">
<option value="0">添加大导航</option>
<%
dim r_type
set r_type=server.CreateObject("adodb.recordset")
	r_type.open "select * from sys_nav where type_id=0 order by order_id asc",sysconn,1,1
	while not r_type.eof
	  if cstr(session("type_id"))=cstr(r_type("id")) then
	  response.Write("<option value=""" & r_type("id") & """ selected>[" & r_type("title") & "] 小导航</option>")
	  else
	  response.Write("<option value=""" & r_type("id") & """>[" & r_type("title") & "] 小导航</option>")
	  end if
	r_type.movenext
	wend
	r_type.close
set r_type=nothing
%>
            </select></td>
            <td colspan="5" class="forumRow"><span class="red">(标题填写 &quot;-&quot; ,则为分割栏; &quot;#&quot; + 表名,则显示该表的所有项)</span></td>
        </tr>
        
<tr class="forumRow">
            <td colspan="2" align="center" class="forumRaw">&nbsp;</td>
            <td class="forumRaw"><input name="title" type="text" class="input2" id="title" style="width:97%;"/></td>
            <td align="center" class="forumRaw"><span class="forumRow2">
              <input name="url" type="text" class="input3" id="url" style="width:95%;" />
            </span></td>
            <td align="center" class="forumRaw"><select name="target" id="target">
              <option value="main">main</option>
              <option value="_blank">_blank</option>
              <option value="_top">_top</option>
              <option value="_parent">_parent</option>
            </select></td>
            <td width="40" align="center" class="forumRaw"><span class="forumRow2">
              <input name="show" type="checkbox" id="show" value="1" checked />
            </span></td>
            <td width="40" align="center" class="forumRaw"><input name="order_id" type="text" id="order_id" style="width:25px;" value="0" size="5" /></td>

<td class="manage-edits">
<button type="submit" name="submit" class="btn" style="width:100%;">
<span class="ico icon-plus">&nbsp;</span>&nbsp;添加导航</button>
</td></tr>
</form>	



<tr class="forumRaw">
<td colspan="2" align="center" class="forumRaw">编号</td>
<td align="center" class="forumRaw">标题</td>
<td align="center" class="forumRaw">地址</td>
<td align="center" class="forumRaw">打开</td>
<td width="40" align="center" class="forumRaw">显示</td>
<td width="40" align="center" class="forumRaw">排序</td>
<td class="manage-edits">修改操作</td>
</tr>


<%do while not row_b.eof%>
<form name="article_type" method="post" action="?edit=update">
<tr class="forumRow">
<td width="40" align="center">
<input name="id" type="hidden" id="id" value="<%=row_b("id")%>" />
<%=row_b("id")%></td>
<td width="20" align="center"><img src="<%=Rpath%>public/images/ico/add_class_s.gif" width="9" height="9" border="0" /></td>
<td align="left"><input name="title" type="text" class="input2 red" id="title" style="font-weight:bold;width:100px;" value="<%=row_b("title")%>" />  <input name="db_showSet" type="text" id="db_showSet" value="<%=row_b("db_showSet")%>"></td>
<td align="center"><input name="url" type="text" class="input3" id="url" value="<%=row_b("url")%>" style="width:95%;" /></td>
<td align="center"><span class="forumRow">
		      <select name="target" id="target">
                <option value="main" <%if row_b("target")="main" then%>selected<%end if%>>main</option>
                <option value="_blank" <%if row_b("target")="_blank" then%>selected<%end if%>>_blank</option>
                <option value="_top" <%if row_b("target")="_top" then%>selected<%end if%>>_top</option>
                <option value="_parent" <%if row_b("target")="_parent" then%>selected<%end if%>>_parent</option>
              </select>
		    </span></td>
		    <td align="center"><input name="show" type="checkbox" id="show" value="1" <%if row_b("show")=1 then%>checked<%end if%> /></td>
            <td align="center"><input style="background-color:#F7F7EE; width:25px;" name="order_id" type="text" value="<%=row_b("order_id")%>" size="8"></td>

<td class="manage-edits" id="<%=row_b("id")%>">
<div class="btn-group"><a class="btn btn-mini delete" href="javascript:if(confirm('确定要删除此导航栏吗？删除后不可恢复!')){window.location.href='?act=del&id=<%=row_b("id")%>';}else{return false;}"><span class="icon icon-remove">&nbsp;</span> 删除</a><button name="update" type="submit" class="btn btn-mini btn-success" id="update"><span class="icon icon-white icon-ok">&nbsp;</span> 保存</button>
</div></td>
</tr>
        </form>
		<%
			  sql1="select * from sys_nav where type_id="&row_b("id")&" order by order_id asc"
			  set row_s=server.CreateObject("ADODB.Recordset")
			  row_s.open sql1,sysconn,1,1
			  do while not row_s.eof
		%>
        
		<form name="article_m_type" method="post" action="?edit=update">
<tr class="forumRow">
<td width="40" align="center">
<input name="id" type="hidden" id="id" value="<%=row_s("id")%>" /><%=row_s("id")%></td>
          
<td align="center"><img src="<%=Rpath%>public/images/ico/type_ico.gif" /></td>
<td class="forumRow2"><input name="title" type="text" class="input2" id="title" value="<%=row_s("title")%>" /></td>
<td align="center"><input name="url" type="text" class="input3" id="url" value="<%=row_s("url")%>" style="width:95%;" /></td>
<td align="center"><span class="forumRow">
<select name="target" id="target">
<option value="main" <%if row_s("target")="main" then%>selected<%end if%>>main</option>
<option value="_blank" <%if row_s("target")="_blank" then%>selected<%end if%>>_blank</option>
<option value="_top" <%if row_s("target")="_top" then%>selected<%end if%>>_top</option>
<option value="_parent" <%if row_s("target")="_parent" then%>selected<%end if%>>_parent</option>
</select>
</span></td>

<td align="center"><input name="show" type="checkbox" id="show" value="1" <%if row_s("show")=1 then%>checked<%end if%> /></td>
<td align="center"><input style="background-color:#FAFAF5;width:25px;" name="order_id" type="text" value="<%=row_s("order_id")%>" size="8"></td>

<td class="manage-edits" id="<%=row_s("id")%>">
<div class="btn-group"><a class="btn btn-mini delete" href="javascript:if(confirm('确定要删除此导航栏吗？删除后不可恢复!')){window.location.href='?act=del&id=<%=row_s("id")%>';}else{return false;}"><span class="icon icon-remove">&nbsp;</span> 删除</a><button name="update" type="submit" class="btn btn-mini btn-success" id="update"><span class="icon icon-white icon-ok">&nbsp;</span> 保存</button>
</div></td>

</tr>
</form>

		<%
			row_s.movenext
			loop
			row_s.close
			set row_s=nothing
		%>	
<%
	row_b.movenext
	loop
    row_b.close
set row_b=nothing
%>
<%
	end if
%>
      </table>

</td></TR></TABLE>
</body>
<%
   id=request("id")
   edit=request("edit")
   
   title   =request.form("title")
   db_showSet   =request.form("db_showSet")
   order_id=request.form("order_id")
   type_id =request.form("type_id")
   url     =request.form("url")
   show    =request.form("show")
   target  =request.form("target")
   
IF order_id="" or isNumeric(order_id)=false then order_id=0
IF type_id="" or isNumeric(type_id)=false then type_id=0


if edit<>"" then
   if db_showSet<>"" then
      db_showSets=split(db_showSet,"|")
      if ubound(db_showSets)>=1 then db_table=db_showSets(1)
      if db_table="" then
      response.Write("<script language=javascript>alert('设置有误!');history.go(-1)</script>") 
      response.end()
      end if
   end if

   IF title=""  then 
      response.Write("<script language=javascript>alert('分导航名称不能为空!');history.go(-1)</script>") 
      response.end()
   End IF

set rs=server.createobject("adodb.recordset")
 if edit="update" and id<>"" and isnumeric(id) then
    editStr="修改"
    sql="select * from sys_nav where id="&id 
    rs.open sql,sysconn,1,3
 else
    editStr="添加"
    sql="select * from sys_nav"
    rs.open sql,sysconn,1,3
    rs.addnew
 end if

'######### 写入数据 #############
    rs("title")   =title
	rs("db_showSet") =db_showSet
    rs("db_table")   =db_table
    rs("order_id")=order_id
	rs("url")     =url
	rs("show")    =show
	rs("target")  =target
	
 if edit="add" then rs("type_id") =type_id
    session("type_id")=type_id

    rs.update
    rs.close
set rs=nothing

'Response.Write "<script>alert('成功"&editStr&"分导航！');window.location.href='?';</script>" 
Response.Write "<script>window.location.href='?';</script>" 

end if
%>






<%
'############ 删除相应的分导航 ###########
if request("act")="del" then
   id=Request.QueryString("id")
if id<>"" and isnumeric(id) and id<>0 then
   sql="delete * from sys_nav where id="&id&" or type_id="&id
   sysconn.execute(sql)
   'response.Write "<script>alert('成功刪除分导航!');window.location.href='?';</script>" 
   Response.Write "<script>window.location.href='?';</script>" 
end if	
end if
%>