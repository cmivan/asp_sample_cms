<!--#include file="../../../system/core/initialize_system.asp"-->
<%
'<><><><><><><><><><><><><><><><><><><><><><><><><><><><>
'        <><> 程序设计 编写 ：凡人
'        <><> 联系qq       ：394716221
'<><><><><><><><><><><><><><><><><><><><><><><><><><><><>

'绑定操作数据对象
dim connBeaut,MMconn
set connBeaut=conn

'读取配置数据文件
on error resume next
dim ConfigConn,ConfigStr,ConfigMdb
    ConfigMdb="Config.mdb"           '数据库文件目录
    ConfigStr="DRIVER=Microsoft Access Driver (*.mdb);DBQ="&server.mappath(ConfigMdb)
set ConfigConn=Server.CreateObject("ADODB.Connection") 
    ConfigConn.Open ConfigStr
 If err then
    err.clear
    set ConfigConn = nothing
    response.write "请检查配置文件!"
    response.end
 end If


'用于规则数据绑定
set MMconn=ConfigConn    


'系统配置变量
dim web_info,web_template,web_cform,PageNum,web_root
dim db_table,Num,web_Num
    web_info     = "天方网站代码生成系统"         '系统信息
	web_template = "template/"                 '含分类的页面目录
    web_cform    = "tf_"                       '控件前缀
    PageNum      = 3                           '生成的页面、路径上限
	'web_root     = "wwwroot/../../system"      '网站主文件夹
	web_root     = "../../../system/models"      '网站主文件夹
	db_table     = lcase(request.querystring("Etable"))  '要操作的表

%>


<style>
.db_table a{width:160px;float:left;text-indent:4px;font-family:Verdana;font-size:13px;}
.db_table a:hover{background-color:#FF9900;color:#FFFFFF;text-decoration:none;}
.db_table a.on{background-color:#FF9900;color:#FFFFFF;font-weight:bold;}
.db_link  a:hover{background-color:#FFFFFF;color:#FF9900;text-decoration:none;font-weight:bold}
</style>
<title><%=web_info%>，数据库操作...</title>

<TABLE border="0" align="center" cellpadding="0" cellspacing="10" class="forum1">
<tr><td>
<%
dim sysPage
    sysPage=request.QueryString("sysPage")	
 if sysPage<>"" then session("sysPage")=sysPage
 if sysPage="main" then session("sysPage")=""
%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop table">
<tr class="forumRaw"><td>

<a class="btn btn-mini" href="?sysPage=main">
<span class="ico icon-folder-open">&nbsp;</span>&nbsp;代码生成器</a>
<a class="btn btn-mini" href="?sysPage=manage">
<span class="ico icon-list">&nbsp;</span>&nbsp;管理规则</a>
<a class="btn btn-mini" href="?sysPage=edit&Etype=add">
<span class="ico icon-plus">&nbsp;</span>&nbsp;添加规则</a>

</td></tr>
</table>
<%
select case session("sysPage")
  case ""
%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="forum2">
<tr class="forumRaw"><td colspan="2" style="padding-left:4px;">
<strong>○ <%=web_info%></strong> / QQ:394716221
</td></tr>
<tr class="forumRow2"><td width="200" valign="top">	
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="forum2 table db_table">
<%
'读取导航中设置的
set GNav=sysconn.execute("select * from sys_nav where type_id=0 and db_showSet<>''")
    do while not GNav.eof
    Arrdb_showSet=GNav("db_showSet")
    db_showSet=split(Arrdb_showSet,"|")
    if ubound(db_showSe)>=2 then
%>
<tr class="forumRow"><td>
<%
if request.querystring("Etable")="" then response.Redirect("?Etable="&db_showSet(1))
if request.querystring("Etable")=db_showSet(1) then

dim db_title
    db_title=db_showSet(0)
%>
<a href='javascript:void(0);' class="on">- <%=db_showSet(1)%>&nbsp;→(<%=db_title%>)</a>
<%else%>
<a href='?Etable=<%=db_showSet(1)%>'>+ <%=db_showSet(1)%></a>
<%end if%>	
</td></tr>

<%
'///////  录入字段
dim fileds_txt
    fileds_txt="0"
	
set get_filed_conn = server.createobject("adodb.recordset")
    get_filed_conn_str = "select * from "&db_showSet(1)
    get_filed_conn.open get_filed_conn_str,connBeaut,0,1
	for i=0 to get_filed_conn.fields.count-1
	   fileds_txt=fileds_txt&","&get_filed_conn.fields.item(i).name
       set AddRs=server.createobject("adodb.recordset")
           AddRsStr="select * from sys_db_fields where db_table='"&db_showSet(1)&"'"
           AddRsStr=AddRsStr&" and db_field='"&get_filed_conn.fields.item(i).name&"'"
           AddRs.open AddRsStr,MMconn,1,3
		   if AddRs.eof then
	          AddRs.addnew
		      AddRs("db_title")=get_filed_conn.fields.item(i).name
			  AddRs("db_field")=get_filed_conn.fields.item(i).name
			  AddRs("db_type") =get_filed_conn.fields.item(i).type
			  AddRs("db_table")=GNav("db_table")
		      AddRs.update
		   end if
           AddRs.close
       set AddRs=nothing
	next
	get_filed_conn.close
set	get_filed_conn=nothing

MMconn.execute("delete from sys_db_fields where db_table='"&db_showSet(1)&"' and not in ("&fileds_txt&")")

    GNav.movenext
end if
    loop
set GNav=nothing
%>
</table>
</td>
<td width="100%" valign="top">
<%
   dim Etable
   Etable=request.querystring("Etable")
if Etable<>"" then
   '### 获取数据表数据 ###
   set GNav=sysconn.execute("select * from sys_nav where db_table='"&db_table&"'")
       if not GNav.eof then db_template=GNav("template")
   set GNav=nothing

'### 读取相应的参数
%>
<form id="form1" name="form1" method="post" action="?Etable=<%=db_table%>">
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop table">

<tr class="forumRaw">
<td align="center" class="forumRow">选择 </td>
<td align="center" class="forumRow">序号</td>
<td class="forumRow">字段</td>
<td class="forumRow">字段描述</td>
<td width="60" align="center" class="forumRow">类型</td>
<td width="80" align="center" class="forumRow">输入框/代码</td>
<td align="center" class="forumRow">排序</td>
<td align="center" class="forumRow">主键</td>
<td align="center" class="forumRow">JS</td>
</tr>
				
<%
'///////  读取导航中设置的
if Etable<>"" then

'///// 重排 /////
db_ReID=request.QueryString("db_ReID")
db_oNum=request.QueryString("db_oNum")
if db_ReID<>"" and isNumeric(db_ReID) and db_oNum<>"" and isNumeric(db_oNum) then  
   MMconn.execute("update sys_db_fields set orderID="&db_oNum&" where id="&db_ReID)
   response.Redirect("?Etable="&Etable)
end if

set ReRs=server.createobject("adodb.recordset")
    ReRsStr="select * from sys_db_fields where db_table='"&Etable&"' order by orderID asc,id asc"
	ReRs.open ReRsStr,MMconn,1,3
    dim db_reOrderNum
	do while not ReRs.eof
	   db_reOrderNum=db_reOrderNum+1
	   ReRs("orderID")=db_reOrderNum
	   ReRs.update
	   ReRs.movenext
	loop
    ReRs.close
set ReRs=nothing

set GNav=MMconn.execute("select * from sys_db_fields where db_table='"&Etable&"' order by orderID asc,id asc")
    i=0
    do while not GNav.eof
%>				
<tr class="forumRow">
<td align="center"><input name="field_<%=i%>" type="checkbox" id="field_<%=i%>" class="delitem" value="<%=GNav("db_field")%>" checked="checked" /></td>
<td align="center"><%=GNav("id")%><input type="hidden" name="fieldID_<%=i%>" value="<%=GNav("id")%>" /></td>
<td class="forumRow" style="padding-left:5px;">
<%=GNav("db_field")%></td>
<td><input name="field_tip_<%=i%>" type="text" id="field_tip_<%=i%>" value='<%=GNav("db_title")%>' /></td>
<td align="center"><input name="field_type_<%=i%>" type="hidden" id="field_type_<%=i%>" value="<%=GNav("db_type")%>" />
<%=get_type(cstr(GNav("db_type")))%></td>
<td>
<select name="field_input_type_<%=i%>">
<%
 set rs=ConfigConn.execute("select * from sys_db_type order by db_title asc,orderID asc,id desc")
	 do while not rs.eof
if GNav("db_input")=rs("id") then
 response.Write("<option selected value="&rs("id")&">"&rs("db_title")&"</option>")
else
 response.Write("<option value="&rs("id")&">"&rs("db_title")&"</option>")
end if
	 rs.movenext
	 loop
 set rs=nothing
%>
</select></td>

<td align="center">
<input name="orderID" type="text" id="orderID" style="width:30px;" value="<%=GNav("orderID")%>" size="10" onchange="window.location.href='?Etable=<%=Etable%>&db_ReID=<%=GNav("id")%>&db_oNum='+this.value;"  />
</td>
<td align="center">
<input type="radio" name="field_input_key" value="<%=GNav("id")%>" <%if GNav("db_key")=1 then%>checked<%end if%> /></td>
<td align="center">
<input name="field_input_checkJS<%=i%>" type="checkbox" id="field_input_checkJS<%=i%>" value="1" <%if cstr(GNav("db_checkJS"))="1" then%>checked<%end if%> /></td>
</tr>
<%
    i=i+1
    GNav.movenext
    loop
set GNav=nothing
end if
%>

<tr class="forumRaw"><td align="center">
<input name="" type="checkbox" id="delsel" value="" checked="checked" />				  </td>
<td colspan="10" align="left" class="forumRow"><input type="hidden" name="db_max" value="<%=i%>" />

<input type="hidden" name="db_form" value="yes" />
<select name="db_template" id="db_template">
<%
 set rs=ConfigConn.execute("select * from sys_db_template order by orderid asc,id desc")
	 do while not rs.eof
	 if db_template=rs("id") then
	    response.Write("<option selected value="&rs("id")&">"&rs("db_title")&"</option>")
	 else
	    response.Write("<option value="&rs("id")&">"&rs("db_title")&"</option>")
	 end if
	 rs.movenext
	 loop
 set rs=nothing
%>
</select>
<button type="submit" name="button2" id="button2" class="btn"><span class="ico icon-ok">&nbsp;</span>&nbsp;生成 (Run!)</button>

</td></tr></table>
</form>
<%end if%>
</td></tr></table>
<%
case "edit"
    Etype    = request.querystring("Etype")
    Eid      = request.querystring("Eid")
    Ftype    = request.form("Ftype")
    ErrInfo  = "<div style='font-size:12px;'>访问出错...</div>"
if not isempty(Ftype) then
set Econn=server.createobject("adodb.recordset")
 if Etype="add" then
	Econn_str="select * from sys_db_type"
	Econn.open Econn_str,MMconn,1,3
	Econn.addnew
	BackTip="添加成功!"
elseif Etype="edit" and not isempty(Eid) and isNumeric(Eid) then
    Econn_str="select * from sys_db_type where id="&int(Eid)
	Econn.open Econn_str,MMconn,1,3
	BackTip="修改成功!"
else
	response.write(ErrInfo)
    response.end() 	
end if 

if not Econn.eof then

   Econn("db_title")   = cstr(request.form("s_db_title"))
   'Econn("db_input")=cstr(request.form("s_db_input"))
   Econn("db_checkJS") = cstr(request.form("s_db_checkJS"))
   Econn("db_request") = cstr(request.form("s_db_request"))
   Econn("db_write")   = cstr(request.form("s_db_write"))
   Econn("db_read")    = cstr(request.form("s_db_read"))
   Econn("db_form")    = cstr(request.form("s_db_form"))
   
   s_db_input = request.form("s_db_input")
   s_db_input = replace(s_db_input,"< textarea","<textarea")
   s_db_input = replace(s_db_input,"< /textarea","</textarea")
   
   Econn("db_input")=cstr(s_db_input)

   response.write("<script>alert('" & BackTip & "');</script>")
end if
	Econn.update
	Econn.close
set Econn= nothing
end if

'判断是修改，还是添加||| 显示数据读取部分
if Etype="edit" and not isempty(Eid) and isNumeric(Eid) then
set Rconn=MMconn.execute("select * from sys_db_type where id="&int(Eid))
 if not Rconn.eof then
    s_db_title    = Rconn("db_title")
    s_db_input    = Rconn("db_input")
    s_db_checkJS  = Rconn("db_checkJS")
    s_db_request  = Rconn("db_request")
    s_db_write    = Rconn("db_write")
    s_db_read     = Rconn("db_read")
    s_db_form     = Rconn("db_form")
	
s_db_input = replace(s_db_input,"<textarea","< textarea")
s_db_input = replace(s_db_input,"</textarea","< /textarea")

 else
	response.write(ErrInfo)
    response.end() 	   
 end if
set Rconn=nothing
elseif Etype="add" then
else
    response.write(ErrInfo)
    response.end()   
end if
%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop table">

<form id="E_form" name="E_form" method="post" action="?Etype=<%=Etype%>&Eid=<%=Eid%>" >
<tr class="forumRaw">
<td width="10%" align='right'>标题：</td><td width="87%">
<input name="s_db_title" type="text" value="<%=s_db_title%>" size="40" /></td></tr>
<tr class="forumRow"><td align='right' valign="top">接收数据：</td><td>
<textarea name="s_db_request" rows="10" style="width:99%" /><%=s_db_request%></textarea></td></tr><tr class="forumRow">
<td align='right' valign="top">写入数据：</td><td>
<textarea name="s_db_write" rows="10" style="width:99%" /><%=s_db_write%></textarea></td></tr><tr class="forumRow">
<td align='right' valign="top">读取数据：</td><td>
<textarea name="s_db_read" rows="10" style="width:99%" /><%=s_db_read%></textarea></td></tr>
<tr class="forumRow"><td align='right' valign="top">显示数据：</td><td>
<textarea name="s_db_input" rows="10" style="width:99%" /><%=s_db_input%></textarea></td></tr>
<tr class="forumRow"><td align='right' valign="top">表单数据：</td><td>
<textarea name="s_db_form" rows="10" style="width:99%" /><%=s_db_form%></textarea></td></tr>
<tr class="forumRow"><td align='right' valign="top">JS检测：</td><td>
<textarea name="s_db_checkJS" rows="10" style="width:99%" /><%=s_db_checkJS%></textarea></td></tr>
<tr class="forumRaw"><td align="right">&nbsp;</td><td>
<button type="submit" name="button" id="button" class="btn"><span class="ico icon-ok">&nbsp;</span>&nbsp;保存 (Save!)&nbsp;</button>
<input name="Ftype" type="hidden" value="<%=Etype%>" /></td>
</tr>
</form>
</table>



<%
case "manage"

   Did=request.querystring("Did")
   CopyID=request.querystring("CopyID")
   ReID=request.querystring("ReID")
   oNum=request.querystring("oNum")
   
if Did<>"" and isNumeric(Did) then
   MMconn.execute("delete * from sys_db_type where id="&int(Did))
   response.Redirect("?")
elseif CopyID<>"" and isNumeric(CopyID) then
   set Rs=MMconn.execute("select * from sys_db_type where id="&int(CopyID))
       if not Rs.eof then
       set AddRs=server.createobject("adodb.recordset")
           AddRsStr="select * from sys_db_type"
           AddRs.open AddRsStr,MMconn,1,3
	       AddRs.addnew
		   AddRs("db_title")   =Rs("db_title")&"(2)"
		   AddRs("db_input")   =Rs("db_input")
		   AddRs("db_checkJS") =Rs("db_checkJS")
		   AddRs("db_request") =Rs("db_request")
		   AddRs("db_write")   =Rs("db_write")
		   AddRs("db_read")    =Rs("db_read")
		   AddRs("db_form")    =Rs("db_form")
		   AddRs.update
           AddRs.close
       set AddRs=nothing
	   response.Redirect("?")
	   end if
   set Rs=nothing
   
elseif ReID<>"" and isNumeric(ReID) and oNum<>"" and isNumeric(oNum) then  
   MMconn.execute("update sys_db_type set orderID="&oNum&" where id="&ReID)
   response.Redirect("?")
end if

'///// 重排 /////
set ReRs=server.createobject("adodb.recordset")
    ReRsStr="select * from sys_db_type order by db_title asc,orderID asc,id desc"
	ReRs.open ReRsStr,MMconn,1,3
    dim reOrderNum
	do while not ReRs.eof
	   reOrderNum=reOrderNum+1
	   ReRs("orderID")=reOrderNum
	   ReRs.update
	   ReRs.movenext
	loop
    ReRs.close
set ReRs=nothing


set Lconn=MMconn.execute("select * from sys_db_type order by db_title asc,orderID asc,id desc")
%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop table">
<tr class="forumRaw">
<td>&nbsp;名称/标题</td>
<td width="60" align="center">排序</td>
<td class="manage-edits">修改操作</td>
</tr>
<%do while not Lconn.eof%>
<tr class="forumRow"><td>
&nbsp;<a href="?sysPage=edit&page=<%=page%>&Etype=edit&Eid=<%=Lconn("id")%>"><%=Lconn("db_title")%></a></td>
<td align="center"><input name="orderID" type="text" id="orderID" style="width:40px; text-align:center; color:#FF6600; font-weight:bold" value="<%=Lconn("orderID")%>" onchange="window.location.href='?ReID=<%=Lconn("id")%>&oNum='+this.value;" /></td>
<td class="manage-edits">

<div class="btn-group"><a class="btn btn-mini delete" href="javascript:if(confirm('是否确定删除?')){window.location.href='?page=<%=page%>&Did=<%=Lconn("id")%>';}else{return false;}"><span class="icon icon-remove">&nbsp;</span> 删除</a><a class="btn btn-success btn-mini update" href="?page=<%=page%>&CopyID=<%=Lconn("id")%>"><span class="icon icon-white icon-file">&nbsp;</span> 复制</a></div>

</td></tr>
<%
    Lconn.movenext
    loop
set Lconn =nothing
%>  
</table>
<%end select%>
</td></tr>
</TABLE>






<%
'///////////////////////////////////////////////////////////////////////
'- 返回信息框的   msgbox 函数
'- msg(str,Num) ,str="提示的内容" , Num:1=提示  2=提示、返回  3=提示、结束
'//////////////////////////////////////////////////////////////////////
sub msg(str,Num)
    select case Num
	       case 1
		   response.write("<script>alert('"&str&"');</script>")
		   case 2
		   response.write("<script>alert('"&str&"');history.go(-1);</script>")
		   response.end()
		   case 3
		   response.write("<script>alert('"&str&"');</script>")
		   response.end()
	end select
end sub


'## 返回模板目录
function getTemplatePath(id)
  if isnumeric(id) then
   set gt=MMconn.execute("select db_path from sys_db_template where id="&id)
       if not gt.eof then getTemplatePath=gt("db_path")
   set gt=nothing
  end if
end function


'///////// 返回表格<table>、<tr class="forumRow">、<td>、<div>../////////
function Totable(tablestr,t_type)
tablestr="<"&t_type&" class=""forumRow"">"&tablestr&"</"&t_type&">"&chr(13)
end function

'///////// 返回读取数据库的字段类型//////////////////////
function get_type(Num)
 set rs=ConfigConn.execute("select top 1 * from sys_db_ftype where db_Num="&Num)
	 if not rs.eof then
	    get_type=rs("db_tip")
	 else
	    get_type="none"
	 end if
 set rs=nothing
end function

'///////// 返回输入框方式的函数//////////////////////////////
function getCode(f_tip,f_name,db_filed,id) 
 dim show_input 
 set rs=ConfigConn.execute("select top 1 * from sys_db_type where id="&id)
  if not rs.eof then
	 show_input=rs(db_filed)
	 show_input=replace(show_input,"{tip}",f_tip)
	 show_input=replace(show_input,"{wc}",web_cform)
	 show_input=replace(show_input,"{name}",f_name)
	 show_input=replace(show_input,"{value}","<%|="&web_cform&f_name&"%|>")
	 show_input=replace(show_input,"{content}",web_cform&f_name)
	 show_input=replace(show_input,"%|","%")
  end if
 set rs=nothing
 getCode=show_input & chr(13)
end function

'///////// 字符替换，将特定的字符替换成要生成的网站代码////////////////
function strTocode(mybody)
	'-- 替换配置文件
	     mybody=replace(mybody,"{db_table}",db_table)
		 mybody=replace(mybody,"{db_title}",db_title)
	'-- 编辑页面的
	     mybody=replace(mybody,"{接收数据}",f_request)
		 mybody=replace(mybody,"{写入数据}",f_write)
		 mybody=replace(mybody,"{读取数据}",f_read)
	     mybody=replace(mybody,"{表单数据}",field_td)
		 mybody=replace(mybody,"{数据判断}",f_requestorm)
		 mybody=replace(mybody,"{表单判断}",field_js)
    '-- 列表页面的	
		 mybody=replace(mybody,"$admin_list_db_1$",admin_list_db_1)
		 mybody=replace(mybody,"$admin_list_db_2$",admin_list_db_2)
    '-- 详细页面的
	     mybody=replace(mybody,"$admin_detail_db_1$",admin_detail_db_1)
end function

'//////////////////////////////////////////////////////////////////
'  生成文件/目录 的函数(go)    sub web_create("文件名","文件内容","类型")
'//////////////////////////////////////////////////////////////////
function web_create(s_name,s_centent,s_type)
    on error resume next  '容错模式(防止内存不足)
    select case s_type
	       case "folder"                          '文件夹不存在则生成
		        folder_name=lcase(server.mappath(s_name))
	            folder_name=replace(folder_name,"\","/")
		        folder_name=split(folder_name,"/")
				i_paht=ubound(folder_name)
				if instr(folder_name(i_paht),".")<>0 then i_paht=i_paht-1
				
		    set fso=createobject("scripting.filesystemobject")
				for i=0 to i_paht
            		if folder_name(i)<>"" then
		      		   folder_names=folder_names&folder_name(i)&"/"
			   		if fso.folderexists(folder_names)=false then fso.createfolder(folder_names)
	        		end if 
				next
	        set fso=nothing 
			
		   case "file"                            '文件不存在则生成文件
		   
'               生成完整路径
		        temp_paths=lcase(s_name)
				temp_path =split(temp_paths,"/")
		        temp_path_str=temp_path(ubound(temp_path))
				full_path =replace(temp_paths,temp_path_str,"")
				call web_create(full_path,,"folder")

				set stm=server.createobject("adodb.stream") 
   				    stm.type=2 '以本模式读取 
   				    stm.mode=3 
    				stm.charset="utf-8"
    				stm.open
					stm.writetext s_centent 
    				stm.savetofile server.mappath(s_name),2 
    				stm.flush 
    				stm.close
				set stm=nothing

		   case "getfile"                        '读取文件内容
				dim str 
				set stm=server.createobject("adodb.stream") 
    				stm.type=2 '以本模式读取 
    				stm.mode=3 
    				stm.charset="utf-8"
    				stm.open
					stm.loadfromfile server.mappath(s_name) 
    				str=stm.readtext 
    				stm.close
				set stm=nothing 
   				    web_create=str 
					
		   case "copy"                        '复制文件
   				set fso=createobject("scripting.filesystemobject") 
   				set c=fso.getfile(s_name) '被拷贝文件位置
       				c.copy(s_centent)
   				set c=nothing
   				set fso=nothing
    end select
end function
%>





<%
'///////// 生成文件 /////////
'f_request    - $接收数据组$ 
'f_write      - $写入数据组$
'f_read       - $读取数据组$
'field_td     - $表单数据组$
'f_requestorm - $数据判断$
'field_js     - $数据判断$
if db_table<>"" and request.form("db_form")="yes" then
if request.form("db_max")<>"" and request.form("db_max")>0 then

   '###判断主键
	  db_key     =request.form("field_input_key")   '主键
	  db_template=request.form("db_template")       '模板
	  
   '###更新数据
   if isnumeric(db_template) then
      sysconn.execute("update sys_nav set template="&db_template&" where db_table='"&db_table&"'")
	  db_tempPath=getTemplatePath(db_template)&"/"
   end if
	  
   if db_key ="" then call msg("请选择主键!",2)
   

   
   
   redim web_file(PageNum)      '页面路径 数组
   redim web_page(PageNum)      '页面个数 数组
   '读取系统模板文件数据 web_file()
   web_Num=0
   web_template=web_template&db_tempPath
   Set oFso = Server.CreateObject("Scripting.FileSystemObject")
   Set oFout = oFso.GetFolder(Server.Mappath(web_template))
       For Each File In oFout.Files
           web_page(web_Num)=web_create(web_template&File.Name,,"getfile")
           web_file(web_Num)=web_root&"/"&db_table&"/"&File.Name
           web_Num=web_Num+1
       Next
   Set oFout = Nothing
   Set oFso = Nothing
   
   

   for g_Num=0 to request.form("db_max")
   
       fieldID  =request.form("fieldID_" & g_Num)                       '字段ID   
       sys_field=request.form("field_" & g_Num)                        '字段名称
       sys_field_show=request.form("field_tip_" & g_Num)              '显示名称
       sys_field_type=request.form("field_type_" & g_Num)              '字段类型
       sys_field_max=request.form("field_max_" & g_Num)                '最大长度
       sys_field_input_type=request.form("field_input_type_" & g_Num)  '输入框类型
       sys_field_check=request.form("field_input_checkJS" & g_Num) '是否js检测


       '###记录数据####
       if sys_field<>"" and fieldID<>"" and isnumeric(fieldID) then
       set AddRs=server.createobject("adodb.recordset")
           AddRsStr="select * from sys_db_fields where id="&fieldID
           AddRs.open AddRsStr,MMconn,1,3
           if not AddRs.eof then
              AddRs("db_title")  =sys_field_show
              AddRs("db_input")  =sys_field_input_type
              if cstr(sys_field_check)="1" then
	             sys_field_check=1
              else
	             sys_field_check=0
              end if
	          AddRs("db_checkJS")=sys_field_check
			  
              if cstr(fieldID)=cstr(db_key) then
	             AddRs("db_key")=1
              else
	             AddRs("db_key")=0
              end if
              AddRs.update
           end if
           AddRs.close
       set AddRs=nothing
       end if


if fieldID<>"" and isnumeric(fieldID) and cstr(fieldID)<>cstr(db_key) then 
   intype  = sys_field_input_type    '输入框类型
   f_name  = sys_field               '输入框name

'-编辑页面(1)-{接收数据}  
   f_request=f_request & getCode(sys_field,f_name,"db_request",intype)
'-编辑页面(2)-{写入数据}
   f_write=f_write & getCode(sys_field,f_name,"db_write",intype)
'-编辑页面(3)-{读取数据}
   f_read=f_read & getCode(sys_field,f_name,"db_read",intype)  
'-编辑页面(4)-{表单显示}
   DBinput=getCode("",f_name,"db_input",intype)
   field_td=field_td & getCode(sys_field_show,DBinput,"db_form",intype)
'-编辑页面(5)-{数据判断}
   f_requestorm=f_requestorm & getCode(sys_field,f_name,"db_checkCode",intype)
'-编辑页面(6)-{表单判断}
   if sys_field_check<>"" then
      field_js=field_js & getCode(sys_field,f_name,"db_checkJS",intype)	
   end if



'/////// 列表页面-$admin_list_db_1$
temp_sys_field=sys_field
call Totable(sys_field_show,"td")
     admin_list_db_1=admin_list_db_1&sys_field_show
	 
	 sys_field="<"&"%=Lconn(""" & sys_field & """)%"&">"
call Totable(sys_field,"td")
     admin_list_db_2=admin_list_db_2&sys_field

'/////// 详细页面-$admin_detail_db_1$
	 sys_field="<"&"%=detail_conn(""" & temp_sys_field & """)%"&">"
call Totable(sys_field,"td")	 
     temp_detail_db=sys_field_show&sys_field
call Totable(temp_detail_db,"tr")

admin_detail_db_1=admin_detail_db_1&temp_detail_db
		
end if
next
 
end if


'/////// 将页面 赋值模板，替换,生成
dim pageBody,BackTip
for Num=0 to PageNum
	pageBody=web_page(Num)
	call strTocode(pageBody)                             '执行替换
	call web_create(web_file(int(Num)),,"folder")        '生成目录
	call web_create(web_file(int(Num)),pageBody,"file")  '生成文件
	BackTip=BackTip&"Beaut -> "&web_file(Num)&"\n"       '生成文件的返回信息
next
    BackTip="Time : "&now()&"\n\n"&BackTip               '返回提示信息
    call msg(BackTip,2)
end if

%>





