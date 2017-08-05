<select name="type_id">
<%
	'///######### 大类
Set rsc=Conn.Execute("select * from "&db_table&"_type where type_id=0 order by order_id asc")
	while not rsc.eof
	selected=""
	if cstr(session("type_id"))=cstr(rsc("id")) then selected="selected"
	   response.Write("<option value='" & rsc("id") & "' "&selected&">" & rsc("title") & "</option>")
	'///######### 小类
	Set rsm=Conn.Execute("select * from "&db_table&"_type where type_id="&rsc("id")&" order by order_id asc")
		do while not rsm.eof
		selected=""
		typeS = rsc("id")&","&rsm("id")
		if request.QueryString("id")<>"" then
		   if session("type_id")=typeS then selected="selected"
		elseif session("type_id")<>"" then
		   '用于保持添加内容后分类不变
		   if session("type_id")=typeS then selected="selected"
		end if
		response.write "<option style=color:#999999 value='" & typeS & "' "&selected&">  ├ "& rsm("title") &"</option>"
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