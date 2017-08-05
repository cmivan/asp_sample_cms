<%
'*********************
'<><>< URL类 ><><>
'*********************

Class articlesClass

  public thisConn
  public table
  public title
  public rs
  public typeB_id
  public typeS_id
  

  Private Sub Class_Initialize
	  set thisConn = conn
	  typeB_id = 0
	  typeS_id = 0
  End Sub


  '***********************************
  '   检测表是否存在 
  '***********************************
  public function checkTable() 
    if table="" then
       response.Write("出错!")
       response.End()
    end if
  End function
  
  

  '***********************************
  '   rsUpdate
  '***********************************
  public function rsUpdate(id)
      dim sql
	  call checkTable()
	  set rs=server.createobject("adodb.recordset")
	      sql = "select * from " & table
		  if text.isNum(id) then
			 sql = sql & " where id=" & id
			 rs.open sql,thisConn,1,3
		  else
			 rs.open sql,thisConn,1,3
			 rs.addnew
		  end if
	  set rsUpdate = rs
  end function

  '***********************************
  '   rsClose
  '***********************************
  public function rsUpdateClose()
	      rs.update
	      rs.close
	  set rs = nothing
  end function


  '***********************************
  '   显示指定记录
  '***********************************
  public function view(id)
     '创建数据对象
	 view = false
	 if text.isNum(id) then
	   set cmrs = New data
		   cmrs.where "id",id
		   cmrs.open table,1,false
		   if not cmrs.eof then set view = cmrs.rs
	   set cmrs = nothing
	 end if
  end function
  
 
  '***********************************
  '   显示指定记录 或 最新记录
  '***********************************
  public function pageView(id)
     '创建数据对象
	 pageView = false
	 set cmrs = New data
		 if text.isNum(id) and id<>0 then
		 cmrs.where "id",id
		 else
		 cmrs.orderBy "order_id","asc"
		 cmrs.orderBy "id","desc"
		 end if
		 cmrs.open table,1,false
		 if not cmrs.eof then set pageView = cmrs.rs
	 set cmrs = nothing
  end function

  
  '***********************************
  '   显示指定记录字段
  '***********************************
  public function viewKey(tables,id,key)
     '创建数据对象
	 viewKey = false
	 if tables="" then tables=table
	 if tables<>"" and text.isNum(id) then
	   set cmrs = New data
		   cmrs.where "id",id
		   cmrs.open tables,1,false
		   if not cmrs.eof then viewKey = cmrs.rs(key)
	   set cmrs = nothing
	 end if
  end function
  
  
  '***********************************
  '   输出分类列表 
  '***********************************
  public function typeItems(table,tid,oid,html)
      'on error resume next
	  dim style,thisHtml,newHtml
	  set rs=server.createobject("adodb.recordset") 
		  exec="select * from "&table&"_type where type_id="&tid&" order by order_id asc" 
		  rs.open exec,conn,1,1 
		  do while not rs.eof
			  style=""
			  if cstr(rs("id"))=cstr(oid) then style=" style=""font-weight:bold; color:#ff0000;"""
			  if tid=0 then
				 newHtml = "<a href="""&URI.reUrl("typeB_id="&rs("id")&"&typeS_id=null")&""" "&style&">&nbsp;"&rs("title")&"&nbsp;</a>"
			  else
				 newHtml = "<a href="""&URI.reUrl("typeS_id="&rs("id"))&""" "&style&">&nbsp;"&rs("title")&"&nbsp;</a>"
			  end if
			  thisHtml = thisHtml & replace(html,"{link}",newHtml)
		  rs.movenext
		  loop
		  rs.close
	  set rs=nothing
	  typeItems = thisHtml
  end function
  
  
  '***********************************
  '   文章分类框 
  '***********************************
  public function typeBox()
      on error resume next
	  dim newHtml,oid
	  newHtml = "<select name=""type_id"">"
	  oid = session("type_id")
	  '///######### 大类
	  Set rsc=conn.Execute("select * from "&table&"_type where type_id=0 order by order_id asc")
		  while not rsc.eof
		  selected=""
		  if cstr(oid)=cstr(rsc("id")) then selected="selected"
			 newHtml = newHtml & "<option value='" & rsc("id") & "' "&selected&">" & rsc("title") & "</option>"
		  '///######### 小类
		  Set rsm=conn.Execute("select * from "&table&"_type where type_id="&rsc("id")&" order by order_id asc")
			  do while not rsm.eof
			  selected=""
			  typeS = rsc("id")&","&rsm("id")
			  if request.QueryString("id")<>"" then
				 if oid=typeS then selected="selected"
			  elseif oid<>"" then
				 '用于保持添加内容后分类不变
				 if oid=typeS then selected="selected"
			  end if
			  newHtml = newHtml & "<option style=color:#999999 value='" & typeS & "' "&selected&">  ├ "& rsm("title") &"</option>"
			  rsm.movenext
			  loop
			  rsm.close
		  Set rsm=nothing				 
		  rsc.movenext
		  wend
		  rsc.close
	  set rsc=nothing
	  newHtml = newHtml & "</select>"
	  typeBox = newHtml
  end function
  
  

  '***********************************
  '   后台.编辑状态.获取分类
  '***********************************
  public function editGetTypes(type_id)
    on error resume next
    dim type_ids
	    type_ids = split(type_id,",")
	if ubound(type_ids)=1 then
	   if text.isNum(type_ids(0)) and text.isNum(type_ids(1)) then
		  typeB_id = type_ids(0)
		  typeS_id = type_ids(1)
		  session("type_id") = type_id   '记录本次操作分类
	   else
		  alert.msgBack(sys_tip_errtype)
	   end if
	else
	   if text.isNum(type_id) then
		  typeB_id = type_id
		  typeS_id = null
		  session("type_id") = type_id   '记录本次操作分类
	   else
		  alert.msgBack(sys_tip_errtype)
	   end if
	end if
  end function
  
  '***********************************
  '   后台.编辑状态.获取详情
  '***********************************
  public function viewGetTypes(typeB_id,typeS_id)
    on error resume next
	if text.isNum(typeS_id)=false then
	   session("type_id") = typeB_id
	else
	   session("type_id") = typeB_id & "," & typeS_id
	end if
  end function  



  
  '***********************************
  '   删除项目 
  '***********************************
  public function del()
    on error resume next
    call checkTable()
	dim delID
	delID = request.QueryString("id")
	if request("act")="del" and text.isNum( delID ) then
		set rs = server.createobject("adodb.recordset")
			rs.open "select * from "&table&" where id="&delID,thisConn,2,3
			rs.delete
			rs.update
			rs.close
		set rs = nothing
		if err<>0 then
		   call alert.msgTo("删除可能不成功","?")
		else
		   call alert.msgTo("成功删除","?")
		end if
	end if
  End function



  '***********************************
  '   设置热门或者审核 
  '***********************************
  public function setCheck()
    on error resume next
    call checkTable()
	dim r_cmd,r_val,r_id
	r_cmd = request.QueryString("cmd")
	r_val = request.QueryString("val")
	r_id  = request.QueryString("r_id")
	r_cmd = lcase(r_cmd)
	if text.isNum(r_id) and (r_cmd="ok" or r_cmd="hot" or r_cmd="news" or r_cmd="tj") then
	   set rs=server.createobject("adodb.recordset") 
		   exec="select * from "&table&" where id = " & r_id
		   rs.open exec,thisConn,1,3 
		   if not rs.eof then
		      r_val = text.getBool(r_val)
			  if int(r_val) = 1 then
				 rs(r_cmd) = 0
			  elseif int(r_val) = 0 then
				 rs(r_cmd) = 1
			  end if
		   rs.update
		   end if
		   rs.close
	   set rs=nothing
	end if
  end function



  '***********************************
  '   批量处理 
  '***********************************
  public function batchEdit()
    on error resume next
    call checkTable()
	dim edits,edit_item,type_id
	edits     = request.Form("edits")
	edit_item = request.Form("edit_item")
	type_id   = request.Form("type_id")
	
	if edits<>"" then
	  edit_items = split(edit_item,",")
	  if ubound(edit_items)>=0 then
	  '__________判断，读取到的数据符合
		 for i=0 to ubound(edit_items)
		   if text.isNum( edit_items(i) ) then
			  '__________ 开始执行__________
			  if edits="del" then
				 '__________ 删除
					 sql = "delete * from "&table&" where id="&edit_items(i)
					 thisConn.execute sql
				 '__________ 审核
			  elseif edits="check" then
					 sql = "update "&table&" set ok=1 where id="&edit_items(i)
					 thisConn.execute sql
				 '__________ 未审核
			  elseif edits="not_check" then
					 sql = "update "&table&" set ok=0 where id="&edit_items(i)
					 thisConn.execute sql
				 '__________ 移动
			  elseif edits="move" then
				 type_ids = split(type_id,",")
				 if ubound(type_ids)=1 then
					if text.isNum( type_ids(0) ) and text.isNum( type_ids(1) ) then
					   sql = "update "&table&" set typeB_id="&type_ids(0)&",typeS_id="&type_ids(1)&" where id="&edit_items(i)
					   thisConn.execute sql
					else
					   errStr = "选择分类有误!"
					end if  
				 else
					if text.isNum( type_id ) then
					   sql = "update "&table&" set typeB_id="&type_id&",typeS_id=null where id="&edit_items(i)
					   thisConn.execute sql
					else
					   errStr = "选择分类有误!"
					end if
				 end if
			  else
				 errStr = "请选择相应的操作!"
			  end if
		   end if
		 next
		'完成批量操作
		 errStr = "ok"
	  else
		 errStr = "请选择要操作的项目!"
	  end if
	end if
	
	'返回提示信息
	if errStr="ok" then
	   select case edits
	   case "del"
		  errStr="成功删除!"
	   case "check"
		  errStr="成功审核!"
	   case "not_check"
		  errStr="成功取消审核!"
	   case "move"
		  errStr="成功移动!"
	   end select
	end if
	
	if errStr<>"" then alert.msgReload(errStr)
  end function




  
  
  '***********************************
  '   删除分类项目 
  '***********************************
  public function delTypeB(typeId)
    on error resume next
    call checkTable()
	'删除大类
	thisConn.execute("delete from "&table&"_type where id=" & int(typeId))  
	'删除相应小类
	thisConn.execute("delete from "&table&"_type where type_id=" & int(typeId))
	'删除相应文章
	thisConn.execute("delete from "&table&" where typeB_id=" & int(typeId))
	call alert.reFresh("成功删除相应的大类、小类、相应的内容!","?")
  End function
  
  public function delTypeS(typeId)
    on error resume next
    call checkTable()
	'删除小类
	thisConn.execute("delete from "&table&"_type where id=" & int(typeId))
	'删除相应文章
	thisConn.execute("delete from "&table&" where typeS_id=" & int(typeId))
	call alert.reFresh("成功删除分类、相应的内容!","?")
  End function




  '***********************************
  '   重新排序 
  '***********************************
  public function listReOrder(key,okey,id,oid,types,where,order_type)
    call checkTable()	
	if table<>"" and key<>"" and okey<>"" and id<>"" and oid<>"" then

	    dim typeTable
	    typeTable = table & "_type"
	
		'创建数据对象
		set cmrs = New data
		
		if isarray(where) then
		   for each whereVal in where
			  if isarray(whereVal) then cmrs.where whereVal(0),whereVal(1)
		   next
		end if

		'处理排序方式
		if order_type="" then order_type = 0
		
		dim orderStr,whereKey,orderKey
		'执行重新排序
		if types = "up" then
		   if order_type=0 then
		      whereKey = okey&" >"
			  orderKey = "asc"
		   else   
		      whereKey = okey&" <"
			  orderKey = "desc"
		   end if
		   orderStr = "排序已到上限!"
		elseif types = "down" then
		   if order_type=0 then
		      whereKey = okey&" <"
			  orderKey = "desc"   
		   else
		      whereKey = okey&" >"
			  orderKey = "asc"
		   end if
		   orderStr = "排序已到上限!"
		end if
		
		cmrs.where whereKey,oid
		cmrs.orderBy okey,orderKey
		cmrs.open typeTable,1,false
		if not cmrs.eof then
		  at_id = cmrs.rs(key)
		  at_order_id = cmrs.rs(okey)
		  '创建数据对象
		  set dn1Rs = New data
			  dn1Rs.sets okey,oid
			  dn1Rs.where key,at_id
			  dn1Rs.update typeTable
		  set dn1Rs = nothing
		  '------------------
		  set dn2Rs = New data
			  dn2Rs.sets okey,at_order_id
			  dn2Rs.where key,id
			  dn2Rs.update typeTable
		  set dn2Rs = nothing
		  call alert.reFresh("更新成功!","?")
		else
		  call alert.reFresh(orderStr,"?")
		end if		
		
	end if

  end function
  
  
  

  
  '***********************************
  '   无内容 
  '***********************************
  public Sub noContent()
	response.Write("<table width=""100%"" cellpadding=""0"" cellspacing=""1"" class=""forum2"">")
	response.Write("<tr class=""forumRow""><td class=""noInfo"">暂无"&title&"相应内容!</td></tr></table>")
  End Sub
  



End Class



'*********************
'<><>< 实例化对象 ><><>
'*********************
set articles = New articlesClass
%>