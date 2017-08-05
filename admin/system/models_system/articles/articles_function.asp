<%
'//批量处理 articles_batch_edit
function articles_batch_edit(db_table)
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
					 sql = "delete * from "&db_table&" where id="&edit_items(i)
					 conn.execute sql
				 '__________ 审核
			  elseif edits="check" then
					 sql = "update "&db_table&" set ok=1 where id="&edit_items(i)
					 conn.execute sql
				 '__________ 未审核
			  elseif edits="not_check" then
					 sql = "update "&db_table&" set ok=0 where id="&edit_items(i)
					 conn.execute sql
				 '__________ 移动
			  elseif edits="move" then
				 type_ids = split(type_id,",")
				 if ubound(type_ids)=1 then
					if text.isNum( type_ids(0) ) and text.isNum( type_ids(1) ) then
					   sql = "update "&db_table&" set typeB_id="&type_ids(0)&",typeS_id="&type_ids(1)&" where id="&edit_items(i)
					   conn.execute sql
					else
					   errStr = "选择分类有误!"
					end if  
				 else
					if text.isNum( type_id ) then
					   sql = "update "&db_table&" set typeB_id="&type_id&",typeS_id=null where id="&edit_items(i)
					   conn.execute sql
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











'//设置热门或者审核 articles_hot_ok_setting
function articles_set_check(db_table)
	dim r_ok,r_hot,r_id
	r_ok = request.QueryString("ok")
	r_hot= request.QueryString("hot")
	r_id = request.QueryString("r_id")
	if text.isNum(r_id) then
	   set rs=server.createobject("adodb.recordset") 
		   exec="select * from "&db_table&" where id = "&r_id
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
end function











'//设置热门或者审核 articles_hot_ok_setting
function articles_del(db_table)
	if request("act")="del" and text.isNum( request.QueryString("id") ) then
		set rs = conn.execute("delete from "&db_table&" where id=" & request.QueryString("id") )
		alert.msgReload("成功删除")
	end if
end function


%>