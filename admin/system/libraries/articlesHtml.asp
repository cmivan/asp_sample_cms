<%
'*********************
'<><>< URL类 ><><>
'*********************

Class articlesHtmlClass

  public table
  public title

  Private Sub Class_Initialize
  
  End Sub


  '***********************************
  '   无内容 
  '***********************************
  Public Sub noContent()
	response.Write("<table width=""100%"" cellpadding=""0"" cellspacing=""1"" class=""forum2"">")
	response.Write("<tr class=""forumRow""><td class=""noInfo"">暂无"&title&"相应内容!</td></tr></table>")
  End Sub
  
  
  '***********************************
  '   审核或是否热门按钮
  '***********************************
  Public Sub btnYesNo(key,keyVal,idVal)
	if keyVal=1 then
	   response.Write("<a href="""&URI.reUrl("cmd="&key&"&val=1&r_id="&idVal)&""" class=""yes"">√</a>")
	else
	   response.Write("<a href="""&URI.reUrl("cmd="&key&"&val=0&r_id="&idVal)&""" class=""no"">×</a>")
	end if
  
'	if keyVal=1 then
'	   response.Write("<a href="""&URI.reUrl(""&key&"=1&r_id="&idVal)&""" class=""yes"">√</a>")
'	else
'	   response.Write("<a href="""&URI.reUrl(""&key&"=0&r_id="&idVal)&""" class=""no"">×</a>")
'	end if
  End Sub
  
  '***********************************
  '   审核或是否热门按钮
  '***********************************
  Public Sub checkYesNo(key,keyVal,Val)
    on error resume next
	if Val="" then Val = "是否"
	if cstr(keyVal)="1" then
	   response.Write("<label><input style=""width:auto;display:inline-block;"" id="""&key&""" name="""&key&""" type=""checkbox"" value=""1"" checked/>&nbsp;"&Val&"</label>")
	else
	   response.Write("<label><input style=""width:auto;display:inline-block;""  id="""&key&""" name="""&key&""" type=""checkbox"" value=""1""/>&nbsp;"&Val&"</label>")
	end if
  End Sub
  
  '***********************************
  '   分类的上下移动箭头 
  '***********************************
  Public Sub typeArrow(id_b,id_s)
    dim ids,imgStr
    if (id_b<>"" and isnumeric(id_b)) then ids = "id_b=" & id_b
	if (id_s<>"" and isnumeric(id_s)) then ids = ids & "&id_s=" & id_s
	imgStr = " width=""12"" height=""12"" border=""0""></a>"
	response.Write("<a href=""" & URI.reUrl("order=up&" & ids) & """><img src="""&Rpath&"public/images/ico/up_ico.gif""" & imgStr)
	response.Write("&nbsp;")
	response.Write("<a href=""" & URI.reUrl("order=down&" & ids) & """><img src="""&Rpath&"public/images/ico/down_ico.gif""" & imgStr)
  End Sub
  
  
  '***********************************
  '   用于编辑页面 
  '***********************************
  Public Sub editBox(boxId,boxVal,boxSkin,boxW,boxH)
    if boxSkin="" then boxSkin = "Default"
	if boxW="" then boxW = "100%"
	if boxH="" then boxH = "400"
    response.Write("<textarea style=""display:none"" id="""&boxId&""" name="""&boxId&""" >"&boxVal&"</textarea>")
	response.Write("<textarea style=""display:none"" id="""&boxId&"___Config"" name="""&boxId&"___Config"" >"&boxVal&"</textarea>")
	response.Write("<iframe id="""&boxId&"___Frame"" src="""&Rpath&"public/editbox/fckeditor/editor/fckeditor.html?InstanceName="&boxId&"&amp;Toolbar="&boxSkin&""" width="""&boxW&""" height="""&boxH&""" frameborder=""0"" scrolling=""no""></iframe>")
  End Sub
  
  
  
  '***********************************
  '   普通多行文本框 
  '***********************************
  Public Sub noteBox(boxId,boxVal)
    if boxId="" then boxId = "note"
    response.Write("<textarea name="""&boxId&""" style=""width:60%"" rows=""4"" id="""&boxId&""">"&boxVal&"</textarea>")
  End Sub
  
  
  '***********************************
  '   确定提交按钮 
  '***********************************
  Public Sub submitBtn(Id,Val)
    if Id="" then Id = "submit"
	if Val="" then Val = "保存"
    response.Write("<button id="""&Id&""" type=""submit"" class=""btn""><i class=""icon-ok"">&nbsp;</i> 确认"&Val&"</button>")
  End Sub
  
  
  '***********************************
  '   删除按钮 
  '***********************************
  Public Sub btnDel(Url,Val)
    if Url="" then Url = "?"
	if Val="" then Val = "删除后不可恢复,确定要删除吗？"
    response.Write("<a class=""btn btn-mini cmbtn"" href=""javascript:void(0);"" title="""&Val&""" url="""&URI.reUrl(Url)&""">")
    response.Write("<span class=""icon icon-remove"">&nbsp;</span> 删除</a>")
  End Sub
  
  
  '***********************************
  '   修改按钮 
  '***********************************
  Public Sub btnEdit(Url,Val)
    if Url="" then Url = "javascript:void(0);"
	if Val="" then Val = "修改"
    response.Write("<a class=""btn btn-mini btn-success"" href=""article_edit.asp"&URI.reUrl(Url)&""">")
	response.Write("<span class=""icon icon-white icon-edit"">&nbsp;</span> "&Val&"</a>")
  End Sub
  
  
  '***********************************
  '   管理按钮 
  '***********************************
  Public Sub btnManage(Url,Val)
    if Url="" then Url = "javascript:void(0);"
	if Val="" then Val = "管理"
    response.Write("<a class=""btn btn-mini btn-success"" href=""article_manage.asp"&URI.reUrl(Url)&""">")
	response.Write("<span class=""icon icon-white icon-edit"">&nbsp;</span> "&Val&"</a>")
  End Sub

  
  '***********************************
  '   上传图片按钮 
  '***********************************
  Public Sub upImgBtn(Val,Ids)
    if Ids="" then Ids = "article_update.small_pic"
	if Val="" then Val = "浏览图片"
    response.Write("<button type=""button"" class=""btn"" onClick=""showUploadDialog('image', '"&Ids&"', '')"">")
	response.Write("<span class=""ico icon-picture"">&nbsp;</span>&nbsp;"&Val&"</button>")
  End Sub
  
  
  '***********************************
  '   上传文件按钮 
  '***********************************
  Public Sub upFileBtn(Val,Ids)
    if Ids="" then Ids = "article_update.small_pic"
	if Val="" then Val = "浏览文件"
    response.Write("<button type=""button"" class=""btn"" onClick=""showUploadDialog('file', '"&Ids&"', '')"">")
	response.Write("<span class=""ico icon-picture"">&nbsp;</span>&nbsp;"&Val&"</button>")
  End Sub
  
  
  '***********************************
  '   管理列表大标题 
  '***********************************
  Public Sub mainTitle(Val)
    if Val="" then Val = "管理"
    response.Write("<span class=""ico icon-list"">&nbsp;</span>&nbsp;"&Val)
  End Sub


End Class



'*********************
'<><>< 实例化对象 ><><>
'*********************
set articlesHtml = New articlesHtmlClass
%>