<!--#include file="article_config.asp"-->
<%
articles.table = db_table

'<><><><><>处理提交数据部分<><><><><>
dim edit_id
    edit_id=request("id")
   
if text.isNum(edit_id) then
   editStr = sys_str_edit
else
   editStr = sys_str_add
end if
   
if request.Form("edit")="ok" then


'******接收数据部分******
	'——————分类
	if is_types then
		type_id=request.form("type_id")
		call articles.editGetTypes(type_id)
	end if
	
	tf_typeB_id = articles.typeB_id
	tf_typeS_id = articles.typeS_id
	

	
	'——————可变动(go)——————
	tf_title=request.form("tf_title")
	tf_content=request.form("tf_content")
	tf_note=request.form("tf_note")
	tf_big_pic=request.form("tf_big_pic")
	tf_small_pic=request.form("tf_small_pic")
	tf_add_data=request.form("tf_add_data")
	tf_order_id = request.form("tf_order_id")
	tf_hits=request.form("tf_hits")
	tf_ok=text.getBool(request.form("tf_ok"))
	tf_hot=text.getBool(request.form("tf_hot"))
	tf_news=text.getBool(request.form("tf_news"))
	tf_tj=text.getBool(request.form("tf_tj"))
	tf_color=request.form("tf_color")
	tf_strong=request.form("tf_strong")
	tf_size=request.form("tf_size")
	tf_toUrl=request.form("tf_toUrl")
	'——————可变动(end)——————


'******写入数据部分******
articles.rsUpdate( edit_id )

    if text.isNum(edit_id) and articles.rs.eof then
	  alert.msgBack( sys_tip_none )
    else
	
	  '——————可变动(go)——————
	  articles.rs("title")=tf_title
	  
	  if is_content then articles.rs("content")=tf_content
	  if is_note then articles.rs("note")=tf_note
	  if is_types then 
		  articles.rs("typeB_id") =tf_typeB_id
		  articles.rs("typeS_id") =tf_typeS_id
	  end if
	  if is_big_pic then articles.rs("big_pic")=tf_big_pic
	  if is_small_pic then articles.rs("small_pic")=tf_small_pic

	  if is_ok then articles.rs("ok")=text.getBool( tf_ok )
	  if is_hot then articles.rs("hot")=text.getBool( tf_hot )
	  if is_news then articles.rs("news")=text.getBool( tf_news )
	  if is_tj then articles.rs("tj")=text.getBool( tf_tj )
	  
	  articles.rs("add_data")=tf_add_data
	  articles.rs("order_id")=text.getNum( tf_order_id )
	  articles.rs("hits")=tf_hits
	  articles.rs("color")=tf_color
	  articles.rs("strong")=tf_strong
	  articles.rs("size")=tf_size
	  articles.rs("toUrl")=tf_toUrl
	  '——————可变动(end)——————
	  
	  if err<>0 then call alert.reFresh(Err.Description & editStr&sys_tip_false,"?id="&edit_id)
	end if
articles.rsUpdateClose()

    call alert.reFresh(editStr&sys_tip_ok,"article_manage.asp?typdB_id="&typdB_id&"&typdS_id="&typdS_id)
end if



'******读取数据部分******
set viewRS = articles.view( request.QueryString("id") )
if viewRS<>false then

	'——————可变动(go)——————
	tf_title=viewRS("title")
	
	if is_content then tf_content=viewRS("content")
	if is_note then tf_note=viewRS("note")
	'--------------------------
	if is_types then 
		tf_typeB_id  =viewRS("typeB_id")
		tf_typeS_id  =viewRS("typeS_id")
		'记录分类,用于分类下拉
		call articles.viewGetTypes(tf_typeB_id,tf_typeS_id)
	end if
	'--------------------------
	if is_big_pic then tf_big_pic=viewRS("big_pic")
	if is_small_pic then tf_small_pic=viewRS("small_pic")
	'--------------------------
	if is_ok then tf_ok=text.getBool( viewRS("ok") )
	if is_hot then tf_hot=text.getBool( viewRS("hot") )
	if is_news then tf_news=text.getBool( viewRS("news") )
	if is_tj then tf_tj=text.getBool( viewRS("tj") )
	
	tf_add_data=viewRS("add_data")
	tf_order_id = text.getNum( viewRS("order_id") )
	tf_hits= text.getNum( viewRS("hits") )
	tf_color=viewRS("color")
	tf_strong=viewRS("strong")
	tf_size=viewRS("size")
	tf_toUrl=viewRS("toUrl")
	'——————可变动(end)——————

end if

%>

<body>
<TABLE border="0" align="center" cellpadding="0" cellspacing="10" class="forum1">
<TR><td>

<!--#include file="../../models_system/articles/articles_edit_head.asp"-->

<table border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop table">
<form name="article_update" class="well form-vertical" method="post" action="">

<tr class="forumRow"><td align="right">标题：</td>
<td><input name="tf_title" type="text" id="tf_title" value="<%=tf_title%>"/></td>
</tr>

<%if is_types then '设置分类%>
<tr class="forumRow"><td align="right">分类：</td><td><%=articles.typeBox()%></td></tr>
<%end if%>


<%if is_small_pic then%>
<tr class="forumRow">
<td align="right" class="forumRow">预览图：</td><td>
<input name="tf_small_pic" type="text" id="tf_small_pic" value="<%=tf_small_pic%>" size="45" maxlength="255" />
<%=articlesHtml.upImgBtn("上传图片","article_update.tf_small_pic")%>
</td></tr>
<%end if%>


<%if is_big_pic then%>
<tr class="forumRow">
<td align="right" class="forumRow">详细图：</td><td>
<input name="tf_big_pic" type="text" id="tf_big_pic" value="<%=tf_big_pic%>" size="45" maxlength="255" />
<%=articlesHtml.upImgBtn("上传图片","article_update.tf_big_pic")%>
</td></tr>
<%end if%>


<%if is_toUrl then%>
<tr class="forumRow">
<td align="right">链接地址：</td>
<td><input name="tf_toUrl" type="text" id="tf_toUrl" value="<%=tf_toUrl%>" size="45" maxlength="255" /> 
(<span class="red"> 链接地址如: http://www.baidu.com</span> ) </td>
</tr>
<%end if%>


<%if is_note then%>
<tr class="forumRow"><td align="right">简要：</td><td>
<%=articlesHtml.noteBox("tf_note",tf_note)%>
</td></tr>
<%end if%>


<%if is_content then%>
<tr class="forumRow"><td align="right">内容：</td><td>
<%=articlesHtml.editBox("tf_content",tf_content,"","","")%>
</td></tr>
<%end if%>


<%if is_ok=true or is_hot=true or is_news=true or is_tj=true then%>
<tr class="forumRow">
  <td align="right">是否：</td><td>
  <%if is_ok then%><%=articlesHtml.checkYesNo("tf_ok",tf_ok,"通过")%><%end if%>
  <%if is_hot then%><%=articlesHtml.checkYesNo("tf_hot",tf_hot,"热门")%><%end if%>
  <%if is_news then%><%=articlesHtml.checkYesNo("tf_news",tf_news,"新品")%><%end if%>
  <%if is_tj then%><%=articlesHtml.checkYesNo("tf_tj",tf_tj,"推荐")%><%end if%>
</td></tr>
<%end if%>

<tr class="forumRaw">
<td width="88" align="right" valign="top">
  <input name="id" type="hidden" value="<%=id%>" />
  <input name="edit" type="hidden" value="ok" />
</td><td><%=articlesHtml.submitBtn("",editStr)%></td></tr>
</form>
</table></td>
</tr></table>
</body>
</html>