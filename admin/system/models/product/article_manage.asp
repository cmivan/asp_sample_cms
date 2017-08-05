<!--#include file="article_config.asp"-->
<%
dim typeB_id,typeS_id,keywords

	typeB_id = request.QueryString("typeB_id")
	typeS_id = request.QueryString("typeS_id")
	keywords = text.noSql(request("keywords"))

	'//创建文章对象，并绑定数据库
    articles.table = db_table
    articles.title = db_title
    '//设置文章是否被推荐等
    articles.setCheck()
    '//删除文章(单项)
    articles.del()
    '//批量操作(删除、移动、审核等)
    articles.batchEdit()
%>

<body>
<table border="0" align="center" cellpadding="0" cellspacing="10" class="forum1"><tr><td>
<!--#include file="../../models_system/articles/articles_manage_head.asp"-->

<%if is_types then%>
<!--#include file="../../models_system/articles/articles_manage_type.asp"-->
<%end if%>

<%
set cmrs = New data
	if text.isnum(typeB_id) then cmrs.where "typeB_id",typeB_id
	if text.isnum(typeS_id) then cmrs.where "typeS_id",typeS_id
	if keywords<>"" then cmrs.likes "title",keywords
	cmrs.orderBy "order_id","asc"
	cmrs.orderBy "id","desc"
	cmrs.open db_table,12,true
	if cmrs.eof then
	   articles.noContent() '没有文章内容
	else
%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop table">
<form name="news_category" method="post" class="form-search">

<tr class="forumRaw">
<td class="manage-ids">ID</td>
<td class="manage-check">&nbsp;</td>
<td>&nbsp;&nbsp;<%=db_title%>&nbsp;名称\标题</td>
<%if is_ok then%><td class="manage-oks">审</td><%end if%>
<%if is_hot then%><td class="manage-oks">热</td><%end if%>
<%if is_news then%><td class="manage-oks">新</td><%end if%>
<%if is_tj then%><td class="manage-oks">推</td><%end if%>
<td class="manage-edits">修改操作</td>
</tr>

<%do while not cmrs.eof%>

<tr class="forumRow">
<td class="manage-ids"><%=cmrs.rs("id")%></td>
<td class="manage-check"><input name="edit_item" type="checkbox" class="delitem" value="<%=cmrs.rs("id")%>" /></td>
<td>&nbsp;&nbsp;<a href="article_edit.asp?id=<%=cmrs.rs("id")%>"><%=text.keyColor( cmrs.rs("title") ,keywords)%></a></td>


<%if is_ok then%>
<td class="manage-oks">
<%=articlesHtml.btnYesNo("ok",cmrs.rs("ok"),cmrs.rs("id"))%>
</td>
<%end if%>


<%if is_hot then%>
<td class="manage-oks">
<%=articlesHtml.btnYesNo("hot",cmrs.rs("hot"),cmrs.rs("id"))%>
</td>
<%end if%>


<%if is_news then%>
<td class="manage-oks">
<%=articlesHtml.btnYesNo("news",cmrs.rs("news"),cmrs.rs("id"))%>
</td>
<%end if%>


<%if is_tj then%>
<td class="manage-oks">
<%=articlesHtml.btnYesNo("tj",cmrs.rs("tj"),cmrs.rs("id"))%>
</td>
<%end if%>

<td align="center" class="manage-edits" id="<%=cmrs.rs("id")%>">
<div class="btn-group">
<%if allow_del then%>
<%=articlesHtml.btnDel( "act=del&id="&cmrs.rs("id") , "" )%>
<%end if%>
<%=articlesHtml.btnEdit( "id="&cmrs.rs("id") , "" )%>
</div>
</td></tr>

<%cmrs.nexts:loop:cmrs.close%>

<%if allow_del then%>
<tr class="forumRaw">
<td align="center">&nbsp;</td>
<td class="manage-check"><input type="checkbox" id="delsel" /></td>
<td colspan="6" align="left">
<table border="0" cellpadding="0" cellspacing="2" id="manage-edit-box"><tr><td>
<select name="manage-edit-action" id="manage-edit-action">
<option value="del">删除</option>
</select>
</td></tr></table>
</td></tr>
<%end if%>
			
</form>
</table>

<%
'page
'maxPage
'allCount
dim page,pageCounts,allCounts
page       = cmrs.page
pageCounts = cmrs.pageCounts
allCounts  = cmrs.allCounts
    cmrs.close
set cmrs = nothing
%>
<!--#include file="../../models_system/articles/articles_paging.asp"-->
				
<%end if%>

</td></tr></table>
</body></html>