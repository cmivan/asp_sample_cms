<!--#include file="article_config.asp"-->
<!--#include file="../../libraries/articles.asp"-->
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
<!--#include file="../../models_system/articles/articles_manage_type.asp"-->
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
<td class="manage-oks">审</td>
<td class="manage-oks">热</td>
<td class="manage-edits">修改操作</td>
</tr>

<%do while not cmrs.eof%>

<tr class="forumRow">
<td class="manage-ids"><%=cmrs.rs("id")%></td>
<td class="manage-check"><input name="edit_item" type="checkbox" class="delitem" value="<%=cmrs.rs("id")%>" /></td>
<td>&nbsp;&nbsp;<a href="article_edit.asp?id=<%=cmrs.rs("id")%>"><%=text.keyColor( cmrs.rs("title") ,keywords)%></a></td>

<td class="manage-oks">
<%=articlesHtml.btnYesNo("ok",cmrs.rs("ok"),cmrs.rs("id"))%>
</td>

<td class="manage-oks">
<%=articlesHtml.btnYesNo("hot",cmrs.rs("hot"),cmrs.rs("id"))%>
</td>

<td align="center" class="manage-edits" id="<%=cmrs.rs("id")%>">
<div class="btn-group">
<%=articlesHtml.btnDel( "?act=del&id="&cmrs.rs("id") , "" )%>
<%=articlesHtml.btnEdit( "article_edit.asp?id="&cmrs.rs("id") , "" )%>
</div>
</td></tr>

<%cmrs.nexts:loop:cmrs.close%>

<tr class="forumRaw">
<td align="center">&nbsp;</td>
<td class="manage-check"><input type="checkbox" id="delsel" /></td>
<td colspan="5" align="left">
<table border="0" cellpadding="0" cellspacing="2" id="manage-edit-box"><tr><td>
<select name="manage-edit-action" id="manage-edit-action">
<option value="del">删除</option>
</select>
</td></tr></table>

</td></tr>			
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
%>
<!--#include file="../../models_system/articles/articles_paging.asp"-->
				
<%end if%>

</td></tr></table>
</body></html>