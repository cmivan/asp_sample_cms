<!--#include file="article_config.asp"-->
<%
'//创建文章对象，并绑定数据库
set articles = New articlesClass
    articles.thisConn = sysconn
    articles.table = db_table
    articles.title = db_title
    '//删除文章(单项)
    articles.del()
%>
<body>
<table border="0" align="center" cellpadding="0" cellspacing="10" class="forum1">
<tr><td>

<table border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop table">
<tr class="forumRaw"><td align="left">

<a class="btn btn-mini" href="article_edit.asp">
<span class="icon icon-plus">&nbsp;</span>&nbsp;<%=db_title%>添加</a>
<a class="btn btn-mini" href="article_manage.asp">
<span class="icon icon-list">&nbsp;</span>&nbsp;<%=db_title%>管理</a>

&nbsp;<%=db_title%>&nbsp;管理列表&nbsp;

</td></TR></table>

<%
set cmrs = New data
    cmrs.thisConn = sysconn
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
<td>&nbsp;<%=db_title%>&nbsp;名称\标题</td>
<td class="manage-date">录入时间</td>
<td class="manage-edits">修改操作</td>
</tr>

<%do while not cmrs.eof%>
<tr class="forumRow">
<td class="manage-ids"><%=cmrs.rs("id")%></td>
<td>&nbsp;<a href="article_edit.asp?id=<%=cmrs.rs("id")%>"><%=text.keyColor( cmrs.rs("title") ,keywords)%></a></td>
<td class="manage-date"><%=date2.ymd( cmrs.rs("add_data") )%></td>
<td class="manage-edits" id="<%=cmrs.rs("id")%>">
<div class="btn-group"><a class="btn btn-mini delete" href="javascript:if(confirm('确定要删除此<%=db_title%>吗？删除后不可恢复!')){window.location.href='article_manage.asp?act=del&id=<%=cmrs.rs("id")%>';}else{return false;}"><span class="icon icon-remove">&nbsp;</span> 删除</a><a class="btn btn-success btn-mini update" href="article_edit.asp?id=<%=cmrs.rs("id")%>"><span class="icon icon-white icon-edit">&nbsp;</span> 修改</a></div>
</td></tr>
<%cmrs.nexts:loop:cmrs.close%>		
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
<%
	end if
    set cmrs = nothing
set articles = nothing
%>
</td></TR></TABLE>
</body>
</html>