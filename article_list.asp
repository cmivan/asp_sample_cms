<!--#include file="system/core/initialize.asp"-->
<!--#include file="system/libraries/articles.asp"-->
<%
'***************'url'***************
'text.echo( URI.reUrl("id=12&action=yes") )
'response.Write("<hr>")
'text.echo( URI.encodeUrl("创建的目录") )
'response.Write("<hr>")
'call autoRedirect("www","5cmlabs.com")


'***************'text'***************
'text.echo( file2.folderCreate("./_创建的目录/") )
'response.Write("<hr>")
'text.echo( http.httpSave("http://www.5cmlabs.com","_创建的目录/_创建的文件.txt") )


'***************'alert'***************
'call alert.reFresh("跳转提示","http://www.5cmlabs.com")
'call alert.msgBack("跳转提示")
'call alert.msgTo("跳转提示","http://www.5cmlabs.com")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!--配置文件-->
<%
 db_table ="product"   '表段
 
 keyword = text.noSql(request("keyword"))
 typeB_id=request.QueryString("typeB_id")
 typeS_id=request.QueryString("typeS_id")
%>
<link href="public/assets/css/bootstrap.css" rel="stylesheet" type="text/css" />
<body><br />
<TABLE width="800" border="0" align="center" cellpadding="0" cellspacing="10" bgcolor="#FFFFFF" class="article_list">
<TR><td class="forumRow">
<Form name="search" method="get" action="" class="form-search">
<table border="0" cellpadding="0" cellspacing="0">
<tr><td><input name="keyword" type="text" id="keyword" value="<%=keyword%>" style="font-size: 9pt" size="25" /></td>
<td>&nbsp;</td>
<td align="center"><button type="submit" class="btn"><span class="icon-search"></span> 搜索</button>
<input name="typeB_id" type="hidden" id="typeB_id" style="font-size: 9pt" value="<%=typeB_id%>" size="25" />
<input name="typeS_id" type="hidden" id="typeS_id" style="font-size: 9pt" value="<%=typeS_id%>" size="25" /></td>
</tr></table>
</FORM>

<!--#include file="article_type.asp"-->

<%
response.Write(SC("seo.title","note"))
set myrs = New data

	if text.isnum(typeB_id) then myrs.where "typeB_id",typeB_id
	if text.isnum(typeS_id) then myrs.where "typeS_id",typeS_id
	if keywords<>"" then myrs.likes "title",keywords

	myrs.orderBy "order_id","asc"
	myrs.orderBy "id","desc"
	myrs.open db_table,12,true
	if not myrs.eof then
%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forMy forMAargin">
<%do while not myrs.eof%>		
<tr>
  <td>
<a href="article_view.asp<%=URI.reUrl("page=null&keyword=null&id="&myrs.rs("id"))%>" target="_blank">
<%if keyword<>"" then%>
  <%=text.keyColor( myrs.rs("title") , keyword )%>
<%else%>
  <%=text.strCut( myrs.rs("title") ,25)%>
<%end if%>
</a>
<!--热门，最新...-->
<%if myrs.rs("hot")=1 then%><span class="icon-search"></span><%end if%>
<%if myrs.rs("news")=1 then%><span class="icon-search"></span><%end if%>
</td>
<td width="80" align="left" class="forumRow2"><span title="<%=myrs.rs("add_data")%>"><%=date2.ymd( myrs.rs("add_data") )%></span>
</td></tr>
<%myrs.nexts:loop:myrs.close%>

</table>
<%else%>	
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1">
<tr><td align="center">暂无 <%=db_title%> 相应内容!</td></tr></table>	
<%end if%>

<%
'page
'maxPage
'allCount
dim page,pageCounts,allCounts,pageAdds
page       = myrs.page
pageCounts = myrs.pageCounts
allCounts  = myrs.allCounts
pageAdds = 3
%>
<!--#include file="article_paging.asp"-->

</td></TR></TABLE>

</body>
</html>