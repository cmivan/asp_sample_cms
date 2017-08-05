<!--#include file="system/core/initialize.asp"-->
<!--#include file="system/libraries/articles.asp"-->
<!--#include file="system/helpers/md5_helpers.asp"-->
<!--#include file="system/libraries/reSizeImg.asp"-->
<%
'// 获取详细信息
dim id,pageTitle,pageContent
    id = input.getnum("id")
	pageTitle = "{title}"
	pageContent = "{content}"

articles.table = "product"
set rs = articles.pageView(id)
    if rs<>false then
	   pageTitle = rs("title")
	   pageContent = rs("content")
	end if
set rs = nothing
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

<strong><%=pageTitle%></strong>
<br />

<%ListHB = text.getImgSrc( pageContent )%>
<%=reSizeImg.imgSplice(ListHB,600,5,"")%>

<hr />

<%=pageContent%>

</td></TR></TABLE>

</body>
</html>