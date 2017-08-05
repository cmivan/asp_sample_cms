<!--#include file="../../system/core/initialize_system.asp"-->
<script language="javascript">
$(function(){
	$(".menu").find("dl").find("dt").find("a").click(function(){
		$(this).parent().parent().parent().find("dd").css({"display":"none"});
		$(this).parent().parent().find("dd").css({"display":"block"});
	});
	//<><><><><><><><><><><><><><>
	$(".menu").find("dl").eq(0).find("dd").css({"display":"block"});
});
</script>
<style type="text/css">
*{margin:0px;}
body{
	background-color:#9E9C92;background-image:url(../../public/images/left_top_bg.gif);
	background-position:left top;background-repeat:repeat-y;overflow-x:hidden;
}
</style>

<base target="main">
<body>
<div class="menu">

<%
'多语言切换按钮
 if urlStr<>"" then
    urlStrs=split(urlStr,"|")
	response.Write("<table width=""100%"" border=""0""><tr>")
    for each lannav in urlStrs
        nav=split(lannav,",")
        if ubound(nav)=1 then
		   response.Write("<td align=""center"" class=""lannav"">")
		   response.Write("<a href="""&GetUrl(nav(1))&""" "&OnUrl(nav(1))&" target=""_top"" >"&nav(0)&"</a>")
		   response.Write("</td>")
        end if
    next
	response.Write("</tr></table>")
end if

set rs = sysconn.execute("select * from sys_nav where type_id=0 and show=1 order by order_id asc,id desc")
 if rs.eof then
	response.Write "&nbsp;暂无"&db_title&"分导航"
 else
	do while not rs.eof
%>
<!-- Item Strat -->
<dl>
<%
if rs("url")<>"" then 
   thisUrl=urlPath&rs("url")
else
   thisUrl="javascript:void(0);"
end if
%>
<dt><a href="<%=thisUrl%>"><%=rs("title")%></a></dt>
<dd><ul>
<%
dim db_showSet
    db_showSet=rs("db_showSet")
if db_showSet<>"" and ubound(split(db_showSet,"|"))=2 then
'########## 情况3, 符合数组
   arr_nav(2)=""
   arr_nav=split(db_showSet,"|")
if arr_nav(0)<>"" and arr_nav(1)<>"" and arr_nav(2)<>"" and isnumeric(arr_nav(2)) then
    select case arr_nav(2)
		   case 2    '只有管理、添加菜单
%>
<li><a target="main" href="<%=urlPath&arr_nav(1)%>/article_manage.asp"><%=arr_nav(0)%>管理</a></li>
<li><a target="main" href="<%=urlPath&arr_nav(1)%>/article_edit.asp"><%=arr_nav(0)%>添加</a></li>
<%		  
		   case 3    '管理、添加、分类菜单具备
%>
<li><a target="main" href="<%=urlPath&arr_nav(1)%>/article_manage.asp"><%=arr_nav(0)%>管理</a></li>
<li><a target="main" href="<%=urlPath&arr_nav(1)%>/article_edit.asp"><%=arr_nav(0)%>添加</a></li>
<li><a target="main" href="<%=urlPath&arr_nav(1)%>/article_type.asp"><%=arr_nav(0)%>分类</a></li>
<%	
		   case else '只有管理菜单
%>
<li><a target="main" href="<%=urlPath&arr_nav(1)%>/article_manage.asp"><%=arr_nav(0)%>管理</a></li>
<%	   
    end select
end if
end if

set js = sysconn.execute("select * from sys_nav where type_id="&rs("id")&" and show=1 order by order_id asc,id desc")
	do while not js.eof
%>
<%
if js("title")="-" then
'########## 情况1, "-" 为 分行符
%>
<div class="line">&nbsp;</div>
<%
elseif left(js("title"),1)="#" then
'########## 情况2, "#company" 其中"#"为标识符 "company" 为表名

Tstr=js("title")
Tstr=replace(Tstr,"#","")
if Tstr<>"" and CheckTable(Tstr) then
    set ts = conn.execute("select * from "&Tstr&" order by order_id asc,id desc")
		do while not ts.eof
%>
<li><a target="main" href="<%=urlPath&Tstr%>/article_edit.asp?id=<%=ts("id")%>"><%=ts("title")%></a></li>
<%
		ts.movenext
		loop
		ts.close
	set ts=nothing
end if
%>
<%else%>
<li><a target="main" href="<%="../"&js("url")%>"><%=js("title")%></a></li>
<%end if%>
<%
		js.movenext
		loop
		js.close
	set js=nothing
%>
</ul></dd></dl>
<%		
	rs.movenext
	loop
 end if
    rs.close
set rs=nothing
%>


<%
if session("super")=1 then%>
<!-- Item Strat -->
<dl><dt><a href="javascript:void(0);">系统管理</a></dt><dd><ul>
<%
set js = sysconn.execute("select * from sys_admin_nav where IsSys=1 order by order_id asc,id desc")
	do while not js.eof
%>
<li><a target="main" href="<%="../"&js("url")%>"><%=js("title")%></a></li>
<%
	js.movenext
	loop
	js.close
set js=nothing
%>
</ul></dd></dl>
<%end if%>


<%
'<><><>是否显示在后台导航版权信息
if IsShow then%>
<dl class="copyright">
<dt><a>版权信息</a></dt><dd><ul><li>
<a href="<%=Official%>" target="_blank" style=" color:#333333;line-height:20px;padding:3px; padding-left:12px; margin:0; text-indent:0; background:none; border-left:#CCCCCC 1px solid; border-top:#CCCCCC 1px solid; width:136px;">开发：<%=Author%><br>联系：<%=Contact%><br>官网：<%=Sysname%><br></a>
</li></ul></dd></dl>
<%end if%>

<dl><dt><a href="?login=out" target="_top">退出管理</a></dt></dl>
<!-- Item End -->
</div>
</body>
