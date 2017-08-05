<!--#include file="../../system/core/initialize_system.asp"-->
<style>
*{margin:0px;}
.TopNav {
padding-right:40px;}
.TopNav a{
float:right;
padding:15px;
padding-top:5px;
padding-bottom:5px;
margin-right:2px;
text-align:center;
text-decoration:none;
}
.TopNav a:hover{
text-decoration:none;
background-color:#000000;
margin-bottom:0px;
}
</style>
<body>
<%if request.QueryString("page")="" then%>
<div class="topnav">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr><td width="170" style="padding:15px;padding-right:5px;padding-left:10px;"><img src="../../public/images/logo.png" width="174" height="50" align="absmiddle"></td>
<td width="200" valign="bottom" style="padding:15px;padding-left:0px;font-size:13px;color:#000000;font-weight:bold;">企业管理系统 V<%=versions%> </td>
<td valign="bottom" class="TopNav">&nbsp;</td></tr></table>
</div>
<%elseif request.QueryString("page")="top" then%>
<div class="topline"><img src="../../public/images/left_top.gif"></div>
<%elseif request.QueryString("page")="bottom" then%>
<div class="bottomline"><img src="../../public/images/left_bottom.gif"></div>
<%end if%>
</body>