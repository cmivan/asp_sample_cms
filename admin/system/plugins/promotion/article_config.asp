<!--#include file="../../../system/core/initialize_system.asp"-->
<!--#include file="../../libraries/articlesHtml.asp"-->
<%
'### 配置信息 ###
 db_table  = "sys_promotion"   '表段
 db_title  = "推广"   '标题
 db_types  = 0        '0为无分类\1为只有大类\2为为只\3为不能删除添加
 db_showOK = 0        '0为不显示审核等信息，1则显示
%>