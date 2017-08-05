<!--#include file="../../system/core/initialize_system.asp"-->
<!--#include file="../libraries/articlesHtml.asp"-->
<%
'### 配置信息 ###
dim db_title
 db_title = "网站信息配置"

'### 修改配置信息 ######
if request.form("formAction")="save" then
set config=server.createobject("adodb.recordset")
    exec="select * from sys_config"
    config.open exec,sysconn,1,3
    config("title")=request.form("title")
    config("url")=request.form("url") 
    config("contact")=request.form("contact")
    config("tel")=request.form("tel")
    config("fax")=request.form("fax")
    config("mobile")=request.form("mobile")
    config("email")=request.form("email")
    config("qq")=request.form("qq")
    config("msn")=request.form("msn")
    config("address")=request.form("address")
    config("keywords")=request.form("keywords") 
    config("description")=request.form("description") 
    config("copyright")=request.form("copyright") 
    config.update 
    config.close 
	if err=0 then call alert.reFresh("网站基本信息设置成功!","?")
end if
%>


<body>
<table border="0" align="center" cellpadding="0" cellspacing="10" class="forum1">
<form id="config" name="config" method="post" action="">
<tr><td>

<!--#include file="../models_system/articles/articles_edit_head.asp"-->   
     
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop table">
<tr class="forumRow">
         <td width="30%" align="right">网站名称：</td>
         <td width="70%"><input name="title" type="text" value="<%=c_title%>" /></td>
       </tr>
       <tr class="forumRow">
         <td align="right">地址：</td>
         <td><input name="address" type="text" value="<%=c_address%>" size="40" /></td>
       </tr>
       <tr class="forumRow">
         <td align="right">联系人：</td>
         <td><input name="contact" type="text" value="<%=c_contact%>" /></td>
       </tr>
       <tr class="forumRow">
         <td align="right">传真：</td>
         <td><input name="fax" type="text" value="<%=c_fax%>" /></td>
       </tr>
       <tr class="forumRow">
         <td align="right">网址：</td>
         <td><input name="url" type="text" value="<%=c_url%>" /></td>
       </tr>
       <tr class="forumRow">
         <td align="right">邮箱：</td>
         <td><input name="email" type="text" value="<%=c_email%>" /></td>
       </tr>
       <tr class="forumRow" style="display:none">
         <td align="right">电话：</td>
         <td><input name="tel" type="text" value="<%=c_tel%>" /></td>
       </tr>

       <tr class="forumRow" style="display:none">
         <td align="right">手机：</td>
         <td><input name="mobile" type="text" value="<%=c_mobile%>" /></td>
       </tr>

       <tr class="forumRow" style="display:none">
         <td align="right">QQ：</td>
         <td><input name="qq" type="text" value="<%=c_qq%>" /></td>
       </tr>
       <tr class="forumRow" style="display:none">
         <td align="right">MSN：</td>
         <td><input name="msn" type="text" value="<%=c_msn%>" size="40" /></td>
       </tr>

       <tr class="forumRow" style="display:none">
         <td align="right">网站关键字：</td>
         <td><input name="keywords" type="text" value="<%=c_keywords%>" size="40"/></td>
       </tr>
       <tr class="forumRow" style="display:none">
         <td align="right">关键字描述：</td>
         <td><input name="description" type="text" value="<%=c_description%>" size="40" /></td>
       </tr>
       <tr class="forumRow">
         <td align="right">版权信息：</td>
         <td><%=articlesHtml.noteBox("copyright",c_copyright)%></td>
       </tr>
       <tr class="forumRaw">
         <td>&nbsp;<input type="hidden" name="formAction" value="save"/></td>
         <td><%=articlesHtml.submitBtn("","修改")%></td>
       </tr>
     </table></td>
   </tr>
   </form>
</table>
</body>
</html>