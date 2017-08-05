<!--#include file="email_config.asp"-->
<% 
set rs=server.createobject("adodb.recordset") 
    sql="select * from Email_default order by id desc" 
    rs.open sql,conn,1,1   
	if not rs.eof then
	   frommail   =rs("frommail")
	   mailsubject=rs("mailsubject")
	   mailbody   =rs("mailbody")
	end if
	rs.close
set rs=nothing

%>

<br />
<table border="0" align=center cellpadding=0 cellspacing=10 bgcolor="#FFFFFF">
  <tr> 
    <td width="600" align="center" valign="top" class="forumRow">
<!--#include file="email_top.asp"-->

	
    <form name="sendmail" action="EMAIL_SendTo.asp" method="post" >
      <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forMy">
          <tr> 
            <td colspan="2" align="center" class="forumRaw"><b>发 送 邮 件</b></td>
          </tr>
          <tr> 
            <td width="96" height="25" align="right" class="forumRow"> 
            发信人地址：</td>
            <td width="493" height="30" class="forumRow"> 
            <input type="text" name="frommail" value="<%=frommail%>">            </td>
          </tr>
          <tr> 
            <%email=request("email")%>
            <td align="right" class="forumRow"> 
            收信人地址：</td>
            <td class="forumRow"> 
            <input name="tomail" type="text" value="<%=email%>">            
            <font color="#000000">(如果为空，从数据库取地址群发) </font></td>
          </tr>
          <tr> 
            <td align="right" class="forumRow"> 
            信件标题：</td>
            <td class="forumRow"> 
              <input name="mailsubject" type="text" value="<%=mailsubject%>" size="30">
              <font color="#000000">(</font><font color="#000000">如果为空，则显示<font color="#FF0000">&quot;<%=rs("mailsubject")%>&quot;</font>) </font></td>
          </tr>
          <tr> 
            <td align="right" class="forumRow"> 
            信件内容：</td>
            <td class="forumRow"> 
              <textarea name="mailbody" cols="50" rows="8"><%=mailbody%></textarea>
              <br> </td>
          </tr>
          <tr> 
            <td height="25" class="forumRow"> 
            <div align="center">&nbsp;</div></td>
            <td height="22" class="forumRow"><input type="submit" name="Submit" value="确认发送"> &nbsp;&nbsp;&nbsp; 
            <input type="reset" name="Submit" value="清除重来"> <input type="hidden" name="action" value="我发">            </td>
          </tr>
        </table>
    </form></td>
  </tr>
</table>
