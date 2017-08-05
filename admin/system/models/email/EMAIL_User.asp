<!--#include file="email_config.asp"-->

<% if request("del")<>"" then 
   conn.Execute("delete * from email where id="&request("del"))
   end if %>
<SCRIPT LANGUAGE="JavaScript">
<!--//
function check()
{
   if (isNaN(go2to.page.value))
		alert("请正确填写转到页数！");
   else if (go2to.page.value=="") 
	     {
		alert("请输入转到页数！");
		 }
   else
		go2to.submit();
}
//-->
</SCRIPT>
<% 
     set rs=server.createobject("adodb.recordset") 
     if not isempty(request("page")) then   
     pagecount=cint(request("page"))   
     else   
     pagecount=1   
     end if

     if key="" then
     sql="select * from email order by id desc" 
     else
     sql="select * from email where dateandtime like '%"&key&"%' order by id desc" 
     end if

	 rs.open sql,conn,1,1
     rs.pagesize=15
     if pagecount>rs.pagecount or pagecount<=0 then              
     pagecount=1              
     end if              
     rs.AbsolutePage=pagecount %>
<br />

<table border="0" align=center cellpadding=0 cellspacing=10 class="forum1">
<tr><td>
<!--#include file="email_top.asp"-->

<%if rs.pagecount>1 then %>  
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="forum2">
<tr class="forumRow">
<td height="25" colspan="4"> <% response.write"<form name=go2to form method=Post action='?key="+key+"'>"
		     response.write "<font color=ffffff>◆&nbsp;</font>"                                              
		     if pagecount=1 then                                                        
			 response.write "<font class=ver>首页 上一页</font>&nbsp;"
			 else                                                        
             response.write "<a href=?page=1&key="+key+"><font class=ver>首页</font></a>&nbsp;" 
	         response.write "<a href=?page="+cstr(pagecount-1)+"&key="+key+"><font color='0000BE'>上一页</font></a>&nbsp;"
			 end if                                             
             if rs.PageCount-pagecount<1 then                                                        
             response.write "<font class=ver>下一页 尾页</font>"                                                    
			 else                                                        
             response.write "<a href=?page="+cstr(pagecount+1)+"&key="+key+"><font class=ver>下一页</font></a>&nbsp;"
			 response.write "<a href=?page="+cstr(rs.PageCount)+"&key="+key+"><font class=ver>尾页</font></a>"           
             end if 
			 response.write "<font class=ver>&nbsp;页次:<font class=ver>"&pagecount&"</font>/"&rs.pagecount&"页</font>" 
			response.write "<font class=ver> 转到第<input type='text' name='page' size=2 maxLength=3 style='font-size: 9pt; color:#00006A; position: relative; height: 18' value="&PageCount&">页</font>&nbsp;&nbsp;&nbsp;&nbsp;"                               
			response.write "<input class=button1 type='button' value='确 定' onclick=check() style='font-family: 宋体; font-size: 9pt; color: #000073; position: relative; height: 19'>" %> </td>
</tr></table>
<%end if%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="forum2">
<tr class="forumRaw">
<td width="9%" align="center" > 
         <font color="#000000">ID</font></td>
          <td align="center"> 
          <font color="#000000">邮件地址</font></td>
          <td colspan="3" align="center"> 
          <font color="#000000">用户名</font></td>
          <td  width="80" align="center"> 
         <font color="#000000">操作</font>
</td></tr>
  
<% do while not rs.eof %>

<tr class="forumRow">
          <td width="50" align="center"> 
          <font color=000000><%=rs("id")%></font></td>
          <td align="center" bgcolor="#ECF5FF"  class=forumRow> 
          <font color="000000"><a href="EMAIL_Send.asp?email=<%=rs("email")%>"><%=rs("email")%></a></font></td>
          <td colspan="3" align="center" bgcolor="#ECF5FF"  class=forumRow> 
          <font color="000000"><%=rs("dateandtime")%></font></td>
          <td align="center" bgcolor="#ECF5FF" class="forumRow" > 
          <font color="#000000">[<a href="?del=<%=rs("id")%>"><font class=ver>删除</font></a>]</font>          </td>
        </tr>
<% i=i+1                                                                                                  
          rs.movenext                                                                                                  
          if i>=rs.PageSize then exit do 
		  loop                                                                    
          rs.close                                                                                                
          set rs=nothing                                                                                                
          conn.close                                                                                                
          set conn=nothing 
'end if
%>

</table>
