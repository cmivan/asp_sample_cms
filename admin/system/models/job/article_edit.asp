<!--配置文件-->
<!--#include file="article_config.asp"-->
<%
'///////////  处理提交数据部分 ////////// 
   edit_id   =request("id")
if edit_id="" or isnumeric(edit_id)=false then
   editStr="添加"
   else
   editStr="修改"
end if
if request.Form("edit")="ok" then
'///////////  写入数据部分 //////////
   set rs=server.createobject("adodb.recordset") 
    if edit_id="" or isnumeric(edit_id)=false then
       exec="select * from "&db_table                       '判断，添加数据
	   rs.open exec,conn,1,3
       rs.addnew
	else
	   exec="select * from "&db_table&" where id="&edit_id  '判断，修改数据
       rs.open exec,conn,1,3
	end if
	if edit_id<>"" and isnumeric(edit_id) and rs.eof then
	   response.Write("写入数据失败!")
	else

	rs("zpdx")  =request.Form("zpdx")
	rs("zprs")  =request.Form("zprs")
	rs("gzdd")  =request.Form("gzdd")
	rs("gzdy")  =request.Form("gzdy")
	rs("yxqx")  =request.Form("yxqx")
	rs("fbrq")  =date()
	rs("zpyq")  =request.Form("zpyq")

    if sys_back then sys_backs="window.location.href='article_manage.asp';"
    response.Write("<script>alert('"&editStr&"操作成功!');"&sys_backs&"</script>")
	end if
	rs.update
	rs.close
set rs=nothing
end if

'#################  读取数据部分 #################
   id=request.QueryString("id") 
if id<>"" and isnumeric(id) then
   edit_stat="edit"
   set rs=server.createobject("adodb.recordset") 
       exec="select * from "&db_table&" where id="&id
       rs.open exec,conn,1,1 
	   if not rs.eof then
    	  zpdx =rs("zpdx")
    	  zprs =rs("zprs")
    	  gzdd =rs("gzdd")
    	  gzdy =rs("gzdy")
    	  yxqx =rs("yxqx")
    	  fbrq =rs("fbrq")
		  zpyq =rs("zpyq")
	   end if
	   rs.close
   set rs=nothing
else
   edit_stat="add"  '/// 当前状态:添加
end if


  '/// 处理添加时间
  if isdate(add_data)=false then add_data=now()
%>


<body>
<TABLE width="600" border="0" align="center" cellpadding="0" cellspacing="10" bgcolor="#FFFFFF">
  <TR><td class="forumRow">
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forMy forMAargin">
        <tr>
          <td align="center" class="forumRow"><strong>&nbsp;&nbsp;<%=db_zpdx%> <%=editStr%></strong></td>
        </tr>
      </table>

<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forMy">
        <form name="article_update" method="post" action="">
<tr> 
<td width="19%" height="25" align="right" valign="top" class="forumRow"> 
                      招聘对象</td>
                    <td width="81%" class="forumRow"> 
                      <input name="zpdx" type="text" class="inputtext" id="zpdx" value="<%=zpdx%>"  size="25" maxlength="30">                    </td>
                  </tr>
                  <tr> 
                    <td height="22" align="right" valign="top" class="forumRow"> 
                      招聘人数</td>
                    <td class="forumRow"> 
                      <input name="zprs" type="text" class="inputtext" id="zprs" value="<%=zprs%>"  size="5" maxlength="30">
                      人</td>
                  </tr>
                  <tr> 
                    <td height="22" align="right" valign="top" class="forumRow"> 
                      工作地点</td>
                    <td height="20" class="forumRow"> 
                      <input name="gzdd" type="text" class="inputtext" id="gzdd" value="<%=gzdd%>"  size="30" maxlength="50">                    </td>
                  </tr>
                  <tr> 
                    <td height="22" align="right" valign="top" class="forumRow"> 
                      工资待遇</td>
                    <td height="10" class="forumRow"> 
                      <input name="gzdy" type="text" class="inputtext" id="gzdy" value="<%=gzdy%>"  size="20" maxlength="50">                    </td>
                  </tr>

<tr> 
<td height="22" align="right" valign="top" class="forumRow"> 
发布日期</td>
<td height="3" class="forumRow"><%if fbrq<>"" then%><%=fbrq%><%else%><%=Date()%><%end if%></td>
</tr>


                  <tr> 
                    <td height="22" align="right" valign="top" class="forumRow"> 
                      有效期限</td>
                    <td height="0" class="forumRow"> 
                      <input name="yxqx" type="text" class="inputtext" id="yxqx" value="<%=yxqx%>"  size="5" maxlength="30">
                      天</td>
                  </tr>
                  <tr> 
                    <td height="22" align="right" valign="top" class="forumRow"> 
                      招聘要求</td>
                    <td height="10" class="forumRow"> 
                      <textarea name="zpyq" cols="40" rows="5" class="inputtext" id="zpyq"><%=zpyq%></textarea></td>
                  </tr>
		  
          <tr>
            <td align="right" valign="top" class="forumRow">
                <input name="id" type="hidden" value="<%=id%>" />
				<input name="edit" type="hidden" value="ok" />
            </td>
            <td class="forumRow">
              <input name="submit" type="submit" value="确认<%=editStr%>" />
            </td>
          </tr>
        </form>
      </table>
</td></TR></TABLE>
</body>
</html>