<!--配置文件-->
<!--#include file="article_config.asp"-->
<%
'<><><><><>处理提交数据部分<><><><><>
dim edit_id
   edit_id=request("id")
if isnum(edit_id) then
   editStr=sys_str_edit
else
   editStr=sys_str_add
end if
   
if request.Form("edit")="ok" then
  tf_title=request.form("tf_title")tf_content=request.form("tf_content")tf_note=request.form("tf_note")'——————分类
type_id=request.form("type_id")
'//系判断是否为数组，再判断是否为数字(否则失败)
type_ids=split(type_id,",")
if ubound(type_ids)=1 then
   if isnum(type_ids(0)) and isnum(type_ids(1)) then
	  typeB_id=type_ids(0)
	  typeS_id=type_ids(1)
	  session("type_id")=type_id   '记录本次操作分类
   else
	  call errtips(sys_tip_errtype)
   end if
else
   if isnum(type_id) then
	  typeB_id=type_id
	  typeS_id=null
	  session("type_id")=type_id   '记录本次操作分类
   else
	  call errtips(sys_tip_errtype)
   end if
end iftf_typeS_id=request.form("tf_typeS_id")tf_big_pic=request.form("tf_big_pic")tf_small_pic=request.form("tf_small_pic")tf_add_data=request.form("tf_add_data")tf_order_id = request.form("tf_order_id")tf_hits=request.form("tf_hits")tf_ok=request.form("tf_ok")tf_hot=request.form("tf_hot")tf_news=request.form("tf_news")tf_tj=request.form("tf_tj")tf_color=request.form("tf_color")tf_strong=request.form("tf_strong")tf_size=request.form("tf_size")tf_toUrl=request.form("tf_toUrl")
  

'///////////  写入数据部分 //////////
set rs=server.createobject("adodb.recordset") 
    if isnum(edit_id) then
	   exec="select * from "&db_table&" where id="&edit_id  '判断，修改数据
       rs.open exec,conn,1,3
    else
       exec="select * from "&db_table                       '判断，添加数据
	   rs.open exec,conn,1,3
       rs.addnew
    end if
    if isnum(edit_id) and rs.eof then
	   response.Write(sys_tip_none)
    else
       articles.rs("title")=tf_titlearticles.rs("content")=tf_contentarticles.rs("note")=tf_notearticles.rs("typeB_id") =typeB_id
articles.rs("typeS_id") =typeS_idarticles.rs("typeS_id")=tf_typeS_idarticles.rs("big_pic")=tf_big_picarticles.rs("small_pic")=tf_small_picarticles.rs("add_data")=tf_add_dataif txt.isNum(tf_order_id) then articles.rs("order_id")=tf_order_idarticles.rs("hits")=tf_hitsarticles.rs("ok")=tf_okarticles.rs("hot")=tf_hotarticles.rs("news")=tf_newsarticles.rs("tj")=tf_tjarticles.rs("color")=tf_colorarticles.rs("strong")=tf_strongarticles.rs("size")=tf_sizearticles.rs("toUrl")=tf_toUrl
	   
	   '<><>判断是否正确写入<><>
	   if err<>0 then call backPage(editStr&sys_tip_false,"article_edit.asp?id="&edit_id,0)
	end if
	rs.update
	rs.close
set rs=nothing
    call backPage(editStr&sys_tip_ok,"article_manage.asp?typdB_id="&typdB_id&"&typdS_id="&typdS_id,0)
	
end if

'<><><><><><>读取数据部分<><><><><><>
id=request.QueryString("id") 
if isnum(id) then
   set rs=conn.execute("select * from "&db_table&" where id="&id) 
	   if not rs.eof then
tf_title=rs("title")tf_content=rs("content")tf_note=rs("note")typeB_id  =rs("typeB_id")
typeS_id  =rs("typeS_id")
'记录分类,用于分类下拉
if isnum(typeS_id) then
   session("type_id")=typeB_id
else
   session("type_id")=typeB_id&","&typeS_id
end iftf_typeS_id=rs("typeS_id")tf_big_pic=rs("big_pic")tf_small_pic=rs("small_pic")tf_add_data=rs("add_data")tf_order_id=rs("order_id")
if txt.isNum(tf_order_id)=false then tf_order_id=0tf_hits=rs("hits")tf_ok=rs("ok")tf_hot=rs("hot")tf_news=rs("news")tf_tj=rs("tj")tf_color=rs("color")tf_strong=rs("strong")tf_size=rs("size")tf_toUrl=rs("toUrl")
	   end if
	   rs.close
   set rs=nothing  
end if
%>


<script>
function CheckForm()
{
 
}
///返回提示框
function Tips(ThisF){
  ThisF.style.background="#FF6600";
}
</script>
<body>
<TABLE border="0" align="center" cellpadding="0" cellspacing="10" class="forum1">
<TR><td>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="forum2 forumtop">
<tr class="forumRaw forumtitle">
<td align="center"><strong>&nbsp;&nbsp;<%=db_title%> <%=editStr%></strong></td>
</tr></table>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="forum2">
<form name="FormE" method="post" action="" onSubmit="return CheckForm(this);">
<tr class="forumRow"><td align="right">title：</td><td><input type="text" name="tf_title" id="tf_title" value="<%=tf_title%>" />
</td></tr><tr class="forumRow"><td align="right">content：</td><td><textarea name="tf_content" style="display:none"><%=tf_content%></textarea>
<IFRAME ID="eWebEditor1" SRC="../../public/editBox/ewebeditor/ewebeditor.asp?id=tf_content&style=s_light" FRAMEBORDER="0" SCROLLING="no" WIDTH="100%" HEIGHT="400"></IFRAME>



</td></tr><tr class='forumRow'><td align='right'>note：</td><td><textarea name="tf_note" style="width:100%" rows="4"><%=tf_note%></textarea>



</td></tr><tr class="forumRow"><td align="right">typeB_id：</td><td><!--#include file="../../public/articles_type.asp"-->







</td></tr><tr class='forumRow'><td align='right'>typeS_id：</td><td><textarea name="tf_typeS_id" style="width:100%" rows="4"><%=tf_typeS_id%></textarea>



</td></tr><tr class="forumRow"><td align="right">big_pic：</td><td><input name="tf_big_pic" id="tf_big_pic" type="text" value="<%=tf_big_pic%>" size="45" maxlength="255" />
<input type=button value="浏览图片" onClick="showUploadDialog('image', 'FormE.tf_big_pic', '')" class="button"/>



</td></tr><tr class="forumRow"><td align="right">small_pic：</td><td><input name="tf_small_pic" id="tf_small_pic" type="text" value="<%=tf_small_pic%>" size="45" maxlength="255" />
<input type=button value="浏览图片" onClick="showUploadDialog('image', 'FormE.tf_small_pic', '')" class="button"/>



</td></tr><tr class='forumRow'><td align='right'>add_data：</td><td><textarea name="tf_add_data" style="width:100%" rows="4"><%=tf_add_data%></textarea>



</td></tr><tr class="forumRow"><td align="right">order_id：</td><td><input type="text" name="tf_order_id" id="tf_order_id" value="<%=tf_order_id%>" />

</td></tr><tr class='forumRow'><td align='right'>hits：</td><td><textarea name="tf_hits" style="width:100%" rows="4"><%=tf_hits%></textarea>



</td></tr><tr class='forumRow'><td align='right'>ok：</td><td><textarea name="tf_ok" style="width:100%" rows="4"><%=tf_ok%></textarea>



</td></tr><tr class='forumRow'><td align='right'>hot：</td><td><textarea name="tf_hot" style="width:100%" rows="4"><%=tf_hot%></textarea>



</td></tr><tr class='forumRow'><td align='right'>news：</td><td><textarea name="tf_news" style="width:100%" rows="4"><%=tf_news%></textarea>



</td></tr><tr class='forumRow'><td align='right'>tj：</td><td><textarea name="tf_tj" style="width:100%" rows="4"><%=tf_tj%></textarea>



</td></tr><tr class='forumRow'><td align='right'>color：</td><td><textarea name="tf_color" style="width:100%" rows="4"><%=tf_color%></textarea>



</td></tr><tr class='forumRow'><td align='right'>strong：</td><td><textarea name="tf_strong" style="width:100%" rows="4"><%=tf_strong%></textarea>



</td></tr><tr class='forumRow'><td align='right'>size：</td><td><textarea name="tf_size" style="width:100%" rows="4"><%=tf_size%></textarea>



</td></tr><tr class='forumRow'><td align='right'>toUrl：</td><td><textarea name="tf_toUrl" style="width:100%" rows="4"><%=tf_toUrl%></textarea>



</td></tr>
<tr class="forumRaw">
<td width="88" align="right" valign="top">
<input name="id" type="hidden" value="<%=id%>" />
<input name="edit" type="hidden" value="ok" />
</td>
<td><input name="submit" type="submit" value="确认<%=editStr%>" class="button" /></td>
</tr>
</form>
</table></td>
</tr></table>
</body>
</html>