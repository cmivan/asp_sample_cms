<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop table">
<tr class="forumRow">
<td colspan="3" align="left">&nbsp;&nbsp;分类<span class="icon icon-circle-arrow-right">&nbsp;</span>&nbsp;&nbsp;&nbsp;| <a href="?"><strong>&nbsp;全部&nbsp;</strong></a>
<%=articles.typeItems(db_table,0,typeB_id," | {link}")%> </td></tr>
<%
dim typeSbox
  typeItems = articles.typeItems(db_table,typeB_id,typeS_id," - {link}")
  if typeItems<>"" then
%>
<tr class="forumRow"><td colspan="3" align="left" style="padding-left:50px;">
<%=typeItems%> </td></tr>
<%end if%>
</table>