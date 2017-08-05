<form name="search" method="get" class="form-search forumtop" action="">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
<tr><TD align="center" class="manage-title"><span class="icon icon-folder-open">&nbsp;</span> <strong><%=db_title%></strong> 管理列表</td>
<TD class="manage-search">
  <input name="keywords" type="text" id="keywords" value="<%=keywords%>"/>
  <button type="submit" class="btn"><span class="icon icon-search">&nbsp;</span> 搜索</button>
  <input name="typeB_id" type="hidden" id="typeB_id" value="<%=typeB_id%>" />
  <input name="typeS_id" type="hidden" id="typeS_id" value="<%=typeS_id%>" />
</td></TR></table>
</form>