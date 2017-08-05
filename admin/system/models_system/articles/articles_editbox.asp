<%
'###### 用户使用的编辑框 #######
    editbox=int(session("user_editbox"))
 if editbox="" or isnumeric(editbox)=false then editbox=0

'###### eWebEditor编辑框
if editbox=0 then
%>
<textarea name="content" style="display:none"><%=content%></textarea>
<IFRAME ID="eWebEditor1" SRC="<%=Rpath%>public/editbox/ewebeditor/ewebeditor.asp?id=content&style=s_light" FRAMEBORDER="0" SCROLLING="no" WIDTH="100%" HEIGHT="400"></IFRAME>

<%
'###### FCKeditor编辑框
elseif editbox=1 then

content=replace(content,"'","&#39;")

	Dim oFCKeditor 
	Set oFCKeditor = New FCKeditor 
	oFCKeditor.BasePath = Rpath&"public/editbox/fckeditor/"
	oFCKeditor.ToolbarSet = "Default" 
	oFCKeditor.Width = "100%" 
	oFCKeditor.Height = "400" 
	oFCKeditor.Value = content
	oFCKeditor.Create "content"

'###### eWebEditor编辑框 支持word
elseif editbox=2 then
%>
<textarea name="content" id="content" style="display:none"><%=content%></textarea>
<iframe id="content_html" src="<%=Rpath%>public/editbox/Ewebeditor_Word/ewebeditor.htm?id=content&style=blue" FRAMEBORDER="0" SCROLLING="no" WIDTH="100%" HEIGHT="400"></iframe>
<%
'###### 163编辑框
elseif editbox=3 then
%>
<textarea name="content" id="content" style="display:none"><%=content%></textarea>
<iframe id="Editor" name="content" src="<%=Rpath%>public/editbox/163Editor/editor.html?ID=content&style=blue" frameborder="0" marginheight="0" marginwidth="0" scrolling="No" style="height:320px;width:100%"></iframe>

<%end if%>