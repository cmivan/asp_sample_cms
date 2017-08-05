<!--#include file="../../system/core/initialize_system.asp"-->
<%
'设置密码
dim Report,RePath
%>
<style>td{padding-left:0px;}</style>
<body>
<table border="0" align="center" cellpadding="0" cellspacing="10" class="forum1">
<tr><td valign="top">

<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop table">
<tr class="forumRaw"><td class="mainTitle"><span class="ico icon-list">&nbsp;</span>&nbsp;ASP木马查找和可疑文件搜索功能</td></tr>
</table>

<%
  if session("cmcms.admin")<>"" then
  if request.QueryString("act")<>"scan" then
%>

<form action="?act=scan" method="post" name="form1">
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop table">
<tr class="forumRow">
<td width="15%" align="right">检查的路径：</td>
<td width="85%">
<input name="path" type="text" value="." size="20" />
<a href="javascript:void(0);" onClick="form1.path.value='*';">根目录相对路径“*”</a>;
<a href="javascript:void(0);" onClick="javascript:form1.path.value='\\'">整网站路径“\”</a>;
<a href="javascript:void(0);" onClick="javascript:form1.path.value='.'">程序所在目录“.”</a>
</td></tr>

<tr class="forumRow">
<td align="right">你要干什么：</td>
<td>
<label><input type="radio" name="radiobutton" value="sws" checked> 查ASP木马</label>
<label><input type="radio" name="radiobutton" value="sf"> 搜索符合条件之文件</label>
</td></tr>

<tr class="forumRow">
<td width="110" align="right" class="forumRow">查找内容：</td>
<td align="left" class="forumRow">
<input name="Search_Content" type="text" id="Search_Content" size="20">
* 要查找的字符串，不填就只进行日期检查</td></tr>

<tr class="forumRow">
<td align="right" class="forumRow">修改日期：</td>
<td align="left" class="forumRow">
<input name="Search_Date" type="text" value="<%=Left(Now(),InStr(now()," ")-1)%>" size="20">
* 多个日期用;隔开，任意日期填写
<a href="javascript:void(0);" onClick="javascript:form1.Search_Date.value='ALL'">ALL</a>
</td></tr>

<tr class="forumRow">
<td align="right">文件类型：</td>
<td align="left">
<input name="Search_FileExt" type="text" value="*" size="20">
* 类型之间用,隔开，*表示所有类型 </td></tr>

<tr class="forumRaw">
<td align="left">&nbsp;</td>
<td align="left">
<button type="submit" name="submit" class="btn"><span class="ico icon-search">&nbsp;</span>&nbsp;开始扫描</button>
</td></tr>

</table>
</form>

<%
	else
		server.ScriptTimeout = 600
		if request.Form("path")="" then
			response.Write("No Hack")
			response.End()
		end if
		if request.Form("path")="\" then
			TmpPath = Server.MapPath("\")
		elseif request.Form("path")="." then
			TmpPath = Server.MapPath(".")
		else
			TmpPath = Server.MapPath("\")&"\"&request.Form("path")
		end if
		timer1 = timer
		Sun = 0
		SumFiles = 0
		SumFolders = 1
		If request.Form("radiobutton") = "sws" Then
			DimFileExt = "asp,cer,asa,cdx"
			Call ShowAllFile(TmpPath)
		Else
			If request.Form("path") = "" or request.Form("Search_Date") = "" or request.Form("Search_FileExt") = "" Then
				response.Write("缉捕条件不完全，恕难从命<br><br><a href='javascript:history.go(-1);'>请返回重新输入</a>")
				response.End()
			End If
			DimFileExt = request.Form("Search_fileExt")
			Call ShowAllFile2(TmpPath)
		End If
%>

<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop">
<tr class="forumRaw"><td colspan="7">扫描完毕！一共检查文件夹<font color="#FF0000"><%=SumFolders%></font>个，文件<font color="#FF0000"><%=SumFiles%></font>个，发现可疑点<font color="#FF0000"><%=Sun%></font>个</td>
		      </tr>

<tr class="forumRaw">
<%If request.Form("radiobutton") = "sws" Then%>
<td width="18%">文件相对路径</td>
<td width="22%">特征码</td>
<td width="44%">描述</td>
<td width="16%">创建/修改时间</td>
<%else%>   
<td width="50%">文件相对路径</td>
<td width="25%">文件创建时间</td>
<td width="25%">修改时间</td>
<%end if%>
</tr>
<%=Report%>
</table>
			
<%
timer2 = timer
thetime=cstr(int(((timer2-timer1)*10000 )+0.5)/10)
%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop">
<tr class="forumRaw">
<td align="center">本页执行共用了 <%=thetime%> 毫秒</td>
</tr></table>
<%
	end if
end if
%>
</td></tr>
</table>
</body>
<%

'遍历处理path及其子目录所有文件
Sub ShowAllFile(Path)
	Set FSO = CreateObject("Scripting.FileSystemObject")
	if not fso.FolderExists(path) then exit sub
	Set f = FSO.GetFolder(Path)
	Set fc2 = f.files
	For Each myfile in fc2
		If CheckExt(FSO.GetExtensionName(path&"\"&myfile.name)) Then
			Call ScanFile(Path&Temp&"\"&myfile.name, "")
			SumFiles = SumFiles + 1
		End If
	Next
	Set fc = f.SubFolders
	For Each f1 in fc
		ShowAllFile path&"\"&f1.name
		SumFolders = SumFolders + 1
    Next
	Set FSO = Nothing
End Sub


Function getPath(path)
	dim ArrPaths,thisPath,TempPath
	    TempPath=path
	    ArrPaths=split(TempPath,"\")
	 if ubound(ArrPaths)>0 then getPath=ArrPaths(ubound(ArrPaths))
End Function

'检测文件
Sub ScanFile(FilePath, InFile)
	If InFile <> "" Then
		Infiles = "<font color=red>该文件被<a href=""http://"&Request.Servervariables("server_name")&"/"&tURLEncode(InFile)&""" target=_blank>"& InFile & "</a>文件包含执行</font>"
	End If
	Set FSOs = CreateObject("Scripting.FileSystemObject")
	on error resume next
	set ofile = fsos.OpenTextFile(FilePath)
	filetxt = Lcase(ofile.readall())
	If err Then Exit Sub end if
	if len(filetxt)>0 then
		'特征码检查
		'T_FilePath=FilePath
		'T_FilePath=replace(T_FilePath,"\","/")

'<><><><><><><><><><><><><><><><><><><><><><><><>
	dim ArrPaths,thisPath,TempPath
	    TempPath=FilePath
	    ArrPaths=split(TempPath,"\")
	 if ubound(ArrPaths)>0 then
		for i=0 to ubound(ArrPaths)-1
		    thisPath=thisPath&ArrPaths(i)&"\"
		next
	 end if
	if RePath<>thisPath then
	   RePath=thisPath
	   Report = Report&"<tr class=""forumRaw""><td colspan='4'>"&thisPath&"</td></tr>"
	end if
'<><><><><><><><><><><><><><><><><><><><><><><><>


	
		filetxt = vbcrlf & filetxt
		temp = "<a href="""&tURLEncode(replace(replace(FilePath,server.MapPath("./")&"\","",1,1,1),"\","/"))&""" target=_blank>"&getPath(FilePath)&"</a>"
			'Check "WScr"&DoMyBest&"ipt.Shell"
			If instr( filetxt, Lcase("WScr"&DoMyBest&"ipt.Shell") ) or Instr( filetxt, Lcase("clsid:72C24DD5-D70A"&DoMyBest&"-438B-8A42-98424B88AFB8") ) then
				Report = Report&"<tr class=""forumRow""><td>"&temp&"</td><td>WScr"&DoMyBest&"ipt.Shell 或者 clsid:72C24DD5-D70A"&DoMyBest&"-438B-8A42-98424B88AFB8</td><td><font color=red>危险组件，一般被ASP木马利用</font>"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
			End if
			'Check "She"&DoMyBest&"ll.Application"
			If instr( filetxt, Lcase("She"&DoMyBest&"ll.Application") ) or Instr( filetxt, Lcase("clsid:13709620-C27"&DoMyBest&"9-11CE-A49E-444553540000") ) then
				Report = Report&"<tr class=""forumRow""><td>"&temp&"</td><td>She"&DoMyBest&"ll.Application 或者 clsid:13709620-C27"&DoMyBest&"9-11CE-A49E-444553540000</td><td><font color=red>危险组件，一般被ASP木马利用</font>"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
			End If
			'Check .Encode
			Set regEx = New RegExp
			regEx.IgnoreCase = True
			regEx.Global = True
			regEx.Pattern = "\bLANGUAGE\s*=\s*[""]?\s*(vbscript|jscript|javascript).encode\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr class=""forumRow""><td>"&temp&"</td><td>(vbscript|jscript|javascript).Encode</td><td><font color=red>似乎脚本被加密了</font>"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
			End If
			'Check my ASP backdoor :(
			regEx.Pattern = "\bEv"&"al\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr class=""forumRow""><td>"&temp&"</td><td>Ev"&"al</td><td>e"&"val()函数可以执行任意ASP代码，被一些后门利用。其形式一般是：ev"&"al(X)<br>但是javascript代码中也可以使用，有可能是误报。"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
			End If
			'Check exe&cute backdoor
			regEx.Pattern = "[^.]\bExe"&"cute\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr class=""forumRow""><td>"&temp&"</td><td>Exec"&"ute</td><td><font color=red>e"&"xecute()函数可以执行任意ASP代码，被一些后门利用。其形式一般是：ex"&"ecute(X)</font><br>"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
			End If
			'----------------------Start Update 200605031-----------------------------
			'Check .Create&TextFile and .OpenText&File
			regEx.Pattern = "\.(Open|Create)TextFile\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr class=""forumRow""><td>"&temp&"</td><td>.CreateTextFile|.OpenTextFile</td><td>使用了FSO的CreateTextFile|OpenTextFile函数读写文件"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
			End If
			'Check .SaveT&oFile
			regEx.Pattern = "\.SaveToFile\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr class=""forumRow""><td>"&temp&"</td><td>.SaveToFile</td><td>使用了Stream的SaveToFile函数写文件"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
			End If
			'Check .&Save
			regEx.Pattern = "\.Save\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr class=""forumRow""><td>"&temp&"</td><td>.Save</td><td>使用了XMLHTTP的Save函数写文件"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
			End If
			'------------------              End           ----------------------------
			Set regEx = Nothing
			
		'Check include file
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "<!--\s*#include\s*file\s*=\s*"".*"""
		Set Matches = regEx.Execute(filetxt)
		For Each Match in Matches
			tFile = Replace(Mid(Match.Value, Instr(Match.Value, """") + 1, Len(Match.Value) - Instr(Match.Value, """") - 1),"/","\")
			If Not CheckExt(FSOs.GetExtensionName(tFile)) Then
				Call ScanFile( Mid(FilePath,1,InStrRev(FilePath,"\"))&tFile, replace(FilePath,server.MapPath("\")&"\","",1,1,1) )
				SumFiles = SumFiles + 1
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing
		
		'Check include virtual
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "<!--\s*#include\s*virtual\s*=\s*"".*"""
		Set Matches = regEx.Execute(filetxt)
		For Each Match in Matches
			tFile = Replace(Mid(Match.Value, Instr(Match.Value, """") + 1, Len(Match.Value) - Instr(Match.Value, """") - 1),"/","\")
			If Not CheckExt(FSOs.GetExtensionName(tFile)) Then
				Call ScanFile( Server.MapPath("\")&"\"&tFile, replace(FilePath,server.MapPath("\")&"\","",1,1,1) )
				SumFiles = SumFiles + 1
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing
		
		'Check Server&.Execute|Transfer
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "Server.(Exec"&"ute|Transfer)([ \t]*|\()"".*"""
		Set Matches = regEx.Execute(filetxt)
		For Each Match in Matches
			tFile = Replace(Mid(Match.Value, Instr(Match.Value, """") + 1, Len(Match.Value) - Instr(Match.Value, """") - 1),"/","\")
			If Not CheckExt(FSOs.GetExtensionName(tFile)) Then
				Call ScanFile( Mid(FilePath,1,InStrRev(FilePath,"\"))&tFile, replace(FilePath,server.MapPath("\")&"\","",1,1,1) )
				SumFiles = SumFiles + 1
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing
		
		'Check Server&.Execute|Transfer
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "Server.(Exec"&"ute|Transfer)([ \t]*|\()[^""]\)"
		If regEx.Test(filetxt) Then
			Report = Report&"<tr class=""forumRow""><td>"&temp&"</td><td>Server.Exec"&"ute</td><td><font color=red>不能跟踪检查Server.e"&"xecute()函数执行的文件。请管理员自行检查</font><br>"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
			Sun = Sun + 1
		End If
		Set Matches = Nothing
		Set regEx = Nothing

		'Check RunatScript
		Set XregEx = New RegExp
		XregEx.IgnoreCase = True
		XregEx.Global = True
		XregEx.Pattern = "<scr"&"ipt\s*(.|\n)*?runat\s*=\s*""?server""?(.|\n)*?>"
		Set XMatches = XregEx.Execute(filetxt)
		For Each Match in XMatches
			tmpLake2 = Mid(Match.Value, 1, InStr(Match.Value, ">"))
			srcSeek = InStr(1, tmpLake2, "src", 1)
			If srcSeek > 0 Then
				srcSeek2 = instr(srcSeek, tmpLake2, "=")
				For i = 1 To 50
					tmp = Mid(tmpLake2, srcSeek2 + i, 1)
					If tmp <> " " and tmp <> chr(9) and tmp <> vbCrLf Then
						Exit For
					End If
				Next
				If tmp = """" Then
					tmpName = Mid(tmpLake2, srcSeek2 + i + 1, Instr(srcSeek2 + i + 1, tmpLake2, """") - srcSeek2 - i - 1)
				Else
					If InStr(srcSeek2 + i + 1, tmpLake2, " ") > 0 Then tmpName = Mid(tmpLake2, srcSeek2 + i, Instr(srcSeek2 + i + 1, tmpLake2, " ") - srcSeek2 - i) Else tmpName = tmpLake2
					If InStr(tmpName, chr(9)) > 0 Then tmpName = Mid(tmpName, 1, Instr(1, tmpName, chr(9)) - 1)
					If InStr(tmpName, vbCrLf) > 0 Then tmpName = Mid(tmpName, 1, Instr(1, tmpName, vbcrlf) - 1)
					If InStr(tmpName, ">") > 0 Then tmpName = Mid(tmpName, 1, Instr(1, tmpName, ">") - 1)
				End If
				Call ScanFile( Mid(FilePath,1,InStrRev(FilePath,"\"))&tmpName , replace(FilePath,server.MapPath("\")&"\","",1,1,1))
				SumFiles = SumFiles + 1
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing

		'Check Crea"&"teObject
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "CreateO"&"bject[ |\t]*\(.*\)"
		Set Matches = regEx.Execute(filetxt)
		For Each Match in Matches
			If Instr(Match.Value, "&") or Instr(Match.Value, "+") or Instr(Match.Value, """") = 0 or Instr(Match.Value, "(") <> InStrRev(Match.Value, "(") Then
				Report = Report&"<tr class=""forumRow""><td>"&temp&"</td><td>Creat"&"eObject</td><td>Crea"&"teObject函数使用了变形技术。可能是误报"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
				exit sub
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing
	end if
	set ofile = nothing
	set fsos = nothing
End Sub

'检查文件后缀，如果与预定的匹配即返回TRUE
Function CheckExt(FileExt)
	If DimFileExt = "*" Then CheckExt = True
	Ext = Split(DimFileExt,",")
	For i = 0 To Ubound(Ext)
		If Lcase(FileExt) = Ext(i) Then 
			CheckExt = True
			Exit Function
		End If
	Next
End Function

Function GetDateModify(filepath)
	Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(filepath) 
	s = f.DateLastModified 
	set f = nothing
	set fso = nothing
	GetDateModify = s
End Function

Function GetDateCreate(filepath)
	Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(filepath) 
	s = f.DateCreated 
	set f = nothing
	set fso = nothing
	GetDateCreate = s
End Function

Function tURLEncode(Str)
	temp = Replace(Str, "%", "%25")
	temp = Replace(temp, "#", "%23")
	temp = Replace(temp, "&", "%26")
	tURLEncode = temp
End Function

Sub ShowAllFile2(Path)
	Set FSO = CreateObject("Scripting.FileSystemObject")
	if not fso.FolderExists(path) then exit sub
	Set f = FSO.GetFolder(Path)
	Set fc2 = f.files
	For Each myfile in fc2
		If CheckExt(FSO.GetExtensionName(path&"\"&myfile.name)) Then
			Call IsFind(Path&"\"&myfile.name)
			SumFiles = SumFiles + 1
		End If
	Next
	Set fc = f.SubFolders
	For Each f1 in fc
		ShowAllFile2 path&"\"&f1.name
		SumFolders = SumFolders + 1
    Next
	Set FSO = Nothing
End Sub

Sub IsFind(thePath)
	theDate = GetDateModify(thePath)
	on error resume next
	theTmp = Mid(theDate, 1, Instr(theDate, " ") - 1)
	if err then exit Sub
	
	xDate = Split(request.Form("Search_Date"),";")
	
	If request.Form("Search_Date") = "ALL" Then ALLTime = True
	
	For i = 0 To Ubound(xDate)
		If theTmp = xDate(i) or ALLTime = True Then 
			If request("Search_Content") <> "" Then
				Set FSOs = CreateObject("Scripting.FileSystemObject")
				set ofile = fsos.OpenTextFile(thePath, 1, false, -2)
				filetxt = Lcase(ofile.readall())
				If Instr( filetxt, LCase(request.Form("Search_Content"))) > 0 Then
					temp = "<a href=""http://"&Request.Servervariables("server_name")&"/"&tURLEncode(Replace(replace(thePath,server.MapPath("\")&"\","",1,1,1),"\","/"))&""" target=_blank>"&replace(thePath,server.MapPath("\")&"\","",1,1,1)&"</a>"
					Report = Report&"<tr class=""forumRow""><td>"&temp&"</td><td>"&GetDateCreate(thePath)&"</td><td>"&theDate&"</td></tr>"
					Sun = Sun + 1
					Exit Sub
				End If
				ofile.close()
				Set ofile = Nothing
				Set FSOs = Nothing
			Else
				temp = "<a href=""http://"&Request.Servervariables("server_name")&"/"&tURLEncode(Replace(replace(thePath,server.MapPath("\")&"\","",1,1,1),"\","/"))&""" target=_blank>"&replace(thePath,server.MapPath("\")&"\","",1,1,1)&"</a>"
				Report = Report&"<tr class=""forumRow""><td>"&temp&"</td><td>"&GetDateCreate(thePath)&"</td><td>"&theDate&"</td></tr>"
				Sun = Sun + 1
				Exit Sub
			End If
		End If
	Next
	
End Sub
%>