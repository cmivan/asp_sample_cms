<!--#include file="../../system/core/initialize_system.asp"-->
<%
dim sss
sss = lcase(request.servervariables("QUERY_STRING"))
GuoLv="select,insert,;,update,',delete,exec,admin,count,drop"
GuoLvA=split(GuoLv,",")
for i=0 to ubound(GuoLvA)
  if instr(sss,GuoLvA(i))<>0 then
    response.Redirect "res://shdoclc.dll/dnserror.htm"
    response.end		
  end if
next

dim DataPathBackup,DataPath
    DataPathBackup = "_dbbackup"  '数据库备份目录
	DataPath = "../../" & dbpath
	
	
dim action
dim admin_flag
dim bkfolder,bkdbname,fso,fso1
Dim uploadpath
    action = trim(request("action"))
%>

<body>
<TABLE border="0" align="center" cellpadding="0" cellspacing="10" class="forum1">
<tr><td>

<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop table">
<tr class="forumRaw"><TD height=20 align="left">

<a class="btn btn-mini" href="?action=BackupData">
<span class="ico icon-plus">&nbsp;</span>&nbsp;备份数据</a>
<a class="btn btn-mini" href="?action=RestoreData">
<span class="ico icon-retweet">&nbsp;</span>&nbsp;恢复数据</a>
<a class="btn btn-mini" href="?action=CompressData">
<span class="ico icon-random">&nbsp;</span>&nbsp;数据压缩</a>

</TD></tr></table>

<%
'====================备份数据库=========================
sub BackupData()
%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop table">
<tr class="forumRow"><TD width="56%" height=20>
&nbsp;<B>备份数据</B>( 需要FSO支持，FSO相关帮助请看微软网站 )</TD>
<TD width="44%" class="forumRaw" >&nbsp;</TD>
</tr>
<form method="post" action="admin_db_backup.asp?action=BackupData&act=Backup">	
<tr class="forumRow">
<td height=20>
&nbsp;当前数据库路径(相对)：
<strong><font color="#FF0000"><%=DataPath%></font></strong></td>
<td class="forumRow"><span class="greentext">(读取网站数据库地址)</span></td>
</tr>	

<tr class="forumRow">
<td height=20>&nbsp;备份数据库目录(相对)：
<strong><font color="#FF0000"><%=DataPathBackup%></font></strong>&nbsp;&nbsp;&nbsp;&nbsp;</td>
<td class="forumRow"><span class="greentext">(如目录不存在，程序将自动创建)</span></td>
</tr>	

<tr class="forumRow">
<td height=20>&nbsp;备份数据库名称(名称)：
<strong><font color="#FF0000"><%=newDBname()%></font> </strong>&nbsp;&nbsp;&nbsp;&nbsp;</td>
<td class="forumRow"><span class="greentext">(如备份目录有该文件，将覆盖，如没有，将自动创建)</span></td>
</tr>	

<tr class="forumRaw"><td height=20>
<button type="submit" name="Submit13" class="btn" style="width:100%;"><span class="ico icon-ok">&nbsp;</span>&nbsp;备份数据</button>
</td>
<td class="forumRow"><span class="greentext">(用该功能来备份您的法规数据，以保证您的数据安全)</span></td>
</tr></form>
</table>
<%end sub%>



<%
'====================恢复数据库=========================
sub RestoreData() %>
<table border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop table">
<tr class="forumRow"><TD width="56%" height=20>&nbsp;<B>恢复数据</B>( 需要FSO支持，FSO相关帮助请看微软网站 )</TD>
<TD width="44%" class="forumRaw" >&nbsp;</TD>
</tr>

<form method="post" action="admin_db_backup.asp?action=RestoreData&act=Restore">
<tr class="forumRow"><td height=20>&nbsp;当前数据库路径(相对)：
<strong><font color="#FF0000"><%=DataPath%></font></strong></td>
<td class="forumRow">&nbsp;</td>
</tr>	
  				
<tr class="forumRow">
<td height=20>&nbsp;备份数据库目录(相对)：
<strong><font color="#FF0000"><%=DataPathBackup%></font></strong>				  </td>
<td class="forumRow"><font color="#FF0000" class="greentext">(注：所有路径都是相对与程序空间根目录的相对路径)</font></td>
</tr>	
  				
<tr class="forumRow">
<td height=20>&nbsp;备份数据库路径(相对)：
<input type=text size=30 name=thisDbPath value="<%=newDBname()%>"></td>
<td class="forumRow"><span class="greentext">(请填写备份数据库名称)</span></td>
</tr>
  				
<tr class="forumRow"><td>
<button type="submit" name="Submit13" class="btn" style="width:100%;"><span class="ico icon-ok">&nbsp;</span>&nbsp;恢复数据</button>
</td><td><font color="#FF0000"><span class="greentext">(用该功能来备份您的法规数据，以保证您的数据安全)</span></font></td>
</tr>	
</form>
</table>
<%end sub%>



<%
'====================压缩数据库 =========================
sub CompressData()
%>
<!-- 以下颜色不同部分为客户端界面代码 -->
<table border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop table">
<form action="?action=CompressData" method="post">
<tr class="forumRow">
<td>
&nbsp;<b>注意：</b>
输入数据库所在相对路径,并且输入数据库名称（正在使用中数据库不能压缩，请选择备份数据库进行压缩操作） </td>
</tr>
<tr class="forumRow">
<td>&nbsp;压缩数据库：
<input name="thisDbPath" type="text" value='<%=DataPath%>' size="60">
<button type="submit" name="Submit13" class="btn"><span class="ico icon-ok">&nbsp;</span>&nbsp;开始压缩</button>
</td></tr>
<tr class="forumRow">
<td>&nbsp;<input type="checkbox" name="boolIs97" value="True">
  &nbsp;如果使用 Access 97 数据库请选择
(默认为 Access 2000 数据库)</td>
</tr>
</form>
</table>
<%end sub%>







<%
'备份数据
select case action
case "BackupData"		'备份数据
    call BackupData()
	if request("act") = "Backup" then call updata()
	
case "RestoreData"		'恢复数据
	dim backpath
	call RestoreData()
	
	if request("act") = "Restore" then
	    '获取备份文件地址
	    thisDbPath   = DataPathBackup&"/"&request("thisDbPath")
	    '获取当前数据库地址
	    backpath = DataPath
		if thisDbPath="" then
		   call alert.reFresh("请输入您要恢复成的数据库全名","?action=RestoreData")
		else
		   thisDbPath = server.mappath(thisDbPath)
		end if
		backpath = server.mappath(backpath)
		set Fso=server.createobject("scripting.filesystemobject")
		if fso.fileexists(thisDbPath) then  					
		   fso.copyfile thisDbPath,Backpath
		   call alert.reFresh("成功恢复数据！","?action=RestoreData")
		else
		   call alert.reFresh("备份目录下并无您的备份文件！","?action=RestoreData")
		end if
	end if
case "CompressData"		'数据压缩
    CompressData()
	
	dim thisDbPath,boolIs97
	thisDbPath = request("thisDbPath")
	boolIs97 = request("boolIs97")
	if thisDbPath <> "" then
	   thisDbPath = server.mappath(thisDbPath)
	   '调用服务器端的自定义函数 CompactDB 来压缩数据库
	   call CompactDB(thisDbPath,boolIs97)
	end If
end select



'-------------新加函数-------------------
function newDBname()
   newDBname = date()
   newDBname = newDBname & ".bak"
end function
'-------------检查某一目录是否存在-------------------
Function CheckDir(FolderPath)
	folderpath=Server.MapPath(".")&"\"&folderpath
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    If fso1.FolderExists(FolderPath) then
       CheckDir = True '存在
    else
       CheckDir = False '不存在
    end if
    Set fso1 = nothing
End Function
'-------------根据指定名称生成目录-------------------
Function MakeNewsDir(foldername)
	dim f
    Set fso1 = CreateObject("Scripting.FileSystemObject")
        Set f = fso1.CreateFolder(foldername)
        MakeNewsDir = True
    Set fso1 = nothing
End Function



sub updata()
    '原数据库路径
	thisDbPath = DataPath
	thisDbPath=server.mappath(thisDbPath)
	bkfolder=DataPathBackup    '备份目录
	bkdbname=newDBname()        '备份名称
	Set Fso=server.createobject("scripting.filesystemobject")
	if fso.fileexists(thisDbPath) then
		If CheckDir(bkfolder) = True Then
		fso.copyfile thisDbPath,bkfolder& "\"& bkdbname
		else
		MakeNewsDir bkfolder
		fso.copyfile thisDbPath,bkfolder& "\"& bkdbname
		end if
		call alert.reFresh("成功备份数据库,备份路径为：" &bkfolder& "\"& bkdbname,"?action=BackupData")
	Else
		call alert.reFresh("找不到您所需要备份的文件!","?action=BackupData")
	End if
end sub




'以下为实际压缩数据库的自定义函数，在服务器端运行
'[压缩参数]
Function CompactDB(thisDbPath, boolIs97)
  on error resume next '容错模式
  dim fso, Engine, strthisDbPath,JET_3X
  strthisDbPath = left(thisDbPath,instrrev(thisDbPath,"\"))
  set fso = CreateObject("Scripting.FileSystemObject")
  If fso.FileExists(thisDbPath) Then
     Set Engine = CreateObject("JRO.JetEngine")

'其实，和在Access中压缩数据库一样，我们仍然调用 JRO 来压缩修复数据库
'所不同的是在这里我们没有向Access那样采用“先引用”的方式（工具菜单选择引用）
'而是采用脚本所能使用的“后引用”方式建立 JRO 的实例 CreateObject("JRO.JetEngine")

     Randomize
     an=""
     an= int((999-222+1) * RND +222)
     If boolIs97 = "True" Then
       Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & thisDbPath,"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strthisDbPath & "temp"&an&".mdb;"& "Jet OLEDB:Engine Type=" & JET_3X
     Else
       Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & thisDbPath,"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strthisDbPath & "temp"&an&".mdb"
     End If
	 '操作完成后将已经缩小体积的数据库 COPY 回原位，覆盖原始文件
	 fso.CopyFile strthisDbPath & "temp"&an&".mdb",thisDbPath
	 '删除无用的临时文件
	 fso.DeleteFile(strthisDbPath & "temp"&an&".mdb")
	 
     Set fso = nothing
	Set Engine = nothing
	call alert.reFresh("·数据库," & thisDbPath & ",已经压缩成功!","?action=CompressData")
  Else
	call alert.reFresh("数据库名称或路径不正确,请重试!","?action=CompressData")
  End If
End Function
%>


</td></tr></table>