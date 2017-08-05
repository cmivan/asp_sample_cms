<!--#include file="../../system/core/initialize_system.asp"-->
<%
sss=LCase(request.servervariables("QUERY_STRING"))
GuoLv="select,insert,;,update,',delete,exec,admin,count,drop"
GuoLvA=split(GuoLv,",")
for i=0 to ubound(GuoLvA)
  if instr(sss,GuoLvA(i))<>0 then
    Response.Redirect "res://shdoclc.dll/dnserror.htm"
    response.end		
  end if
next

dim Data_Backup_path
    Data_Backup_path="db_backup"  '数据库备份目录
%>

<body>
<TABLE border="0" align="center" cellpadding="0" cellspacing="10" class="forum1">
<tr><td>

<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop">
<tr class="forumRaw"><TD height=20 align="left">&nbsp;<B><a href="?action=BackupData">备份数据</a></B> - <B><a href="?action=RestoreData">恢复数据</a></B> - <B><a href="?action=CompressData">数据压缩</a></B></TD>
 </tr>
</table>

 

<%
dim action
dim admin_flag
dim bkfolder,bkdbname,fso,fso1
Dim uploadpath
    action=trim(request("action"))
'////////////// 新加函数 ////////////////////////////
function new_dbname()
         new_dbname=date()
		 new_dbname=new_dbname&".bak"
end function
'////////////////////////////////////////////////////




	'备份数据
select case action
case "BackupData"		'备份数据
		if request("act")="Backup" then
			call updata()
		else
			call BackupData()
			
		end if

case "RestoreData"		'恢复数据
	dim backpath
		if request("act")="Restore" then
			Dbpath    =Data_Backup_path&"/"&request("DBpath")        '获取备份文件地址
			
			backpath  =mdbs                '获取当前数据库地址
			
			if dbpath="" then
			call backPage("请输入您要恢复成的数据库全名","?action=RestoreData",0)
			
			else
			Dbpath=server.mappath(Dbpath)
			end if
			backpath=server.mappath(backpath)
		
			Set Fso=server.createobject("scripting.filesystemobject")
			if fso.fileexists(dbpath) then  					
			fso.copyfile Dbpath,Backpath
			
			call backPage("成功恢复数据！","?action=RestoreData",0)
					
			else
			call backPage("备份目录下并无您的备份文件！","?action=RestoreData",0)
			
			end if
		else
			call RestoreData()
		end if

case "CompressData"		'数据压缩
      CompressData()
end select

%>






<%
sub updata()
		Dbpath=mdbs               '原数据库路径
		Dbpath=server.mappath(Dbpath)
		bkfolder=Data_Backup_path    '备份目录
		bkdbname=new_dbname()        '备份名称
		Set Fso=server.createobject("scripting.filesystemobject")
		if fso.fileexists(dbpath) then
			If CheckDir(bkfolder) = True Then
			fso.copyfile dbpath,bkfolder& "\"& bkdbname
			else
			MakeNewsDir bkfolder
			fso.copyfile dbpath,bkfolder& "\"& bkdbname
			end if
            call backPage("成功备份数据库,备份路径为：" &bkfolder& "\"& bkdbname,"?action=BackupData",0)
		Else
			call backPage("找不到您所需要备份的文件!","?action=BackupData",0)
		End if
end sub
%>




<%
'====================备份数据库=========================
sub BackupData()
If IsSqlDataBase = 1 Then
	SQLUserReadme()
	Exit Sub
End If
%>

<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forum2">
<tr class="forumRow"><TD width="56%" height=20>
&nbsp;<B>备份数据</B>( 需要FSO支持，FSO相关帮助请看微软网站 )</TD>
<TD width="44%" class="forumRaw" >&nbsp;</TD>
</tr>
<form method="post" action="admin_db_backup.asp?action=BackupData&act=Backup">	
<tr class="forumRow">
<td height=20>
&nbsp;当前数据库路径(相对)：<strong><font color="#FF0000"><%=mdbs%></font></strong></td>
<td class="forumRow"><span class="greentext">(读取网站数据库地址)</span></td>
</tr>	

<tr class="forumRow">
<td height=20>&nbsp;备份数据库目录(相对)：
<strong><font color="#FF0000"><%=Data_Backup_path%></font></strong>&nbsp;&nbsp;&nbsp;&nbsp;</td>
<td class="forumRow"><span class="greentext">(如目录不存在，程序将自动创建)</span></td>
</tr>	

<tr class="forumRow">
<td height=20>&nbsp;备份数据库名称(名称)：
<strong><font color="#FF0000"><%=new_dbname()%></font> </strong>	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
<td class="forumRow"><span class="greentext">(如备份目录有该文件，将覆盖，如没有，将自动创建)</span></td>
</tr>	

<tr class="forumRaw"><td height=20>
<input type=submit value="备份数据" style="width:100%;" class="button"></td>
<td class="forumRow"><span class="greentext">(用该功能来备份您的法规数据，以保证您的数据安全)</span></td>
</tr></form>
</table>
<%
end sub
%>



<%
'====================恢复数据库=========================
sub RestoreData()
%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forum2">
<tr class="forumRow"><TD width="56%" height=20>&nbsp;<B>恢复数据</B>( 需要FSO支持，FSO相关帮助请看微软网站 )</TD>
<TD width="44%" class="forumRaw" >&nbsp;</TD>
</tr>

<form method="post" action="admin_db_backup.asp?action=RestoreData&act=Restore">
<tr class="forumRow"><td height=20>&nbsp;当前数据库路径(相对)：
<strong><font color="#FF0000"><%=mdbs%></font></strong></td>
<td class="forumRow">&nbsp;</td>
</tr>	
  				
<tr class="forumRow">
<td height=20>&nbsp;备份数据库目录(相对)：
<strong><font color="#FF0000"><%=Data_Backup_path%></font></strong>				  </td>
<td class="forumRow"><font color="#FF0000" class="greentext">(注：所有路径都是相对与程序空间根目录的相对路径)</font></td>
</tr>	
  				
<tr class="forumRow">
<td height=20>&nbsp;备份数据库路径(相对)：
<input type=text size=30 name=DBpath value="<%=new_dbname()%>"></td>
<td class="forumRow"><span class="greentext">(请填写备份数据库名称)</span></td>
</tr>
  				
<tr class="forumRow">
<td><input name="submit" type=submit style="width:100%;border: 1px solid #FFFFFF" value="恢复数据"></td>
<td><font color="#FF0000"><span class="greentext">(用该功能来备份您的法规数据，以保证您的数据安全)</span></font></td>
</tr>	
</form>
</table>


<%
end sub

'------------------检查某一目录是否存在-------------------
Function CheckDir(FolderPath)
	folderpath=Server.MapPath(".")&"\"&folderpath
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    If fso1.FolderExists(FolderPath) then
       '存在
       CheckDir = True
    Else
       '不存在
       CheckDir = False
    End if
    Set fso1 = nothing
End Function
'-------------根据指定名称生成目录-----------------------
Function MakeNewsDir(foldername)
	dim f
    Set fso1 = CreateObject("Scripting.FileSystemObject")
        Set f = fso1.CreateFolder(foldername)
        MakeNewsDir = True
    Set fso1 = nothing
End Function
%>








<%
'====================压缩数据库 =========================
sub CompressData()
%><!-- 以下颜色不同部分为客户端界面代码 -->
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forum2">
<form action="?action=CompressData" method="post">
<tr class="forumRow">
<td>
&nbsp;<b>注意：</b>
输入数据库所在相对路径,并且输入数据库名称（正在使用中数据库不能压缩，请选择备份数据库进行压缩操作） </td>
</tr>
<tr class="forumRow">
<td>&nbsp;压缩数据库：
  <input name="dbpath" type="text" value='<%=mdbs%>' size="60">
<input type="submit" value="开始压缩"></td>
</tr>
<tr class="forumRow">
<td>&nbsp;<input type="checkbox" name="boolIs97" value="True">
  &nbsp;如果使用 Access 97 数据库请选择
(默认为 Access 2000 数据库)</td>
</tr>
</form>
</table>
<%
dim dbpath,boolIs97
dbpath = request("dbpath")
boolIs97 = request("boolIs97")

If dbpath <> "" Then
dbpath = server.mappath(dbpath)
    '调用服务器端的自定义函数 CompactDB 来压缩数据库
    call CompactDB(dbpath,boolIs97)
End If

end sub
%>


<%
'以下为实际压缩数据库的自定义函数，在服务器端运行
'=====================压缩参数=========================
Function CompactDB(dbPath, boolIs97)
 on error resume next '容错模式
Dim fso, Engine, strDBPath,JET_3X
    strDBPath = left(dbPath,instrrev(DBPath,"\"))
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(dbPath) Then
Set Engine = CreateObject("JRO.JetEngine")

'其实，和在Access中压缩数据库一样，我们仍然调用 JRO 来压缩修复数据库
'所不同的是在这里我们没有向Access那样采用“先引用”的方式（工具菜单选择引用）
'而是采用脚本所能使用的“后引用”方式建立 JRO 的实例 CreateObject("JRO.JetEngine")

Randomize
an=""
an= int((999-222+1) * RND +222)
    If boolIs97 = "True" Then
       Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath,"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp"&an&".mdb;"& "Jet OLEDB:Engine Type=" & JET_3X
    Else
       Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath,"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp"&an&".mdb"
    End If
'操作完成后将已经缩小体积的数据库 COPY 回原位，覆盖原始文件
fso.CopyFile strDBPath & "temp"&an&".mdb",dbpath
'删除无用的临时文件
fso.DeleteFile(strDBPath & "temp"&an&".mdb")

Set fso = nothing
Set Engine = nothing

    CompactDB = "·数据库," & dbpath & ",已经压缩成功!"
	call backPage(CompactDB,"",0)

Else
    CompactDB = "数据库名称或路径不正确,请重试!"
	call backPage(CompactDB,"",0)
	
End If

End Function
%>


</td>
  </tr>
</table>
