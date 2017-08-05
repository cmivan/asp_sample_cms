<!--#include file="../../system/core/initialize_system.asp"-->
<%
 response.Buffer=true
 on error resume next '容错模式
'***********************************
'     配置，定义变量
'***********************************
 Dim fixs,Path,db_Table,fPath,mPath
    '-----------
	 mPath = "../../../up/"
     fPath = server.MapPath(mPath)
    '-----------
     fixs="|.gIf|.jpg|.png|.txt|"    '注意定义结构
	 fixs=lcase(fixs)
    '-----------
     Path=request.Form("Path")
	 if Path<>"" then
	    session("Path")=Path      '记录当前目录
	 elseif Path="" and session("Path")<>"" then
	    Path=session("Path")
	 else
	    Path=fPath
	 end if
     
    '-----------
	 db_Table=request.Form("Tables")
	 if len(db_Table)<3 then
	    db_Table="product"             '指定操作的表
	 end if
%>






<%
'***********************************
'    调用分类名称
'***********************************
function get_typeName(gid)
    If gid<>"" and isnumeric(gid) then
       set typeName_rs=server.CreateObject("adodb.recordset")
	       typeName_sql="select * from "&db_Table&"_type where id="&int(gid)
	       typeName_rs.open typeName_sql,conn,1,1
	         If not typeName_rs.eof then get_typeName=typeName_rs("title")  '读取指定字段
	       typeName_rs.close  
       set typeName_rs=nothing
    End If
End function



'***********************************
'    判断目录是否存在
'***********************************
Function CheckDir(ckDirname)
   Dim M_fso
   CheckDir=False
   Set M_fso=CreateObject("Scripting.FileSystemObject")
       If (M_fso.FolderExists(ckDirname)) Then
           CheckDir=True
       End If
   Set M_fso = Nothing
End Function


'***********************************
'    用于输出提示行
'***********************************
Sub Print(str)
   response.Write("<li onmouseover=""this.className='on';"" onmouseout=""this.className='out';""> "&str&"</li>")
   response.Flush()
End Sub


'***********************************
'    样式函数   
'***********************************
Function css(str,s)
    css="<span class="""&s&""" >"&str&"</span>"
End Function










'***********************************
'   查找分类目录下的产品，并录入到数据库
'***********************************
Function DirProduct(TOpath,tid)
  on error resume next     '容错模式
  Set objFso = CreateObject("Scripting.FileSystemObject")
  Set CrntFolder = objFso.GetFolder(TOpath)
      F_File=false
      For Each ConFile In CrntFolder.Files
		 F_Name=ConFile.Name
		 
		 If len(F_Name)>4 then
		    F_fix=lcase(right(F_Name,4))
			F_fix="|"&F_fix&"|"
		   '*********** |检测文件类型是否符合| ************
			If instr(fixs,F_fix)<>0 then

		
         '######################## |写入数据库（重要） | ######################### 
			   set rs=server.CreateObject("adodb.recordset")
			       rs.open "select * from "&db_Table,conn,1,3
			       rs.addnew
				   
				   rs("title")    = left(F_Name,len(F_Name)-4)
			       rs("small_pic")= Path_X_S&"/"&tid&"/"&F_Name
				   rs("big_pic")  = Path_X_B&"/"&tid&"/"&F_Name
				  '--------------------------------
				   content        =Path_X_B&"/"&tid&"/"&F_Name
				   content        ="<p align=""center""><img src="""&content&""" /></p>"
			       rs("content")  = content
			       rs("type_id")  = tid
				   
				  '在未出错的情况下才写入数据库
				   if err=0 then rs.update
				   rs.close
			   set rs=nothing
		  '####################################################################

			 
			   if err=0 then    '数据库成功记录
			      Print("&nbsp;- 已录入: "&ConFile.Name&"  …… "&css("√","green"))
			   else
			      Print("&nbsp;- 检测到文件,录入失败: "&ConFile.Name&"  …… "&css("×","red"))
			   end if
			   F_File=true
			   
		    End If
		 End If


      Next
  Set CrntFolder = nothing
  Set objFso = nothing
  
 '#### 在指定目录下未检测到 相应内容 ####
  If F_File=false then call Print("&nbsp;- 未检测到指定类型 "&css(fixs,"strong")&" 相应的文件  …… "&css("Failed !","red"))
End Function
%>

<style>
body{
font-size:12px;
color:#333333;
font-family:Verdana, Arial, Helvetica, sans-serif;
line-height:20px;
}

body,tr,td,div{
font-size:11px;}

form{
margin:0;}

input{
	border:#666666 1px solid;
	border-bottom:#FFFFFF 1px solid;
	border-right:#FFFFFF 1px solid;
	background-color:#E2E2C7;
	color:#333333;
	width:100%;

}

a{
color:#CCCCCC;}


ul{
margin:12px;
padding-left:12px;
margin-top:0;
margin-bottom:0;}
li{
padding-left:6px;}

hr{
width:100%;
text-align:left;
background-color:#666666;
border-bottom:#666666 1px dotted;}

.red{
color:#FF0000;
font-size:10px;}
.green{
color:#00FF00;
font-size:10px;}
.strong{
color:#FF0000;
font-weight:bold;
font-size:12px;}

#main_div{
width:635px;
margin:auto;
background-color:#EBEBD8;
padding:12px;}
#main_div a{
text-decoration:none;
font-size:11px;
color:#EAEAEA;}
#main_div a:hover div{
background-color:#EBEBD8;}



#main_box{
width:600px;
margin:auto;
background-color:#F4F4EA;
padding:15px;
border:#CCCCCC 1px dotted;
border-bottom:#CCCCCC 1px dotted;
border-right:#CCCCCC 1px dotted;
}



/*目录样式*/
#Folders{
margin:10px;}
#Folders a{
float:left;
padding-top:10px;
padding-bottom:5px;
width:118px;
height:70px;
color:#FB9700;
text-decoration:none;
text-align:center;
float:left;
border:#EBEBD8 1px solid;
border-bottom:#EBEBD8 2px solid;
}
#Folders a:hover{
color:#FF0000;
border:#EBEBD8 1px solid;
border-bottom:#FF6600 2px solid;
background-color:#FFFFFF;
}
#Folders font{
line-height:80%;
}

.on{
background-color:#EBEBD8}
.out{
background-color:#F4F4EA;}
</style>


<%if request.QueryString("page")<>"folder" then%>
<script type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>



<div id="main_div">
<div id="main_box">
<form id="addform" name="addform" method="post" action="">
<table width="100%" border="0" cellpadding="0" cellspacing="5" onMouseOver="this.bgColor='#EBEBD8';" onMouseOut="this.bgColor='';">
  <tr>
    <td>
  {选定路径}/small&nbsp;&nbsp;&nbsp;'小图相对目录
  <br>
  {选定路径}/big&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'大图相对目录
</td></tr>
  <tr>
    <td>
      <input name="Path" type="text" id="Path" value="<%=Path%>" size="80%" />
      
	  <input name="P_Form" type="hidden" id="P_Form" value="ok"/>
	  </td>
    <td width="60" style="padding-left:5px;">
	<input name="Tables" type="text" id="Tables" value="<%=db_Table%>" style="width:60px;" />
	</td>
    <td width="120">
    <button type="button" name="Submit2" onClick="MM_openBrWindow('?page=folder','Folder','width=380,height=280')" class="btn"><span class="ico icon-picture">&nbsp;</span>&nbsp;浏览</button>
    <input type="submit" name="Submit2" value="录入" class="btn"/></td>
  </tr>
</table>
</form>



<ul type="a">
<%
Dim i,TempStr,FlSpace
  
If Path<>"" and request.Form("P_Form")="ok" then
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  onPath= fPath
  xPath = replace(lcase(Path),lcase(onPath),"")   '获取与跟目录的相对目录
  xPath = mPath&xPath
  

  Path_X_S=xPath&"/small"    '小图相对目录
  Path_X_B=xPath&"/big"      '大图相对目录
  
  Path_X_S=replace(Path_X_S,"\","/")
  Path_X_S=replace(Path_X_S,"//","/")
  
  Path_X_B=replace(Path_X_B,"\","/")
  Path_X_B=replace(Path_X_B,"//","/")

  Path_S=server.MapPath(Path_X_S)     '小图全目录
  Path_B=server.MapPath(Path_X_B)     '大图全目录
  
 '######## 相对目录处理 #########
  Path_X_S=replace(Path_X_S,"../","/")
  Path_X_S=replace(Path_X_S,"./","/")
  Path_X_B=replace(Path_X_B,"../","/")
  Path_X_B=replace(Path_X_B,"./","/")

  Path_X_S=replace(Path_X_S,"//","/")
  Path_X_B=replace(Path_X_B,"//","/")
  
  Path_X_S=replace(Path_X_S,"/userfiles","userfiles")
  Path_X_B=replace(Path_X_B,"/userfiles","userfiles")
  
  


'-------- 判断路径是否有效,是否存在 ---------
  call Print(css("ECHO ↓","green"))
  'response.Write("<br />")
  call Print("当前目录:"&Path)
  
     Path_Ok=true
  If CheckDir(Path_S)=FALSE then
	 call Print("检测小图目录:"&Path_S&"  …… "&css("Failed !","red"))
	 Path_Ok=false
  else
     call Print("检测小图目录:"&Path_S&"  …… "&css("ok !","green"))
  End If
  
  If CheckDir(Path_B)=FALSE then
	 call Print("检测大图目录:"&Path_B&"  …… "&css("Failed !","red"))
	 Path_Ok=false
  else
	 call Print("检测大图目录:"&Path_B&"  …… "&css("ok !","green"))
  End If

  
  
  
  
  
'-------- 读取小图目录， 获取分类文件夹---------
If Path_Ok then
  '同时检测到大图和小图目录，才进行读取文件(go)
	  
  call Print("正在分析目录  …… ")
  Set objFso = CreateObject("Scripting.FileSystemObject")
  Set CrntFolder = objFso.GetFolder(Path_S)

  For Each SubFolder In CrntFolder.SubFolders
	  F_id=SubFolder.Name

	  If F_id<>"" and isnumeric(F_id) then
	  
	  '------------------------
	  If get_typeName(F_id)="" then
		 call Print(" →  目录"&css(" &lt;"&F_id&"&gt; ","strong")&"  未找到相应分类!  …… "&css("ok !","green"))
	  else
	  
	     response.Write("<hr />")
		 call Print(" →  检测到分类目录"&css(" &lt;"&get_typeName(F_id)&"&gt; ","strong")&"! 正检测目录下的文件 …… "&css("ok !","green"))

		'读取分类目录下的所有文件，并把符合条件的写入
		 call DirProduct(Path_S&"\"&F_id,F_id)
	  End If
	  '------------------------

	  else
	     If F_id="" then call Print(css(" &lt; 空目录 &gt; ","strong")&"  不符合!  …… "&css("Failed !","red"))
		 If isnumeric(F_id)=false then call Print("目录"&css(" &lt; "&F_id&" &gt; ","strong")&"  不符合分类要求!  …… "&css("Failed !","red"))
	  End If

  Next
  
  Set CrntFolder = nothing
  Set objFso = nothing
 '同时检测到大图和小图目录，才进行读取文件(end)
End If
  
  response.Write("<br /><br />")
If Path_Ok then
   call Print(css("完成录入!","strong")&"  …… "&css("ok !","green"))  
Else
   call Print(css("条件不符,录入失败!","strong")&"  …… "&css("Failed !","red"))  
end if

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End If
  
%>

</ul>
</div>
</div>



<%else%>
<!--页面：文件夹选择-->

<%
 onPath=trim(request.QueryString("Path"))
 upPath=request.QueryString("upPath")
if instr(onPath,"..")<>0 or instr(upPath,"..")<>0 then response.End()
'===============================
 if session("onPath")="" and onPath<>"" then
	if CheckDir(fPath&"/"&onPath) then
       session("onPath")  =fPath&"\"&onPath
    end if
 elseif onPath<>"" then
	if CheckDir(session("onPath")&"/"&onPath) then
       session("onPath")=session("onPath")&"\"&onPath
    end if
 elseif session("onPath")="" and onPath="" and CheckDir(fPath) then
    session("onPath")=fPath
 end if

'////// 上一级目录 ////
if upPath="yes" then
   dim tempPath1,tempPath2,tempPath,arrPath1,arrPath2
       temponPath1=fPath
       temponPath2=session("onPath")
       arrPath1=split(temponPath1,"\")
       arrPath2=split(temponPath2,"\")
if ubound(arrPath1)<ubound(arrPath2) then
   for tempI=0 to ubound(arrPath2)-1
       if tempI=ubound(arrPath2)-1 and tempI<>0 then
          tempPath=tempPath&arrPath2(tempI)
       else
          tempPath=tempPath&arrPath2(tempI)&"\"
       end if
   next
   if CheckDir(tempPath) then session("onPath")=tempPath
end if
end if

 'session.Abandon()
response.Write(session("onPath"))
%>

<script>
function Open_path(str){
	   window.location.href="?page=<%=request("page")%>&path="+str;  
}
function To_path(str){
   var path;
       path="<%=replace(session("onPath"),"\","\\")%>\\";
	   window.opener.addform.Path.value=path+str;
	   window.Topath.Topath_str.value=path+str;
     //window.close();
}

function goto_path(){
    if(window.Topath.Topath_str.value!=""){
	  window.opener.addform.submit();
	  window.close();
	}else{
	  alert("请选择目录!");
	}
	   window.Topath.Topath_str.value=path+str;

}
</script>

<body style="margin:0; padding:0;">
<form name="Topath">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="forum2" >
  <tr class="forumRow">
    <td width="60%" style="padding-left:10px;"><div id="on_path">
      <input name="Topath_str" type="text" id="Topath_str" readonly />
    </div></td>
    <td align="center">
      <input type="button" name="Submit" value="录入"  onclick="goto_path();" class="btn"/>
      <%if ubound(arrPath1)<ubound(arrPath2) then%>
      <a class="btn" href="?page=<%=request("page")%>&upPath=yes"><span class="ico icon-arrow-up">&nbsp;</span>上一级</a>
      <%else%>
      <a class="btn disabled" href="javascript:void(0);"><span class="ico icon-arrow-up">&nbsp;</span>上一级</a>
      <%end if%></td>
    </tr>
</table>
</form>
<div id="Folders">
<%
'---------------------------------------
 Set objFso = CreateObject("Scripting.FileSystemObject")
  Set CrntFolder = objFso.GetFolder(session("onPath"))

  For Each SubFolder In CrntFolder.SubFolders
	  F_id=SubFolder.Name
%>
<a href="javascript:void(0);" onclick='To_path("<%=F_id%>");' ondblclick='Open_path("<%=F_id%>");'>
<font face='wingdings' size='+5'>1</font><br><%=F_id%>
</a>
<%
  Next

  Set CrntFolder = nothing
  Set objFso = nothing
%>
</div>
</body>
<%end if%>

