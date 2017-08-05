<%
'*********************
'<><>< 文件处理类 ><><>
'*********************

Class fileClass

  '***********************************
  '   类初始化
  '***********************************
  Private Sub Class_Initialize
  
  End Sub
  
  
  '***********************************
  '   判断文件是否已经存在
  '***********************************
  public function isExits(fullpath)
	 on error resume next
	 Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
		 If oFSO.FileExists(fullpath) Then
			isExits = true
		 else
			isExits = false
		 end if
	 Set oFSO = nothing
  end function
  
  
  '****************************************
  '** 生成文件,可指定生成编码
  '****************************************
  public function writeToFile(fullpath,Str)
	 Set stm = CreateObject("Adodb.Stream")
		 stm.Type = 2
		 stm.mode = 3
		 stm.charset ="GB2312"
		 stm.Open
		 stm.WriteText Str
		 stm.SaveToFile fullpath, 2
		 stm.flush
		 stm.Close
	 Set stm = Nothing
  end function
  
  
  '***********************************
  '   知道文件路径，获取文件名称
  '***********************************
  public function getName(path)
	 on error resume next
	 dim pathArr,arrNum
	 getName = ""
	 pathArr = replace(path,"\","/")
	 pathArr = split(pathArr,"/")
	 arrNum  = ubound(pathArr)
	 if arrNum>1 then getName = pathArr(arrNum)
  end function
  
  
  '***********************************
  '   判断组件是否已经安装
  '***********************************
  public function isInstall(str)
	 on error resume next
	 isInstall = false
	 dim testObj
	 err = 0
	 set testObj = Server.CreateObject(str)
	 If err=0 Then isInstall = true
	 err = 0
	 set testObj=Nothing
  end function
  
  
  '***********************************
  '   生成目录，支持多级生成
  '***********************************
  public function folderCreate(path)
      on error resume next
	  dim patharr,path_level,i,pathtmp,cpath,CreateDIR,fso
	  path = Server.Mappath(path)
	  path = replace(path,"\","/")
      set fso = server.createobject("Scripting.FileSystemObject")
          patharr = split(path,"/")
          path_level = ubound(patharr)
          for i = 0 to path_level
              if i=0 then pathtmp = patharr(0) & "/" else pathtmp = pathtmp & patharr(i) & "/"
              cpath = left(pathtmp,len(pathtmp)-1)
              if not fso.FolderExists(cpath) then fso.CreateFolder(cpath)
          next
      set fso = nothing
	  if err.number<>0 then
		 folderCreate = false
		 err.Clear
	  else
		 folderCreate = true
	  end if	  
  end function
  
  
  
  '****************************************
  '** 创建目录 \ 支持多级 \文件夹不存在则生成
  '****************************************
  public Function creatfolder(w_path)
	  on error resume next
	  set fso = server.createobject("Scripting.FileSystemObject")
		  w_path=replace(w_path,"/","\")
		  w_path=split(w_path,"\")
		  ' 有后缀的--------------------------------
		  i_path=ubound(w_path)
		  if instr(w_path(ubound(w_path)),".")<>0 then i_path=i_path-1
		 '-----------------------------------------
		  for i=0 to i_path
			  if w_path(i)<>"" then
				 w_paths=w_paths&w_path(i)&"\"
				 if fso.folderexists(w_paths)=false then fso.createfolder(w_paths)
			  end if
		  next
		  creatfolder=true            
	  set fso=nothing
  End Function
  
  
  '****************************************
  '** 创建文件,可指定生成编码
  '****************************************
  public function creatfile(filePath,nr)
	  on error resume next
	  Temp_filePath=filePath
	  Temp_filePath=replace(Temp_filePath,"\\","\")
	  filePath=replace(filePath,"\\","\")
	  filePath=server.MapPath(filePath)
	  T_filePath =filePath
	  
	  '//创建目录
	  call creatfolder(filePath)
	  
	  '判断文件大小，是否要生成
	  if getFileLen(nr) then
	     '调用编码转换
		 call writeToFile(T_filePath,nr)
	  else
		 back = "falis!(Too big!)"
	  end if
	  
	  if err and back="" then
		 back="falis!"
	  elseif back="" then
		 back="ok!"
	  end if
	  
	  response.Write("<a href="""&Temp_filePath&""" target=""_blank""> "&Temp_filePath&"</a> "&back&"<br>")
	  response.Flush()
  end function 
  
  
  '****************************************
  '** 检测文件大小,用于生成文件
  '****************************************
  public function getFileLen(str)
	  on error resume next
	  getFileLen = true
	  dim l,t,i,c
	  l=len(str)
	  t=0
	  for i=1 to l
		 c=Abs(Asc(Mid(str,i,1)))
		 if c>255 then
			t=t+2
		 else
			t=t+1
		 end if
	  next
	  if int(HtmlLimit)<=0 then HtmlLimit=100
	  '限制HTML文件大小
	  if t>HtmlLimit*1024 then getFileLen = false
  end function
    

  '***********************************
  '   读取文件内容 
  '***********************************
  public function fileOpen(fullpath)
    on error resume next 
 	Set fs = CreateObject("Scripting.FileSystemObject") 
 	set f = fs.OpenTextFile(fullpath,1,True) 

	if err<>0 then
	   response.Write(err.description)
	   exit function
	end if
 	fileOpen = f.ReadAll()  
 	f.Close 
 	set fs=nothing 
  end function
  
  
  '***********************************
  '   显示文件列表 
  '***********************************
  public function fileList(fullpath)
	dim f, f1, fc, s
	set fso = server.CreateObject("scripting.filesystemobject")
	if (fso.FolderExists(fullpath)) Then
	   set f = fso.GetFolder(fullpath)
	   set fc = f.Files
	   for Each f1 in fc
	      s = s & f1.name 
	      s = s & "|"
	   next
	   fileList = s
	else
	   fileList = -1
	end if
  end Function
  
  
  '****************************************
  '** 文件的删除,用于生成文件
  '****************************************
  public function del(fullpath)
      on error resume next 
	  set fs = server.CreateObject("Scripting.FileSystemObject")
	   if fs.FileExists(fullpath) then fs.DeleteFile(fullpath)
	  set fs = Nothing
  end function
 
 
End Class



'*********************
'<><>< 实例化对象 ><><>
'*********************
set file2 = New fileClass
%>
