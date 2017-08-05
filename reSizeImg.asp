<%
'***********************************
'    调整图片尺寸
'***********************************
Class reSizeImg

  public thisSize
  public jpegInstall

  '// 类初始化
  private sub Class_Initialize
     thisSize = 250
	 jpegInstall = false
	 if obj_install("Persits.Jpeg") then jpegInstall = true
  end sub
 
  public function thisReSize(path)
	 on error resume next
	 thisReSize = "<img src=""" & reSizeImg(path,thisSize) & """ width=""" & thisSize & """ />"
  end function
  
  '// 调整图片尺寸
  public function reSizeImg(path,jsize)
	 on error resume next
	 reSizeImg = path
	 
	 '// 判断是否安装组件
	 if jpegInstall then
		 dim JWidth,JHeight,Jpath,Jname,Npath,Ipath
		 Jpath = server.MapPath(path)
		 Jname = getImgName(path)
		 if Jname<>"" then
			 Npath = "up/fck/image_small/" & Jname
			 Ipath = server.MapPath("./" & Npath)
			 if imgIsExits(Ipath)=false then
			   Set Jpeg = server.CreateObject("Persits.Jpeg")
				   Jpeg.open Jpath
				   if Jpeg.OriginalHeight>Jpeg.OriginalWidth then
					  JWidth      = jsize*Jpeg.OriginalWidth/Jpeg.OriginalHeight
					  Jpeg.Width  = JWidth
					  Jpeg.Height = jsize
				   else
					  JHeight     = jsize*Jpeg.OriginalHeight/Jpeg.OriginalWidth
					  Jpeg.Width  = jsize
					  Jpeg.Height = JHeight
				   end if
				   Jpeg.save Ipath
			   Set Jpeg = Nothing
			 end if
			 reSizeImg = Npath
			 if err<>0 then reSizeImg = path
		 end if
	 end if
	 
  end function
  
  '// 判断文件是否已经存在
  public function imgIsExits(path)
	 on error resume next
	 Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
		 If oFSO.FileExists( path ) Then
			imgIsExits = true
		 else
			imgIsExits = false
		 end if
	 Set oFSO = nothing
  end function
  
  '// 知道文件路径，获取文件名称
  public function getImgName(path)
	 on error resume next
	 dim pathArr,arrNum
	 getImgName = ""
	 pathArr = replace(path,"\","/")
	 pathArr = split(pathArr,"/")
	 arrNum = ubound(pathArr)
	 if arrNum>1 then getImgName = pathArr(arrNum)
  end function
  
  '// 判断组件是否已经安装
  public function obj_install(strClassString)
	 on error resume next
	 obj_install=False
	 dim xTestObj
	 err = 0
	 set xTestObj = Server.CreateObject(strClassString)
	 If Err=0 Then obj_install = True
	 Set xTestObj=Nothing
	 Err = 0
  end function
  
end Class

set reSize = new reSizeImg
%>
<%'=reSize.thisReSize("up/fck/image/design/case_12_12_31/1Q84/01.jpg")%>