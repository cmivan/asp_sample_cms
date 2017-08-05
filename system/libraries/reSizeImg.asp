<%
'*********************
'<><>< 调整图片尺寸类 ><><>
'*********************
Class reSizeImgClass

  public thisSize
  public jpegInstall


  '***********************************
  '   类初始化
  '***********************************
  private sub Class_Initialize
     thisSize = 250
	 if file2.isInstall("Persits.Jpeg") then
	    jpegInstall = true
	 else
	    jpegInstall = false
	    response.Write("No Persits.Jpeg!")
	 end if
  end sub
  

  '***********************************
  '   调整图片尺寸
  '***********************************
  public function reSizeImg(path,jsize)
	 on error resume next
	 reSizeImg = path
	 
	 if text.isNum(jsize) = false then jsize = thisSize
	 
	 '// 判断是否安装组件
	 if jpegInstall then
	 
		 dim jWidth,jHeight,jPath,jMd5,jName,nPath,iPath
		 jMd5 = md5( path )
		 jPath = server.MapPath( path )
		 jName = file2.getName( path )
		 if jName<>"" then
			 nPath = "up/fck/image_small/" & jsize & "@" & jMd5 & "@" & jName
			 iPath = server.MapPath("./" & nPath)
			 if file2.isExits(iPath)=false then
			   Set Jpeg = server.CreateObject("Persits.Jpeg")
				   Jpeg.open jPath
				   if Jpeg.OriginalHeight>Jpeg.OriginalWidth then
					  jWidth      = jsize*Jpeg.OriginalWidth/Jpeg.OriginalHeight
					  Jpeg.Width  = jWidth
					  Jpeg.Height = jsize
				   else
					  jHeight     = jsize*Jpeg.OriginalHeight/Jpeg.OriginalWidth
					  Jpeg.Width  = jsize
					  Jpeg.Height = jHeight
				   end if
				   Jpeg.save iPath
			   Set Jpeg = Nothing
			 end if
			 reSizeImg = nPath
			 if err<>0 then reSizeImg = path
		 end if
	 end if
	 
  end function
  
  
  
  '***********************************
  '   读取图片
  '***********************************
  public function getImg(path,width)
	 on error resume next
	 getImg = false
	 if text.isNum(width) = false then width = thisSize
	 '// 判断是否安装组件
	 if jpegInstall then
	   Set Jpeg = server.CreateObject("Persits.Jpeg")
		   Jpeg.open path
		   Jpeg.Width  = width
		   Jpeg.Height = width*Jpeg.OriginalHeight/Jpeg.OriginalWidth
		   Jpeg.Quality=100
		   Set getImg = Jpeg
	   Set Jpeg = Nothing
	 end if
  end function


  '***********************************
  '   在指定图片上贴图
  '***********************************
  public function imgMark(path,imgObj,pX,pY)
     on error resume next
	 if reSizeImg.imgIsExits(path) then
	   Set iJpeg = server.CreateObject("Persits.Jpeg")
		   iJpeg.open path
		   iJpeg.DrawImage pX, pY, imgObj
		   iJpeg.Quality=100
		   iJpeg.save path
	   Set iJpeg = Nothing
	 end if
  end function
  
  
  
  '***********************************
  '   生成拼图
  '   * paths 一组图片路径
  '   * width 生成图片保存宽度
  '   * row   图片拼接列
  '   * path  拼图保存位置
  '***********************************
  public function imgSplice(paths,width,row,path)
	 on error resume next
	 
	 '// 初始化数据
	 imgSplice = ""
	 if text.isNum(width) = false then width = 600
	 if text.isNum(row) = false then row = 3
	 if path = "" then path = "up/fck/image_small/"
	 imgNull = "./up/designSplice_bg.jpg"
	 
	 '// 判断是否安装组件
	 if jpegInstall then
		 dim jWidth,jHeight,jPath,jMd5,jName,nPath,iPath
		 for each pss in paths
			jMd5 = jMd5 & pss
		 next
		 jMd5 = md5( jMd5 & width & row )
		 
		 '新图片相对地址,绝对路径
		 nPath = path&width&"@"&row&"@"&jMd5&".jpg"
		 iPath = server.MapPath("./"&nPath)
		 
		 dim Jpeg
		 '判断合成文件是否已经存在
		 if file2.isExits(iPath)=false then
			 dim imgObj,jpegNum,jpegW
				 jpegNum= ubound(paths)
				 jpegW  = int(width/row)
				 
			 dim itemNum,boxRow
				 itemNum= jpegNum-1
				 boxRow = row-1
				 
			 dim jpegArr,jpegObjArr,rowArr
			 redim jpegArr(itemNum)
			 redim jpegObjArr(itemNum)
			 redim rowArr(boxRow)
  
			 '拼接图片初始化参数
			 for jNum = 0 to itemNum
			   '以指定宽度获取拼图块
			   set imgObj = getImg( server.MapPath( paths(jNum) ) , jpegW )
				   '记录当前状态下的各列情况
				   set jpegObjArr(jNum) = imgObj
				   jpegArr(jNum) = rowArr
				   '累加当前列的高度
				   jRow = minIndex(rowArr)
				   rowArr(jRow) = int(rowArr(jRow)) + int(imgObj.height)
			   set imgObj = nothing
			 next
			 
			 '获取最终图片的高度
			 maxHeight = int(maxVal(rowArr))
			 
			 '创建图片
			 Set Jpeg = server.CreateObject("Persits.Jpeg")
				 Jpeg.open server.MapPath(imgNull)
				 Jpeg.Width  = width
				 Jpeg.Height = maxHeight
				 Jpeg.Quality=100
				 Jpeg.save iPath
			 Set Jpeg = Nothing

			 '拼接图片
			 for jNum = 0 to itemNum
			   '以指定宽度获取拼图块
			   rowArrItem = jpegArr(jNum)
			   jRow= minIndex(rowArrItem)
			   pX  = int(jRow * jpegW)
			   pY  = int(rowArrItem(jRow))
			   '将拼图块拼接在新图上
			   set imgObj = jpegObjArr(jNum)
			   IG = imgMark(iPath,imgObj,pX,pY)
			   set imgObj = nothing
			 next
		 end if
		 imgSplice = nPath
		 if err<>0 then imgSplice = ""
		 if imgSplice<>"" then imgSplice = "<img src=""" & imgSplice &""" />"
	 end if
  end function
  
  
  
  
  '***********************************
  '   获取数组最值
  '***********************************
  public function minIndex(arr)
	  minIndex = arrMinMax(arr,"index","min")
  end function
  
  public function minVal(arr)
	  minVal = arrMinMax(arr,"val","min")
  end function
  
  public function maxIndex(arr)
	  maxIndex = arrMinMax(arr,"index","max")
  end function
  
  public function maxVal(arr)
	  maxVal = arrMinMax(arr,"val","max")
  end function
  
  public function arrMinMax(arr,key,T)
    on error resume next
	dim arrNum,arrVal,arrIndex
	arrVal = 0
	arrIndex = 0
	arrNum = ubound(arr)
	if arrNum>0 then
	  arrVal = arr(0)
	  arrIndex = 0
	  for arrI=0 to arrNum
		 if T="min" then
			 if cint(arr(arrI))<cint(arrVal) then
				arrVal = arr(arrI)
				arrIndex = arrI
			 end if
		 else
			 if cint(arr(arrI))>cint(arrVal) then
				arrVal = arr(arrI)
				arrIndex = arrI
			 end if
		 end if
	  next
	end if
	if key="val" then
	   arrMinMax = arrVal
	else
	   arrMinMax = arrIndex
	end if
  end function
  
  
  
  
end Class



'*********************
'<><>< 实例化对象 ><><>
'*********************
set reSizeImg = new reSizeImgClass
%>