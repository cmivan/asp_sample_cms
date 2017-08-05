<%
'*********************
'<><>< 文本处理类 ><><>
'*********************

Class textClass


  '***********************************
  '   类初始化
  '***********************************
  private sub Class_Initialize
  
  end sub
  
'
' 输出消息
' 
' @access: public
' @author: mk.zgc
' @param : string，str,提示消息
' @return: string
' @eq    : echo("提示消息"); 
' 
  public function echo(str)
	  response.write(str)
	  response.flush()
  end function
  
  
  
  '***********************************
  '    返回是否选择项
  '***********************************
  public function getBool(num)
  	  if cstr(num)="1" then
	     getBool = 1
	  else
	     getBool = 0
	  end if
  end function
  
  
  
  '***********************************
  '    返回是否选择项
  '***********************************
  public function isNum(key)
      isNum = false
	  if key<>"" and isnumeric(key) then isNum = true
  end function
  
  
  
  '***********************************
  '    返回是否选择项
  '***********************************
  public function getNum(key)
      getNum = key
      if isNum(key)=false then getNum = 0
  end function

  
  '***********************************
  '    截取指定长度字符
  '***********************************
  public function noSql(key)
      key = replace(key,"'","")
	  key = replace(key,"<","")
	  key = replace(key,">","")
	  key = replace(key,"-","")
	  key = replace(key," ","")
	  key = replace(key,"%","")
	  noSql = key
  end function
  
  
  
  '***********************************
  '    过滤HTML字符
  '***********************************
  public function noHtml(html) 
	  Dim objRegExp, Match, Matches 
	  Set objRegExp = New Regexp 
		 objRegExp.IgnoreCase = true 
		 objRegExp.Global = true 
		 '取闭合的<> 
		 objRegExp.Pattern = "<.+?>" 
		 '进行匹配 
	  Set matches = objRegExp.Execute(html) 
		 '遍历匹配集合，并替换掉匹配的项目 
	  For Each Match in matches 
		 html = Replace(html,Match.Value,"") 
	  Next
		 noHtml = html 
	  Set objRegExp = Nothing 
  End Function
  
  
  
  '***********************************
  '    在HTML中找出图片生成数组
  '***********************************
  public function getImgSrc(html)
	  Dim objRegExp, Match, Matches ,newHTML ,newNum ,newImg
	  Set objRegExp = New Regexp 
		 objRegExp.IgnoreCase = True 
		 objRegExp.Global = True 
		 '取闭合的<> 
		 objRegExp.Pattern = "<img(.+?)src=""(.+?)""(.+?)>" 
		 '进行匹配 
	  Set Matches = objRegExp.Execute(html) 
		 '遍历匹配集合，并替换掉匹配的项目
	  newHTML = ""
	  newNum = 0
	  dim Mnum
	  Mnum = Matches.Count
	  redim Img(Mnum)
	  For Each Match in Matches
		 if newNum<=15 then
			'// 这里用到图片尺寸处理类
			Img(newNum) = Match.SubMatches(1)
			newNum = newNum + 1
		 end if
	  Next
	  getImgSrc = Img 
	  Set objRegExp = Nothing 
  end function
  
  
  
  '***********************************
  '    截取指定长度字符
  '***********************************
  public function strCut(txt,length)
	dim i
	i = 1
	y = 0
	txt = trim(txt)
	txt = noHtml(txt)
	for i=1 to len(txt)
		j = mid(txt,i,1)
		'汉字外的其他符号,如:!@#,数字,大小写英文字母
		if ascw(j)>=0 and ascw(j)<=127 then
		   y = y + 1
		else '汉字
		   y = y + 2
		end if
		if -int(-y) >= length then '截取长度
		   txt = left(txt,i) & "..."
		   exit for
		end if
	next
	strCut = txt
  end function
  
  
  
  '***********************************
  '    返回标题颜色
  '***********************************
  public function strColor(title,color)
	if color<>"" then
	   strColor = "<span style='color:"&color&"'>"&title&"</span>"
	else
	   strColor = title
	end if
  End function
  
  
  
  '***********************************
  '    返回关键字颜色，多用于搜索
  '***********************************
  public function keyColor(str,keyword)
      if keyword<>"" then
		  dim keywords
		  keywords = strColor(keyword,"#ff0000")
		  keyColor = replace(str,keyword,keywords)
	  else
	      keyColor = str
	  end if
  end function
  

 
End Class
  
  
'*********************
'<><>< 实例化对象 ><><>
'*********************
set text = New textClass
%>
