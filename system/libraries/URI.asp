<%
'*********************
'<><>< URL类 ><><>
'*********************

Class uriClass


  '***********************************
  '   类初始化
  '***********************************
  Private Sub Class_Initialize
  
  End Sub
  
  

  '***********************************
  '   获取域名
  '***********************************
  Public function getSite(str)
    str = lcase(str)
	str = replace(str,"http://","")
	str = replace(str,"https://","")
	str = replace(str,"www.","")
	str = replace(str,"/","")
	getSite = str
  end function
  
  
  
  '***********************************
  '   301 域名跳转
  '   调用示例:call autoRedirect("www","???.com")
  '***********************************
  Public sub autoRedirect(www,site)
	dim ScriptAddress,servername,qs,url
	servername = cstr(request.serverVariables("Server_Name"))
	if lcase(site) = lcase(servername) then
	   ScriptAddress = CStr(Request.ServerVariables("SCRIPT_NAME"))
	   url = "http://" & www & "." & site & ScriptAddress
	   qs = Request.QueryString
	   if qs<>"" then url = url & "?" & qs
	   response.status = "301 Moved Permanently" 
	   response.addHeader "Location", url 
	   response.end()
	end if
  End sub
 
 
 
  '***********************************
  '   用于返回带参数的Url值
  '   调用示例:call ReUrl("value=5&page=34")
  '***********************************
  Public function reUrl(keys)
	  dim UrlStr,UrlKey,Uqs,ul,uls,rekey,revalue,key,tempKey,tempVal
		  keys = lcase(keys)
		  UrlStr = lcase(request.QueryString)
		  if keys<>"" then
		     if UrlStr<>"" then UrlStr = "&" & UrlStr
			 '-------------------------------------
			 key = split(keys,"&")
			 for each ul in key
				 uls     = split(ul,"=")
				 rekey   = uls(0)
				 revalue = ul
				 tempKey = "&" & rekey & "="
				 tempVal = tempKey & request.QueryString(rekey)
				 if instr(UrlStr,tempKey)>0 and lcase(uls(1)) <> "null" then
					UrlStr = replace(UrlStr, tempVal , "&"&revalue)
				 elseif instr(UrlStr,tempKey)>0 and lcase(uls(1)) = "null" then
					UrlStr = replace(UrlStr, tempVal ,"")
				 elseif instr(UrlStr,tempKey)<=0 and lcase(uls(1)) <> "null" then
					if UrlStr<>"" then Uqs = "&"
					UrlStr = UrlStr & Uqs & revalue
				 end if
			 next
			 '-------------------------------------
			 if UrlStr<>"" then
			    UrlStr = replace("?" & UrlStr, "&&" ,"&")
				UrlStr = replace(UrlStr, "?&" ,"?")
			 end if
		  end if
	  reUrl = UrlStr
  end function
    
  

  '***********************************
  '   Url编码转换 
  '***********************************
  Public function encodeUrl(str)
	  encodeUrl = server.URLEncode(str)
  end function
  
  Public function uncodeUrl(str)
	  encodeUrl = server.URLEncode(str)
  end function
 
 
End Class



'*********************
'<><>< 实例化对象 ><><>
'*********************
set URI = New uriClass
%>
