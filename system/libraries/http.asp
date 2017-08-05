<%
'*********************
'<><>< URL类 ><><>
'*********************

Class httpClass

  Public charSets '--编码

  '***********************************
  '   类初始化
  '***********************************
  private sub Class_Initialize
     charSets = "gb2312"
  end sub
    

  '***********************************
  '   获取网络数据函数 
  '***********************************
  Public function httpPage(url) 
      dim tempBody
	  tempBody = pageGet(url)
	  httpPage = bytesToBstr(tempBody,"utf-8")
  End Function
  
  
  
  '***********************************
  '   保存网络数据到文件 
  '***********************************
  Public function httpSave(url,path) 
	on error resume next
	dim tempBody
	tempBody = httpPage(url)
	path = server.MapPath(path)
	Set objs = Server.CreateObject("ADODB.Stream")
		With objs
		.Open
		.Charset = "utf-8"
		.Position = objs.Size
		.WriteText = tempBody
		.SaveToFile path,2 
		.Close
		End With
	Set objs = Nothing
	if err.number = 0 then
	   httpSave = true
	else
	   err.clear
	   httpSave = false
	end if
  End Function
  
  
  
  '***********************************
  '   获取网络数据函数GET方式 
  '***********************************
  Public function pageGet(url) 
     on error resume next
     dim xmlpage  
	  
	 if IsNull(url)=True or Len(url)<18 Then
		pageGet = false
		exit Function
	 end If
	 
	 set xmlpage = server.createobject("MSXML2.XMLHTTP")
		 xmlpage.open "GET", url, false,"", ""
		 xmlpage.Send()
		 If xmlpage.Readystate<>4 then
		    Set xmlpage = Nothing
		    pageGet = false
		    Exit function
		 end if
		 pageGet = xmlpage.responseBody
	 set xmlpage = Nothing
	 if err.number<>0 then err.Clear
  End Function 
  
  
  
  '***********************************
  '   获取网络数据函数POST方式 
  '***********************************
  Public function pagePost(url)
     on error resume next
	 dim xmlpage,xmlpost
	 
	 if isNull(url)=True or Len(url)<18 Then
		pagePost = false
		exit Function
	 end If
	 
	 set xmlpage = server.createobject("MSXML2.XMLHTTP")
	     xmlpost = ""
		 'xmlpost = server.URLEncode(keyword)
		 'xmlpost = "keyword=" & xmlpost
		 xmlpage.open "POST", url, false
		 xmlpage.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		 xmlpage.Send( xmlpost )
		 If xmlpage.Readystate<>4 then
		    Set xmlpage = Nothing
		    pagePost = false
		    Exit function
		 end if
		 pagePost = xmlpage.responseBody
		 'httpPage = replace(replace(httpPage , vbCr,""),vbLf,"")
	 set xmlpage = Nothing
	 if err.number<>0 then err.Clear
  End function

  
  
  '***********************************
  '   获取网络数据函数 
  '***********************************
  Public function bytesToBstr(body,chars)         
	  dim obj
	  set obj = Server.CreateObject("adodb.stream")
	      obj.Type = 1
	      obj.Mode =3
	      obj.Open
	      obj.Write body
	      obj.Position = 0
	      obj.Type = 2
	      obj.Charset = chars
	      bytesToBstr = obj.ReadText 
	      obj.Close
	  set obj = nothing
  End Function
 
 
End Class



'*********************
'<><>< 实例化对象 ><><>
'*********************
set http = New httpClass
%>
