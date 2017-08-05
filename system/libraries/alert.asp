<%
'*********************
'<><>< 提示消息处理类 ><><>
'*********************

Class alertClass
  
  Private alertHead
  
  
  '***********************************
  '   类初始化
  '***********************************
  private sub Class_Initialize
  
	alertHead = alertHead & "<meta http-equiv=Content-Type content=text/html; charset=utf-8 />"
	alertHead = alertHead & "<link href='"&Rpath&"public/style/style.css' rel='stylesheet' type='text/css' />"
    alertHead = alertHead & "<link href='"&Rpath&"public/assets/css/bootstrap.css' rel='stylesheet' type='text/css' />"
	
  End Sub
  
  'meta自动跳转到指定页面
  Public function reFresh(str,url)
       dim back
       back = back & alertHead
       back = back & "<meta http-equiv=refresh content=2;url='"&url&"' >"
	   back = back & "<body style=""overflow:hidden;"">"
       back = back & "<br><TABLE align=center cellpadding=0 cellspacing=10 class=""alert alert-error""><tr><td>"
       back = back & str
       back = back & "</td></tr></table><br><br></body>"
	   reFresh = true
	   response.Write(back)
	   response.End()
  End function
  
  
  'js弹出提示，返回上一级
  Public Sub msgBack(str)
       back = alertHead & "<script language='javascript'>alert('"&str&"');history.back(1);</script>"
	   response.Write(back)
	   response.End()
  End Sub
  
  
  'js弹出提示，重新加载
  Public Sub msgReload(str)
       back = alertHead & "<script language='javascript'>alert('"&str&"');window.location.reload();</script>"
	   response.Write(back)
	   response.End()
  End Sub
  
  
  'js弹出提示，返回指定页面
  Public Sub msgTo(str,url)
       back = alertHead & "<script language='javascript'>alert('"&str&"');window.location.href='"&url&"';</script>"
	   response.Write(back)
	   response.End()
  End Sub
      
 
 
End Class

  
'*********************
'<><>< 实例化对象 ><><>
'*********************
set alert = New alertClass
%>
