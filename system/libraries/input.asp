<%
'*********************
'<><>< 操作数据库类 ><><>
'*********************

Class inputClass

  '***********************************
  '   类初始化
  '***********************************
  private sub Class_Initialize
  
  end sub
 
 'getnum /////////////
 Public function getnum(key)
      dim val
      val = request.QueryString(key)
      if val<>"" and text.isNum(val) then getnum = val
 End function
 
 'postnum /////////////
 Public function postnum(key)
      dim val
      val = request.Form(key)
      if val<>"" and text.isNum(val) then getnum = val
 End function
 

 
End Class




'*********************
'<><>< 实例化对象 ><><>
'*********************
set input = New inputClass
%>
