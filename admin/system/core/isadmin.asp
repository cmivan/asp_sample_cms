<%
'### 容错模式
 on error resume next
 
'### 验证是否已登陆
 if errTip="" then errTip = "登陆超时或未登录,请重新登陆后再进行操作!"
 if session("cmcms.admin")="" then call alert.reFresh(errTip,Rpath&Rpath&"../")
 
'### 注销登陆
 if request.QueryString("login")="out" then
    session.Abandon()
    call alert.reFresh("注销成功!",Rpath&"../")
 end if
%>