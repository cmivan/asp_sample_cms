<%
'*********************
'<><>< URL类 ><><>
'*********************

Class dateClass

  '***********************************
  '   类初始化
  '***********************************
  private sub Class_Initialize
  
  end sub
  
  '***********************************
  '    返回时间,多用于订单号
  '***********************************
  Public function times(T)
	  if T="" then
		 T = times( now() )
	  else
		 T = year(T) & month(T) & day(T) & hour(T) & minute(T) & second(T)
	  end if
	  times = T
  end function
  
  
  Public function ymd(T)
	  if T="" then
		 T = ymd( now() )
	  else
		 T = year(T) &"-"& month(T) &"-"& day(T)
	  end if
	  ymd = T
  end function
 
 
End Class



'*********************
'<><>< 实例化对象 ><><>
'*********************
set date2 = New dateClass
%>