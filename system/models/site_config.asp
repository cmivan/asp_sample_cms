<%

'// 获取站点配置信息
function SC(key,filed)
    dim SCrs
	set SCrs = New data
		SCrs.where "key","'"&key&"'"
		SCrs.open "site_config",1,false
		if not SCrs.eof then SC = SCrs.rs(filed)
		SCrs.close
	set SCrs = nothing
end function

'//获取导航信息

%>