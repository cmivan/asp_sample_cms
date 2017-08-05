<%
'
' 输出消息
' 
' @access: public
' @author: mk.zgc
' @param : string，str,提示消息
' @return: string
' @eq    : echo("提示消息"); 
' 
function echo(str)
    response.Write(str)
	response.Flush()
end function

'
' 提交表单返回信息提示
' 
' @access: public
' @author: mk.zgc
' @param : string，key,键; str,提示消息
' @return: string
' @eq    : json_form("键","提示消息"); 
' 
function json_form(key,str)
	echo "{""cmd"":"""&key&""",""info"":""" &str& """}"
	response.End()
end function


' 提交表单返回的错误信息提示
function json_form_no(str)
    json_form "n",str
end function


' 提交表单返回的正确信息提示
function json_form_yes(str)
    json_form "y",str
end function


'弹出消息窗口
function json_form_alt(str)
    json_form "alt",str
end function
%>