<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head><title><%=c_title%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="<%=rpath%>public/style/style.css"/>
<link rel="stylesheet" type="text/css" href="<%=rpath%>public/assets/css/bootstrap.ie6.helper.css"/>
<link rel="stylesheet" type="text/css" href="<%=rpath%>public/assets/css/bootstrap.css"/>
<!--<script language="javascript" src="<%=Rpath%>public/style/jquery.1.4.js"></script>-->
<script language="javascript" src="<%=Rpath%>public/assets/js/jquery.js"></script>
<script language="javascript" src="<%=Rpath%>public/style/jquery.edit.style.js"></script>
<script language="javascript" src="<%=Rpath%>public/style/png24.js"></script>
<script language="javascript">
//文件上传,弹出的窗口 js
function showUploadDialog(otype, olink, othumbnail){
var arr = showModalDialog("<%=Rpath%>public/editBox/fckupLoad/dialog/i_upload.htm?style=popup&type="+otype+"&link="+olink+"&thumbnail="+othumbnail, window, "dialogWidth:0px;dialogHeight:0px;help:no;scroll:no;status:no");
}
//防止刷新的
//function document.onkeydown(){
//	if(event.keyCode==116){
//		window.parent.frames["main"].location.reload();
//		event.keyCode = 0;
//		return false;
//		}
//	}
</script>
</head>
