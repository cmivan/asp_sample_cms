<script type="text/javascript" language="javascript"> 
var ColorHex=new Array('00','33','66','99','CC','FF') 
var SpColorHex=new Array('FF0000','00FF00','0000FF','FFFF00','00FFFF','FF00FF') 
var current=null 
function initcolor(evt) 
{ 
var colorTable='' 
for (i=0;i<2;i++) 
{ 
for (j=0;j<6;j++) 
{ 
colorTable=colorTable+'<tr height=15>' 
colorTable=colorTable+'<td width=15 style="background-color:#000000">' 
if (i==0){ 
colorTable=colorTable+'<td width=15 style="cursor:pointer;background-color:#'+ColorHex[j]+ColorHex[j]+ColorHex[j]+'" onclick="doclick(this.style.backgroundColor)">'} 
else{ 
colorTable=colorTable+'<td width=15 style="cursor:pointer;background-color:#'+SpColorHex[j]+'" onclick="doclick(this.style.backgroundColor)">'} 
colorTable=colorTable+'<td width=15 style="background-color:#000000">' 
for (k=0;k<3;k++) 
{ 
for (l=0;l<6;l++) 
{ 
colorTable=colorTable+'<td width=15 style="cursor:pointer;background-color:#'+ColorHex[k+i*3]+ColorHex[l]+ColorHex[j]+'" onclick="doclick(this.style.backgroundColor)">' 
} 
} 
} 
} 
colorTable='<table border="0" cellspacing="0" cellpadding="0" style="border:1px #000000 solid;border-bottom:none;border-collapse: collapse;width:80px;" bordercolor="000000">' 
+'<tr height=20><td colspan=21 bgcolor=#ffffff style="font:12px tahoma;padding-left:2px;">' 
+'<span style="float:right;padding-right:3px;cursor:pointer;" onclick="colorclose()">×关闭</span>' 
+'</td></table>' 
+'<table width="100%" border="1" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="000000" style="cursor:pointer;">' 
+colorTable+'</table>'; 
document.getElementById("colorpane").innerHTML=colorTable; 
var current_x = document.getElementById("inputcolor").offsetLeft; 
var current_y = document.getElementById("inputcolor").offsetTop; 
//alert(current_x + "-" + current_y) 
document.getElementById("colorpane").style.left = current_x + "px"; 
document.getElementById("colorpane").style.top = current_y + "px"; 
} 
function doclick(obj){ 
//alert(obj); 
document.getElementById("inputcolor").value=obj;   //返回颜色值
document.getElementById("inputcolor").style.background=obj;   //返回颜色
//document.getElementById("inputcolor").style.color=obj;

document.getElementById("colorpane").style.display = "none"; 
} 
function colorclose(){ 
document.getElementById("colorpane").style.display = "none"; 
} 
function coloropen(){ 
document.getElementById("colorpane").style.display = ""; 
} 
window.onload = initcolor; 
</script>


<table border="0" align="left" cellpadding="1" cellspacing="1" class="forum2">
<tr class="forumRaw2">
<td width="28" align="center"><input name="strong" type="checkbox" id="strong" value="1" <%if strong=1 then%>checked="checked"<%end if%> /></td>
<td width="40" align="left">加粗</td>
<td width="40" align="right" style="position:relative">颜色

<div id="colorpane" style="position:absolute;z-index:999;display:none;"></div>
</td>
<td width="60" align="center" style="position:relative"><input name="color" type="text" id="inputcolor" style="width:80px;background-color:<%=color%>" onclick="coloropen(event)" value="<%=color%>"/></td>
<td width="40" align="right">
字号
</td>
<td align="right"><select name="size" id="size">
  <option value="">默认大小</option>
  <%
'if sizes="" or isnumeric(sizes)=false then sizes=13
for i=12 to 20
%>
  <option value="<%=i%>" <%if cstr(i)=cstr(sizes) and sizes<>"" then%>selected="selected"<%end if%>><%=i%>px</option>
  <%next%>
</select></td>
<td width="90"><span style="color: #FF0000">(请适当设置)</span></td>
</tr></table>