<!--#include file="system/core/initialize.asp"-->
<!--#include file="system/libraries/articles.asp"-->
<%
  '//链接SQL(GO)
  dim sqlconn
  set sqlconn=server.createobject("ADODB.connection") 
	  sqlconn.open "server=(local);UID=sa;PWD=3389007;database=master;Provider=SQLOLEDB"
  
  dim sqlpower
	  sqlpower = ".dbo."
	  
  if err<>0 then
	 response.Write("链接失败:"&err.description)
	 response.End()
  end if
  
  
  
  
'//初始化参数
dim s_db_num,s_tb_num
s_db_num = 10
s_tb_num = 100

dim s_db,s_tb
s_db = "GroupData"
s_tb = "Group"


%>
<!--#include file="system/models/getqun.asp"-->
<%



qq = session("qq")
if isNum(qq) then
   db_num = session("db_num_"&qq)
   tb_num = session("tb_num_"&qq)
end if 

'//AJAX查询
go_ajax = request("go_ajax")
if go_ajax="yes" then
   if valOk() then
	  '//执行查询
	  backDB = backData(db_num,tb_num,qq)
	  '//更新session并判断是否已经完成
	  if update_session()=false then
	     response.End()
	  else
	     echo_json(backDB)
	  end if
   end if
   response.End()
end if



'//获取参数
go = request("go")
if go="yes" then

   db_num = request("db_num")
   tb_num = request("tb_num")
   if isNum(db_num)=false then db_num = 1
   if isNum(tb_num)=false then tb_num = 1
   if isNum(request("qq")) then
      qq = request("qq")
	  
	  session("qq") = qq
	  session("db_num_"&qq) = db_num
	  session("tb_num_"&qq) = tb_num
   end if 
   
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<link href="public/assets/css/bootstrap.css" rel="stylesheet" type="text/css" />
<link href="style/css/screen.css" rel="stylesheet" type="text/css" />
<script src="style/js/jquery1.7.js"></script>

<div class="form-search-box">
<div class="main_body_width">

<form action="" method="post" class="form-search">
<select name="db_num" id="db_num">
<option value="">全部</option>
<%for i=1 to s_db_num%>
  <option value="<%=i%>"<%=selected(i,db_num)%>><%=s_db&i%></option>
<%next%>
</select>
<select name="tb_num" id="tb_num">
<option value="">全部</option>
<%for i=1 to s_tb_num%>
  <option value="<%=i%>"<%=selected(i,tb_num)%>><%=i%></option>
<%next%>
</select>
<input name="qq" type="text" id="qq" value="<%=qq%>" />
<%if isNum(session("qq")) and go="" then%>
<button type="submit" class="btn"><span class="icon-search"></span> 继续搜索</button>
<%elseif go<>"" and valOk() then%>
<button type="submit" class="btn"><span class="icon-search"></span> 搜索中...(暂停)</button>
<%else%>
<button type="submit" class="btn"><span class="icon-search"></span> 搜索</button>
<%end if%>
<input name="go" type="hidden" value="yes" />
</form>

</div>
</div>

<%
if qq="" or isNum(db_num)=false then
   echo("<p>找什么？</p>")
   response.End()
end if
%>

<%if go<>"" and valOk() then%>

<script language="javascript">
function loading(num){
	$("#loading_num").text(num + "%");
	$("#loading_run").css({"width":num+"%"});
}
function setSel(obj,key){
	obj.find("option[value='"+key+"']").attr("selected","selected");
}
	
var s_db_num = <%=s_db_num%>;
var s_tb_num = <%=s_tb_num%>;
function getQun(){
	$.post('?',{"go_Ajax":"yes"},function(d){
		if(d.info!=""){
			$("#loadback").append("<li>" + d.info + "</li>");
		}
		var a = parseInt(d.db_num),b=parseInt(d.tb_num);
		
		if(a>=s_db_num&&b>=s_tb_num){
			alert("全部查询完成!");	
		}else{
			var z = (a-1)*100 + b;
			var s = (a+z)*100/(s_db_num*s_tb_num);
			loading(s);
			
			setSel($("#db_num"),a);
			setSel($("#tb_num"),b);
			
			getQun();
		}
	},"json");	
}
$(function(){ getQun(); });
</script>

<!--进度-->
<div class="progress progress-striped active">
  <div class="bar" style="width:0;" id="loading_run"><div id="loading_num">0%</div></div>
</div>

<div id="loadback"></div>


<%end if%>



