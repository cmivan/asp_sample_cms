<%

'### <><><><><><><><><><><><>
'### Time:2011-04-24 
'### For :sql ���ݵ�������access
'### By  :cmivan
'### Contact:619835864
'### <><><><><><><><><><><><>

 response.Buffer=true
 server.ScriptTimeout=3000
 on Error Resume Next

'//��������(GO)
dim sqlconn
set sqlconn=server.createobject("ADODB.connection") 
	sqlconn.open "server=(local);UID=sa;PWD=3389007;database=master;Provider=SQLOLEDB"

dim sqlpower
    sqlpower = ".dbo."
	
if err<>0 then
   response.Write("����ʧ��"&err.description)
   response.End()
end if



'//��ʼ������
dim s_db_num,s_tb_num
s_db_num = 10
s_tb_num = 100

dim s_db,s_tb
s_db = "GroupData"
s_tb = "Group"


'### <><><><><><><><><><><><>


'### <><><><><><><><><><><><>


qq = session("qq")
if isNum(qq) then
   db_num = session("db_num_"&qq)
   tb_num = session("tb_num_"&qq)
end if

%>

<style>
body{font-size:12px;}
#loading{background-color:#eee;border:1px #ccc solid;padding:3px;}
#loading_run{background-color:#C36;border:1px #fff solid;border-top:2px #666 solid;border-left:2px #666 solid;line-height:200%;font-weight:bold;color:#fff;text-align:right;width:100%;}
#loading_num{padding-right:10px;}
</style>
<script type="application/javascript" src="style/js/jquery1.7.js"></script>

<form action="" method="post">
<select name="db_num" id="db_num">
<option value="">ȫ��</option>
<%for i=1 to s_db_num%>
  <option value="<%=i%>"<%=selected(i,db_num)%>><%=s_db&i%></option>
<%next%>
</select>
<select name="tb_num" id="tb_num">
<option value="">ȫ��</option>
<%for i=1 to s_tb_num%>
  <option value="<%=i%>"<%=selected(i,tb_num)%>><%=i%></option>
<%next%>
</select>
<input name="qq" type="text" id="qq" value="<%=qq%>" />
<%if isNum(session("qq")) and go="" then%>
<input name="send" type="submit" id="send" value="��������" />
<%elseif go<>"" and valOk() then%>
<input name="send" type="submit" id="send" value="������...(��ͣ)" />
<%else%>
<input name="send" type="submit" id="send" value="����" />
<%end if%>
<input name="go" type="hidden" value="yes" />
</form>

<%
if qq="" or isNum(db_num)=false then
   echo("<p>��ʲô��</p>")
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
			alert("ȫ����ѯ���!");	
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

<!--����-->
<div id="loading">
<div id="loading_run"><div id="loading_num">100%</div></div>
</div>

<div id="loadback"></div>


<%end if%>
