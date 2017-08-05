<%
function backData(r_db_num,r_tb_num,r_qq)

    dim z,dbase,dbtable,Group,sql
	
	z = (r_db_num-1)*100
	dbase   = s_db & r_db_num
	dbtable = s_tb &(r_tb_num+z)
	Group = dbase & sqlpower & dbtable
	
	sql = "select top 99999999 "&Group&".QunNum from "&Group&" where "&Group&".QQNum="&r_qq&" order by "&Group&".QunNum asc"
	
	dim backTemp
	set rs = sqlconn.execute(sql)
		do while not rs.eof
		   if backTemp<>"" then backTemp = backTemp & ","
		   backTemp = backTemp & rs("QunNum")
		rs.movenext
		loop
	set rs = nothing
	
	'//关闭数据库
	sqlconn.close
	set sqlconn=nothing
	
	
	backData = backTemp

end function


'//更新session
function update_session()
	u_qq = session("qq")
	u_db_num = session("db_num_"&qq)
	u_tb_num = session("tb_num_"&qq)
	if (cint(u_db_num)=cint(s_db_num)) and (cint(u_tb_num)=cint(s_tb_num)) then
	   update_session = false
	else
	   if cint(u_tb_num)<cint(s_tb_num) then
	       u_tb_num = u_tb_num+1
		   session("tb_num_"&qq) = u_tb_num
	   else
		   if cint(u_db_num)<cint(s_db_num) then 
			  u_db_num = u_db_num+1
			  session("tb_num_"&qq) = 1
			  session("db_num_"&qq) = u_db_num
		   end if
	   end if
	   update_session = true
	end if
end function





sub echo(str)
    response.Write(str)
	response.Flush()
end sub

sub echo_json(str)
    qq = session("qq")
	db_num = session("db_num_"&qq)
	tb_num = session("tb_num_"&qq)
    response.Write("{""cmd"":""yes"",""info"":"""&str&""",""db_num"":"""&db_num&""",""tb_num"":"""&tb_num&"""}")
	response.End()
end sub

function selected(key1,key2)
    if cstr(key1)=cstr(key2) then
       selected = " selected=""selected"""
	end if
end function

function isNum(str)
    if cstr(str)="" or isnumeric(str)=false then
	   isNum = false
	else
	   isNum = true
	end if
end function


function valOk()
    qq = session("qq")
	db_num = session("db_num_"&qq)
	tb_num = session("tb_num_"&qq)
	valOk = false
	if isNum(db_num) and  isNum(tb_num) and  isNum(qq) then
	   valOk = true
	end if
end function



'//写入数据
function inserQQqun(qq,qun)
    inserQQqun = false
	if text.isNum(qq) and text.isNum(qun) then
	   set inserRs = conn.execute("select * from QqQun as q left join QqList as qql on q.QqId=qql.id left join QunList as qunl on q.QunId=qunl.id where qql.QqNum="&qq&" and qunl.QunNum="&qun)
	       if not inserRs.eof then
		      inserQQqun = false
		   else

		   end if
	   set inserRs = nothing
	end if
end function
%>