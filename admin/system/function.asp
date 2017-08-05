<%
function dataConn(rootpath,dbpath)
    on error resume next
    dim conn,connstr,rconnstr
	    rconnstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
		
	'//第1次连接
    set conn = server.CreateObject("ADODB.Connection") 
        connstr = rconnstr & server.mappath(rootpath & dbpath)
        conn.open connstr
        If err then
           err.clear
           set conn = nothing
	   
		   '//第2次连接
		   rootpath = rootpath & "../"
		   set conn = server.CreateObject("ADODB.Connection") 
			   connstr = rconnstr & server.mappath(rootpath & dbpath)
			   conn.open connstr
			   If err then
				  err.clear
				  set conn = nothing
				  
				  '//第3次连接
				  rootpath = rootpath & "../"
				  set conn = server.CreateObject("ADODB.Connection") 
					  connstr = rconnstr & server.mappath(rootpath & dbpath)
					  conn.open connstr
					  If err then
						 err.clear
						 set conn = nothing
					  
						 '//第4次连接
						 rootpath = rootpath & "../"
						 set conn = server.CreateObject("ADODB.Connection") 
							 connstr = rconnstr & server.mappath(rootpath & dbpath)
							 conn.open connstr
							 If err then
								err.clear
								set conn = nothing
								response.write(dberr)
								response.end()
							 end If
					  end If
			   end If 
		end If
		'//记录路径
		rpath = rootpath
		set dataConn = conn
end function

'× --------------------------------------
'× -------  用于返回热门最新推荐审核等  -----
'× --------------------------------------
  function getBool(num)
  	  if num="1" then
	     getBool =1
	  else
	     getBool =0
	  end if
  end function
'######用于后台多语言菜单#######################
function OnUrl(lag)
    onStr=lcase(Request.ServerVariables("URL"))
	if instr(onStr,"/"&lcase(lag)&"/")<>0 then
	   onStyle="class=on"
	end if
	OnUrl=onStyle
end function
'///////////////////////////////////////////////
function GetUrl(lag)     
On Error Resume Next
dim strTemp      
    strTemp=lcase(Request.ServerVariables("SCRIPT_NAME"))
    strTemp=replace(strTemp,"/meun.asp","../../../system_admin.asp")
   for each urls in urlStrs
       urls=lcase(urls)
       urlss=split(urls,",")
       strTemp=replace(strTemp,"/"&urlss(1)&"/","/"&lag&"/")
   next 
   GetUrl=strTemp
end function


'判断是否是数字
function isnum(num)
  if num="" or isnumeric(num)=false then
     isnum=false
  else
     isnum=true
  end if
end function

'错误提示信息返回
function errtips(str)
  response.write("<script>alert('"&str&"');history.back(1);</script>")
  response.end()
end function



'× --------------------------------------
'× ----------  返回提示信息 ---------------
'× --------------------------------------
  function backPage(backStr,backUrl,backType)
    back =""
	back =back&"<meta http-equiv=Content-Type content=text/html; charset=utf-8 />"
	back =back&"<link href='"&Rpath&"images/style/style.css' rel='stylesheet' type='text/css' />"
	
	if backType=0 then
	    'meta自动跳转到指定页面
       back =back&"<meta http-equiv=refresh content=2;url='"&backUrl&"' >"
	   back =back&"<body style=""overflow:hidden;"">"
       back =back&"<br><TABLE align=center cellpadding=0 cellspacing=10 class=forum1><tr><td>"
       back =back&"<table width=100% align=center cellpadding=3 cellspacing=1 class=forum2>"
       back =back&"<tr class=forumRow><td class=""noInfo"">"
       back =back&backStr
       back =back&"</tr></table></td></tr></table><br><br>"
	elseif backType=1 then
	    'js弹出提示，返回指定页面
       back =back&"<script language='javascript'>alert('"&backStr&"');window.location.href='"&backUrl&"';</script>"
	elseif backType=2 then
	    'js弹出提示，返回上一级
       back =back&"<script language='javascript'>alert('"&backStr&"');history.back(1);</script>"
	elseif backType=3 then
	    'js弹出提示，返回指定页
       back =back&"<script language='javascript'>window.location.href='"&backUrl&"';</script>"
	end if
	response.Write(back)
	response.End()
 end function


'× --------------------------------------
'× ----------  后台编辑器 ---------------
'× --------------------------------------
 function edit_box(eb_key,eb_content,eb_type)
     '//获取管理员设置
	 if cint(eb_type)=0 then
	    eb_db="box"
	 else
	    eb_db="box_1"
	 end if
	 uid=session("user_id")
	 if uid<>"" and isnumeric(uid) then
	    set eb_rs=sysconn.execute("select E."&eb_db&" from sys_admin A left join sys_ebox E on A.editbox=E.id where A.id="&uid)
		    edit_box=cstr(eb_rs(eb_db))
			edit_box=replace(edit_box,"{key}",eb_key)
			edit_box=replace(edit_box,"{value}",eb_content)
			response.Write(edit_box)
		set eb_rs=nothing
	 end if
 end function
 
 
 
 function getproductmoreVal(product_id,parent_id)
      if product_id<>"" and isnumeric(product_id) and parent_id<>"" and isnumeric(parent_id) then
		  set get_one_rs=conn.execute("select note from product_more where product_id="&int(product_id)&" and parent_id="&int(parent_id))
			  if not get_one_rs.eof then getproductmoreVal = get_one_rs("note")
		  set get_one_rs=nothing
      end if
 end function
 
 
 
Function CheckTable(myTable)
	'列出数据库中的所有表
	dim Crs,getTableName
	set Crs=conn.openSchema(20) 
	Crs.movefirst 
	do Until Crs.EOF 
	if Crs("TABLE_TYPE")="TABLE" then 
	   getTableName=getTableName+Crs("TABLE_NAME")&","
	end if 
	Crs.movenext
	loop
	Crs.close
	set Crs=nothing
	'判断数据库中是否存在此表（下面两行代码使待比较表和指定表前后都加上豆号，以精确比较）
	dim getTableName2,myTable2
	getTableName2=","&getTableName '此字符串后面已经有豆号
	myTable2=","&myTable&","
	If instr(getTableName2,myTable2)<>0 Then
	   CheckTable=1 '存在
	else
	   CheckTable=0 '不存在
	end if
End Function














'//用于从内容中解析出参数,读取指定的图片目录内容
function view_content(content)
    on error resume next
	dim picroot,picpath
	picroot = "up/fck/image/"
    set regex = new regexp
		regex.ignorecase = true
		regex.global = true
		regex.pattern = "{imglist:(.*?)\|(\d+)}"
		set matches = regex.Execute(content)
			for each items in matches
			   p1 = items.submatches(0)
			   p2 = items.submatches(1)
			   '解析出参数，并调用函数
			   if p1<>"" and p2<>"" and isnumeric(p2) then
			      '处理路径
			      picpath = picroot & p1
				  picpath = replace(picpath,"\","/")
				  picpath = replace(picpath,"//","/")
			      '获取并替换数据
			      oldtext = "{imglist:" & p1 & "|" & p2 & "}"
				  newtext = pic_content(picpath,p2)
				  content = replace(content,oldtext,newtext)
			   end if
			next
		set matches = nothing
    set regex = nothing
	view_content = content '//返回数据
end function


'//读取指定目录的图片目录内容
function pic_content(dirPath,T)
	on error resume next
	if right(dirPath,1)<>"/" then dirPath = dirPath & "/"
	
	dim fixs,thisfix,picpath,allretrun
	fixs = "|gif|jpg|png|bmp|"

	picpath = server.MapPath(dirPath)
	if right(picpath,1)<>"\" then picpath = picpath & "\"
	
	set FSO = CreateObject("scripting.filesystemobject") 
	set f = FSO.GetFolder(picpath) 
	set fs = f.files 
	for each fileN in fs
	    thisfile = fileN.name
	    thisfix = lcase(right(thisfile,3))
		thisname = left(thisfile,1)
		'//详细内容中排除大小图
	    if instr(fixs,thisfix)>0 and thisname<>"大" and thisname<>"小" then
		   if T=0 then
		      allretrun = allretrun & "<p><img src='"&dirPath&thisfile&"'/></p>"
		   else
		      allretrun = allretrun & "<p><img src='"&dirPath&thisfile&"'/><br>"&left(thisfile,len(thisfile)-4)&"</p>"
		   end if
		end if
	next
	
	'如果子目录也要的话就把下弄注释去除 
	set fsubfolers = f.SubFolders 
	    for Each dir in fsubfolers 
	        DN = dirPath & dir.name
	        allretrun = allretrun & pic_content(DN,T)
	    next
	set fsubfolers = nothing

	set FSO = nothing
	pic_content = allretrun
end function




'######### 获取上一篇或下一篇 #########
'#### obj = 表，id 值
function this_prev(obj,id,str)
    if str="" then str="上一页"
    dim nav
    nav = ""
    if id<>"" and isnumeric(id) then
       set rsP = conn.execute("select top 1 * from "&obj&" where id>"&id&" order by order_id desc,id asc")
           if not rsP.eof then nav = "<a href='"&obj&"_view.asp?id="&rsP("id")&"'>"&str&"</a>"
       set rsP = nothing
    end if
    this_prev = nav
end function

function this_next(obj,id,str)
    if str="" then str="下一页"
    dim nav
    nav = ""
    if id<>"" and isnumeric(id) then
       set rsP = conn.execute("select top 1 * from "&obj&" where id<"&id&" order by order_id asc,id desc")
           if not rsP.eof then nav = "<a href='"&obj&"_view.asp?id="&rsP("id")&"'>"&str&"</a>"
       set rsP = nothing
    end if
    this_next = nav
end function



%>