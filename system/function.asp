<%
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