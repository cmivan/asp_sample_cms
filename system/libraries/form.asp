<%
'*********************
'<><>< 文件处理类 ><><>
'*********************

Class formClass

  '***********************************
  '   类初始化
  '***********************************
  private sub Class_Initialize
  
  end sub
  

  '***********************************
  '   生成文件。目录 
  '***********************************
  Public sub web_create(s_name,s_centent,s_type)
  on error resume next
	  select case s_type
	  
			 case "folder"                          '文件夹不存在则生成
				  set fso=createobject("scripting.filesystemobject")
					  if fso.folderexists(server.mappath(s_name))= false then  
						 fso.createfolder(server.mappath(s_name))
					  end if                            
				  set fso=nothing 
		 
			 case "file"                            '文件不存在则生成文件
				 if s_name<>"" then
				  set fso=createobject("scripting.filesystemobject")
				  set savefile=fso.opentextfile(server.mappath("editbox\file\"&s_name),2,true)
					  savefile.writeline(s_centent)
					  response.write("已生成文件")
				  set savefile=nothing
				  set fso=nothing	
				  else
					  response.write("缺少文件名~")
				 end if	
	  end select
  end sub
    

  '***********************************
  '   读取文件内容 
  '***********************************
  Public function openfile(path)
	  fileurl = server.MapPath(path)
	  set fso = server.CreateObject("scripting.filesystemobject") '定义FSO
	  set mofile=fso.opentextfile(fileurl,1) '以读的方式打开文件
	  mo_top=mofile.readall() '读取全部内容
	  mofile.close
	  openfile = mo_top
  end function
  
 
 
End Class



'*********************
'<><>< 实例化对象 ><><>
'*********************
set form2 = New formClass
%>
