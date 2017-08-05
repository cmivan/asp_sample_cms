<%
Option Explicit
Response.buffer=true
session.Timeout = 20
validcode("")

Function validcode(key)
	Response.Expires = -1
	Response.AddHeader "Pragma","no-cache"
	Response.AddHeader "cache-ctrol","no-cache"
	On Error Resume Next
	Dim codenum,i,j
	Dim Ados,Ados1
	Randomize timer
	codenum = cint(8999*Rnd+1000)
	
	if key="" then key="code.num"
	Session(key) = codenum
	
	Dim zimg(4),codestr
	codestr = cstr(codenum)
	
	for i=0 To 3
		zimg(i)=cint(mid(codestr,i+1,1))
	next
	
	Dim Pos
	Set Ados=Server.CreateObject("Adodb.Stream")
	Ados.Mode=3
	Ados.Type=1
	Ados.Open
	Set Ados1=Server.CreateObject("Adodb.Stream")
	Ados1.Mode=3
	Ados1.Type=1
	Ados1.Open
	Ados.LoadFromFile(Server.mappath("body.Fix"))
	Ados1.write Ados.read(1280)
	For i=0 To 3
		Ados.Position=(9-zimg(i))*320
		Ados1.Position=i*320
		Ados1.write ados.read(320)
	Next	
	Ados.LoadFromFile(Server.mappath("head.fix"))
	Pos=lenb(Ados.read())
	Ados.Position=Pos
	
	For i=0 To 9 Step 1
		For j=0 To 3
			Ados1.Position=i*32+j*320
			Ados.Position=Pos+30*j+i*120
			Ados.write ados1.read(30)
		Next
	Next
	
	response.ContentType = "image/BMP"
	Ados.Position=0
	response.BinaryWrite Ados.read()
	Ados.Close:set Ados=nothing
	Ados1.Close:set Ados1=nothing
	
	If err Then Session(key) = 9999
End Function
%>