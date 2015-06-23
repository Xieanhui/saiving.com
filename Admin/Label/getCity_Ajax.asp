<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<%session.CodePage="936"%>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
	response.Charset = "gb2312"
	Dim Conn,review_Sql,review_RS,TmpStr
	Dim pid
	TmpStr = "" 
	
	pid = CintStr(request.QueryString("PID"))  
	
	if not isnumeric(pid) then response.Write("省份必须选择"):response.End()
	
	MF_Default_Conn
	   
	TmpStr = "<select name=City id=City>"&vbnewline
	TmpStr = TmpStr & "<option value="""">城市不限</option>"&vbnewline
	TmpStr = TmpStr & Get_FildValue_List("Select CID,City from FS_AP_City where PID="&pid,"",1)
	TmpStr = TmpStr & "</select>"&vbnewline
	response.Write(TmpStr)
	Conn.close
	set Conn=Nothing
	
	
	Function Get_FildValue_List(This_Fun_Sql,EquValue,Get_Type)
	'''This_Fun_Sql 传入sql语句,EquValue与数据库相同的值如果是<option>则加上selected,Get_Type=1为<option>
	Dim Get_Html,This_Fun_Rs,Text
	On Error Resume Next
	This_Fun_Sql = Replace(This_Fun_Sql,"%","")
		if instr(This_Fun_Sql," FS_ME_")>0 then 
			set This_Fun_Rs = User_Conn.execute(This_Fun_Sql)
		else
			set This_Fun_Rs = Conn.execute(This_Fun_Sql)
		end if			
	If Err<>0 then Get_FildValue_List = "<option value="""">"&Err.description&"</option>"&vbnewline
	if isnull(EquValue) then EquValue = ""
	do while not This_Fun_Rs.eof 
		select case Get_Type
		  case 1
			''<option>		
			if instr(This_Fun_Sql,",") >0 then 
				Text = This_Fun_Rs(1)
			else
				Text = This_Fun_Rs(0)
			end if	
			if cstr(EquValue) = cstr(This_Fun_Rs(0)) then 
				Get_Html = Get_Html & "<option value="""&This_Fun_Rs(0)&"""  style=""color:#0000FF"" selected>"&Text&"</option>"&vbNewLine
			else
				Get_Html = Get_Html & "<option value="""&This_Fun_Rs(0)&""">"&Text&"</option>"&vbNewLine
			end if		
		  case else
			exit do : Get_FildValue_List = "<option value="""">Get_Type值传入错误</option>"&vbnewline : exit Function
		end select
		This_Fun_Rs.movenext
	loop
	This_Fun_Rs.close
	Get_FildValue_List = Get_Html
	End Function

%>






