<%
class cls_ActiveXCheck
	Public FileName,WebName,WebUrl,SysName,SysNameE,SysVersion
	Public Function IsObjInstalled(strClassString)
		On Error Resume Next
		Dim xTestObj
		Set xTestObj = Server.CreateObject(strClassString)
		If Err Then
			IsObjInstalled = False
		else	
			IsObjInstalled = True
		end if
		Set xTestObj = Nothing
	End Function
	
	Public Function getver(Classstr)
		On Error Resume Next
		Dim xTestObj
		Set xTestObj = Server.CreateObject(Classstr)
		If Err Then
			getver=""
		else	
			getver=xTestObj.version
		end if
		Set xTestObj = Nothing
	End Function
	
	Public Function GetObjInfo(startnum,endnum)
		dim i,Outstr
		for i=startnum to endnum
			Outstr = Outstr & "<tr  align=center height=18 class=""hback""><TD align=left>&nbsp;" & theObj(i,0) & ""
			Outstr = Outstr & "&nbsp;"&theObj(i,1)&"</font>"
			Outstr = Outstr & "</td>"
			If Not IsObjInstalled(theObj(i,0)) Then 
			Outstr = Outstr & "<td align=left>&nbsp;<span class=""tx""><b>°¡</b></span></td>"
			Else
			Outstr = Outstr & "<td align=left>&nbsp;<b>°Ã</b>" & getver(theObj(i,0)) & "</td>"
			End If
			Outstr = Outstr & "</tr>" & vbCrLf
		next
		Response.Write(Outstr)
	End Function
	
	Public Function cdrivetype(tnum)
		Select Case tnum
			Case 0: cdrivetype = "Œ¥÷™"
			Case 1: cdrivetype = "ø…“∆∂Ø¥≈≈Ã"
			Case 2: cdrivetype = "±æµÿ”≤≈Ã"
			Case 3: cdrivetype = "Õ¯¬Á¥≈≈Ã"
			Case 4: cdrivetype = "CD-ROM"
			Case 5: cdrivetype = "RAM ¥≈≈Ã"
		End Select
	end function
	
	Private Sub Class_Initialize()
		WebName="FoosunCMS"
		WebUrl="http://www.Foosun.Cn"
		SysName="ÃΩ’Î"		
		SysNameE="ActiveXCheck"
		SysVersion=" "
		FileName=Request.ServerVariables("SCRIPT_NAME")
	End Sub
	
	Public Function dtype(num)
		Select Case num
			Case 0: dtype = "Œ¥÷™"
			Case 1: dtype = "ø…“∆∂Ø¥≈≈Ã"
			Case 2: dtype = "±æµÿ”≤≈Ã"
			Case 3: dtype = "Õ¯¬Á¥≈≈Ã"
			Case 4: dtype = "CD-ROM"
			Case 5: dtype = "RAM ¥≈≈Ã"
		End Select
	End Function
	
	Public Function formatdsize(dsize)
		if dsize>=1073741824 then
			formatdsize=Formatnumber(dsize/1073741824,2) & " GB"
		elseif dsize>=1048576 then
			formatdsize=Formatnumber(dsize/1048576,2) & " MB"
		elseif dsize>=1024 then
			formatdsize=Formatnumber(dsize/1024,2) & " KB"
		else
			formatdsize=dsize & "B"
		end if
	End Function
	
	Public Function formatvariables(str)
	on error resume next
	str = cstr(server.htmlencode(str))
	formatvariables=replace(str,chr(10),"<br>")
	End Function
	
	Public Sub ShowFooter()
		dim Endtime,Runtime,OutStr
		Endtime=timer()
		OutStr = ""
		Runtime=FormatNumber((endtime-startime)*1000,2) 
		if Runtime>0 then
			if Runtime>1000 then
				OutStr = OutStr & "<div align=center>“≥√Ê÷¥–– ±º‰£∫‘º"& FormatNumber(runtime/1000,2) & "√Î</div>"
			else
				OutStr = OutStr & "<div align=center>“≥√Ê÷¥–– ±º‰£∫‘º"& Runtime & "∫¡√Î</div>"
			end if	
		end if
		Response.Write(OutStr)
	End Sub
End class
%>





