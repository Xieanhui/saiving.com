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
			Outstr = Outstr & "<td align=left>&nbsp;<span class=""tx""><b>��</b></span></td>"
			Else
			Outstr = Outstr & "<td align=left>&nbsp;<b>��</b>" & getver(theObj(i,0)) & "</td>"
			End If
			Outstr = Outstr & "</tr>" & vbCrLf
		next
		Response.Write(Outstr)
	End Function
	
	Public Function cdrivetype(tnum)
		Select Case tnum
			Case 0: cdrivetype = "δ֪"
			Case 1: cdrivetype = "���ƶ�����"
			Case 2: cdrivetype = "����Ӳ��"
			Case 3: cdrivetype = "�������"
			Case 4: cdrivetype = "CD-ROM"
			Case 5: cdrivetype = "RAM ����"
		End Select
	end function
	
	Private Sub Class_Initialize()
		WebName="FoosunCMS"
		WebUrl="http://www.Foosun.Cn"
		SysName="̽��"		
		SysNameE="ActiveXCheck"
		SysVersion=" "
		FileName=Request.ServerVariables("SCRIPT_NAME")
	End Sub
	
	Public Function dtype(num)
		Select Case num
			Case 0: dtype = "δ֪"
			Case 1: dtype = "���ƶ�����"
			Case 2: dtype = "����Ӳ��"
			Case 3: dtype = "�������"
			Case 4: dtype = "CD-ROM"
			Case 5: dtype = "RAM ����"
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
				OutStr = OutStr & "<div align=center>ҳ��ִ��ʱ�䣺Լ"& FormatNumber(runtime/1000,2) & "��</div>"
			else
				OutStr = OutStr & "<div align=center>ҳ��ִ��ʱ�䣺Լ"& Runtime & "����</div>"
			end if	
		end if
		Response.Write(OutStr)
	End Sub
End class
%>





