<!--#include file="Const.asp" -->
<!--#include file="Function.asp"-->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.Charset = "GB2312"
Dim Conn
MF_Default_Conn


'-------------------------------------------
Dim DivID,Sql,Rs,DisStr,i,ChlidID,StrType
DivID = NoSqlHack(Request.QueryString("act"))
ChlidID = NoSqlHack(Request.QueryString("ID"))
StrType = NoSqlHack(Request.QueryString("StrType"))
If ChlidID = "" Or Not IsNumeric(ChlidID) Then
	ChlidID = 0
Else
	ChlidID = ChlidID
End If
'------------------------		
Select Case DivID
	Case "JobCity"
		Sql = "Select PID,Province From FS_AP_Province Where 1=1 Order By PID Desc"
		Set Rs = Conn.ExeCute(Sql)
		If Not Rs.Eof Then
			DisStr = "<table width=""594"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbnewline
			DisStr = DisStr & "<tr>" & vbnewline & "<td height=""25"" colspan=""6"" align=""left"" valign=""middle"" style=""font-size:12px; color:#3399CC;"">��ѡ��ʡ�� [<span onclick=""GetSelectValue('','City','ѡ�����ص�')"" style=""cursor:hand;""><font color=""#FF0000"">���޹����ص�</font></span>]</td>" & vbnewline & "</tr>" & vbnewline
			Do While Not Rs.Eof
				DisStr = DisStr & "<tr>" & vbnewline
				For i = 1 To 6
				If Rs.Eof Then Exit For	
					DisStr = DisStr & "<td height=""20"" align=""left"" width=""99"" valign=""middle"" style=""font-size:12px; color:#3399CC;"">"
					DisStr = DisStr & "<span onclick=""DisChildDiv(" & Rs(0) & ",'City')"" style=""cursor:hand;""><font color=""#3399CC"">" & Rs(1) & "</font></span>"
					DisStr = DisStr & "</td>" & vbnewline
				Rs.MoveNext
				Next
				DisStr = DisStr & "</tr>" & vbnewline	
			Loop
			DisStr = DisStr & "</table>"
		Else
			DisStr = "<span style=""color:#3399CC"">��û������ʡ��</span>"
		End If
		Rs.Close : Set Rs = Nothing
'-----------
	Case "JobType"
		Sql = "Select TID,Trade From FS_AP_Trade Where 1=1 Order By TID Desc"
		Set Rs = Conn.ExeCute(Sql)
		If Not Rs.Eof Then
			DisStr = "<table width=""594"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbnewline
			DisStr = DisStr & "<tr>" & vbnewline & "<td height=""25"" colspan=""6"" align=""left"" valign=""middle"" style=""font-size:12px; color:#3399CC;"">��ѡ����ҵ [<span onclick=""GetSelectValue('','Job','ѡ����ҵ/ְλ')"" style=""cursor:hand;""><font color=""#FF0000"">������ҵ</font></span>]</td>" & vbnewline & "</tr>" & vbnewline
			Do While Not Rs.Eof
				DisStr = DisStr & "<tr>" & vbnewline
				For i = 1 To 6
				If Rs.Eof Then Exit For	
					DisStr = DisStr & "<td height=""20"" align=""left"" width=""99"" valign=""middle"" style=""font-size:12px; color:#3399CC;"">"
					DisStr = DisStr & "<span onclick=""DisChildDiv(" & Rs(0) & ",'Job_Type')"" style=""cursor:hand;""><font color=""#3399CC"">" & Rs(1) & "</font></span>"
					DisStr = DisStr & "</td>" & vbnewline
				Rs.MoveNext
				Next
				DisStr = DisStr & "</tr>" & vbnewline	
			Loop
			DisStr = DisStr & "</table>"
		Else
			DisStr = "<span style=""color:#3399CC"">��û��������ҵ</span>"
		End If
		Rs.Close : Set Rs = Nothing
	Case "JobTime"
		DisStr = "<table width=""594"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbnewline
		DisStr = DisStr & "<tr>" & vbnewline
		DisStr = DisStr &  "<td height=""20"" align=""left"" width=""99"" valign=""middle"" style=""font-size:12px; color:#3399CC;"">"
		DisStr = DisStr & "<span onclick=""GetSelectValue('','StrTime','ѡ��ʱ�䷶Χ')"" style=""cursor:hand;""><font color=""#FF0000"">����ʱ��</font></span>"
		DisStr = DisStr & "</td>" & vbnewline
		DisStr = DisStr &  "<td height=""20"" align=""left"" width=""99"" valign=""middle"" style=""font-size:12px; color:#3399CC;"">"
		DisStr = DisStr & "<span onclick=""GetSelectValue('1','StrTime','���һ��')"" style=""cursor:hand;""><font color=""#3399CC"">���һ��</font></span>"
		DisStr = DisStr & "</td>" & vbnewline
		DisStr = DisStr &  "<td height=""20"" align=""left"" width=""99"" valign=""middle"" style=""font-size:12px; color:#3399CC;"">"
		DisStr = DisStr & "<span onclick=""GetSelectValue('3','StrTime','�������')"" style=""cursor:hand;""><font color=""#3399CC"">�������</font></span>"
		DisStr = DisStr & "</td>" & vbnewline
		DisStr = DisStr &  "<td height=""20"" align=""left"" width=""99"" valign=""middle"" style=""font-size:12px; color:#3399CC;"">"
		DisStr = DisStr & "<span onclick=""GetSelectValue('7','StrTime','���һ��')"" style=""cursor:hand;""><font color=""#3399CC"">���һ��</font></span>"
		DisStr = DisStr & "</td>" & vbnewline
		DisStr = DisStr &  "<td height=""20"" align=""left"" width=""99"" valign=""middle"" style=""font-size:12px; color:#3399CC;"">"
		DisStr = DisStr & "<span onclick=""GetSelectValue('15','StrTime','�������')"" style=""cursor:hand;""><font color=""#3399CC"">�������</font></span>"
		DisStr = DisStr & "</td>" & vbnewline
		DisStr = DisStr &  "<td height=""20"" align=""left"" width=""99"" valign=""middle"" style=""font-size:12px; color:#3399CC;"">"
		DisStr = DisStr & "<span onclick=""GetSelectValue('30','StrTime','���һ��')"" style=""cursor:hand;""><font color=""#3399CC"">���һ��</font></span>"
		DisStr = DisStr & "</td>" & vbnewline
		DisStr = DisStr &  "<td height=""20"" align=""left"" width=""99"" valign=""middle"" style=""font-size:12px; color:#3399CC;"">"
		DisStr = DisStr & "<span onclick=""GetSelectValue('90','StrTime','�������')"" style=""cursor:hand;""><font color=""#3399CC"">�������</font></span>"
		DisStr = DisStr & "</td>" & vbnewline
		DisStr = DisStr & "</tr>" & vbnewline
		DisStr = DisStr & "</table>" & vbnewline
	Case "GetChlid"
		If StrType = "City" Then
			Sql = "Select CID,City From FS_AP_City Where PID = " & ChlidID & " Order By CID Desc"
			Set Rs = Conn.ExeCute(Sql)
			If Not Rs.Eof Then
				DisStr = "<table width=""594"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbnewline
				Do While Not Rs.Eof
					DisStr = DisStr & "<tr>" & vbnewline
					For i = 1 To 6
					If Rs.Eof Then Exit For	
						DisStr = DisStr & "<td height=""20"" align=""left"" width=""99"" valign=""middle"" style=""font-size:12px; color:#3399CC;"">"
						DisStr = DisStr & "<span onclick=""GetSelectValue(" & Rs(0) & ",'City','" & Rs(1) & "')"" style=""cursor:hand;""><font color=""#3399CC"">" & Rs(1) & "</font></span>"
						DisStr = DisStr & "</td>" & vbnewline
					Rs.MoveNext
					Next
					DisStr = DisStr & "</tr>" & vbnewline	
				Loop
				DisStr = DisStr & "</table>"
			Else
				DisStr = "<span style=""color:#3399CC"">��û�и�ʡ�ݵĳ�����Ϣ</span>"
			End If
			Rs.Close : Set Rs = Nothing
		ElseIf StrType = "Job_Type" Then
			Sql = "Select JID,Job From FS_AP_Job Where TID = " & ChlidID & " Order By JID Desc" 		
			Set Rs = Conn.ExeCute(Sql)
			If Not Rs.Eof Then
				DisStr = "<table width=""594"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbnewline
				Do While Not Rs.Eof
					DisStr = DisStr & "<tr>" & vbnewline
					For i = 1 To 6
					If Rs.Eof Then Exit For	
						DisStr = DisStr & "<td height=""20"" align=""left"" width=""99"" valign=""middle"" style=""font-size:12px; color:#3399CC;"">"
						DisStr = DisStr & "<span onclick=""GetSelectValue(" & Rs(0) & ",'Job','" & Rs(1) & "')"" style=""cursor:hand;""><font color=""#3399CC"">" & Rs(1) & "</font></span>"
						DisStr = DisStr & "</td>" & vbnewline
					Rs.MoveNext
					Next
					DisStr = DisStr & "</tr>" & vbnewline	
				Loop
				DisStr = DisStr & "</table>"
			Else
				DisStr = "<span style=""color:#3399CC"">��û�и���ҵ��ְλ��Ϣ</span>"
			End If
			Rs.Close : Set Rs = Nothing
		End If
End Select
Response.Write DisStr
Conn.Close : Set Conn = Nothing
%>





