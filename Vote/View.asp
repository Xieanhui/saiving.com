<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/zhupic.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn,strShowErr,TID,VS_Sql,VS_RS,VS_RS1,VS_RS2,PerVote,PerHeight,f_ii,InputValue,otherstr,DisMode
Dim Vs_Num
PerHeight = 60 
TID = CintStr(request.QueryString("TID"))
if TID="" or not isnumeric(TID) then response.Write("ͶƱ�����������!"&TID)  :  response.End()
Vs_Num = NoSqlHack(request.QueryString("Vs_Num"))
If Vs_Num = "" Or Not IsNumeric(Vs_Num) Then
	Response.Write("ͶƱ��������!")
	Response.End()
End If	
MF_Default_Conn

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<!--[if !mso]>
<style>
v\:*         { behavior: url(#default#VML) }
o\:*         { behavior: url(#default#VML) }
.shape       { behavior: url(#default#VML) }
</style>
<![endif]-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=request.QueryString("Title")%>---ͶƱ����鿴</title>
<style>
TD {	FONT-SIZE: 9pt}
</style>
</head>
<body>
<%
VS_Sql = "select  * from FS_VS_Steps where TID = "&TID&" order by Steps"
Set VS_RS = CreateObject(G_FS_RS)
VS_RS.Open VS_Sql,Conn,1,1
if VS_RS.eof then 
	response.Write("<table border=0 width=""100%"">")
	Call viewmain(TID)
	response.Write("</table>")	
else
	do while not VS_RS.eof
		response.Write("<table border=0 width=""100%"">")		
		Call viewmain(VS_RS("QuoteID"))
		PerHeight = PerHeight + 280
		response.Write("</table>")
		VS_RS.movenext	
	loop	
end if

sub viewmain(TIDValue)
	f_ii = 1
	VS_Sql = "select Theme,DisMode,IID,ItemName,DisMode,B.ItemMode,VoteCount,PicSrc,ItemDetail from FS_VS_Theme A,FS_VS_Items B where A.TID=B.TID AND A.TID ="&TIDValue
	Set VS_RS1 = CreateObject(G_FS_RS)
	VS_RS1.Open VS_Sql,Conn,1,1
	if not VS_RS1.eof then 
		DisMode=VS_RS1("DisMode")
		if DisMode = 1 then 
			'ֱ��ͼ
			response.Write("<tr><td height="&PerHeight&" valign=""top"" align=""left"" style=""font-size:14px;"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>"&VS_RS1("Theme")&"</b></td></tr>")
			redim total(VS_RS1.recordcount,2)
		end if		
	else
		response.Write("<tr><td>")
	end if		
	do while not VS_RS1.eof
		PerVote = VS_RS1("VoteCount")
		'''''''''''''''''''��ʵƱ��
		VS_Sql = "select count(*) from FS_VS_Items_Result where TID = "& CintStr(TIDValue) &" and IID = "&CintStr(VS_RS1("IID"))			
		set VS_RS2=Conn.execute(VS_Sql)
		if not VS_RS2.eof then 
			''��ǰѡ���Ʊ��
			PerVote = PerVote + VS_RS2(0)
		else
			PerVote = 0	
		end if
		VS_RS2.close
		''''''''''''''''''
		'''''''''''''''''''������д��
		VS_Sql = "select ItemValue from FS_VS_Items_Result where TID = "& CintStr(TIDValue) &" and IID = "&CintStr(VS_RS1("IID"))			
		set VS_RS2=Conn.execute(VS_Sql)
		if not VS_RS2.eof then 
			''��ǰѡ���Ʊ��
			if VS_RS2(0)<>"" then InputValue =  "---"&VS_RS2(0)
		else
			InputValue = ""	
		end if
		VS_RS2.close
		''''''''''''''''''
		'response.Write(VS_RS1.recordcount&":"&VS_RS1("IID")&":"&PerVote&"<br>")
		if VS_RS1("PicSrc")<>"" then 
			otherstr = "  style=""cursor:hand"" onClick=""window.open('"&otherstr&"');""" : InputValue = InputValue & " [����鿴��ͼ]"
		else
			otherstr = "  style=""cursor:help"""
		end if
		if VS_RS1("DisMode") = 1 then 
			'ֱ��ͼ
			total(f_ii,1) = PerVote
			total(f_ii,2)="<span title="""&VS_RS1("ItemName")&InputValue&""""&otherstr&">"&left(VS_RS1("ItemName"),10)&"</span>"
		end if
		
		f_ii = f_ii + 1		
		VS_RS1.movenext
	loop
	if f_ii>1 then 
		if DisMode = 1 then 
			'ֱ��ͼ
			response.Write("<tr><td>")
			call table1(total,0,PerHeight,10,20,768,200,"A",Vs_Num)
			response.Write("</td></tr>")
		end if
	else
		response.Write("û�з��������ļ�¼��</td></tr>")
	end if	
	VS_RS1.close
	set VS_RS1 = nothing
end sub

Function HtmEnd(Msg)
	response.Write(Msg&vbnewline)
	response.Write("</body></html>")
	response.End()
End Function
Sub ConnClose()
	Set Conn = Nothing
End Sub
%>
</body>
</html>