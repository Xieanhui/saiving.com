<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<!--#include file="FS_Inc/Const.asp" -->
<!--#include file="FS_InterFace/MF_Function.asp" -->
<!--#include file="FS_Inc/Function.asp" -->
<%session.CodePage="936"%>
<%
response.Charset="gb2312"
''ǰ̨ҳ��,��JS���õõ� ���ø��ļ��������һЩ����.
Dim Conn
MF_Default_Conn
Dim stype,Id,SpanId
Dim Server_Name,Server_V1,Server_V2,Cookie_Domain,TmpArr
Cookie_Domain = Get_MF_Domain()
if Cookie_Domain="" then    
	Cookie_Domain = "localhost"
else
	if left(lcase(Cookie_Domain),len("http://"))="http://" then Cookie_Domain = mid(Cookie_Domain,len("http://")+1)
	if right(Cookie_Domain,1)="/" then Cookie_Domain = mid(Cookie_Domain,1,len(Cookie_Domain) - 1)
end if	
''������
Dim Main_Name,Name_Str1,V_MainName,V_Str
Server_Name = NoHtmlHackInput(NoSqlHack(LCase(Request.ServerVariables("SERVER_NAME"))))
if Request.ServerVariables("Server_Port")<>80 then
	Server_Name = Server_Name&":"&Request.ServerVariables("Server_Port")
end if
IF trim(Server_Name) <> trim(LCase(Split(Cookie_Domain,"/")(0))) Then
	call HTMLEnd("û��Ȩ�ޣ������"&Cookie_Domain,"http://"&Cookie_Domain)
End If
Server_V1 = NoHtmlHackInput(NoSqlHack(Replace(Lcase(Cstr(Request.ServerVariables("HTTP_REFERER"))),"http://","")))
Server_V1 = Replace(Replace(Server_V1,"//","/"),"///","/")
IF Server_V1 = "" Then
	call HTMLEnd("û��Ȩ�ޣ������"&Cookie_Domain,"http://"&Cookie_Domain)
End If
IF Instr(Server_V1,"/") = 0 Then
	Server_V2 = Server_V1
Else
	Server_V2 = Split(Server_V1,"/")(0)
End If	
If Instr(Server_Name,".") = 0 Then
	Main_Name = Server_Name
Else
	Name_Str1 = Split(Server_Name,".")(0)
	Main_Name = Trim(Replace(Server_Name,Name_Str1 & ".",""))
End If
If Instr(Server_V2,".") = 0 Then
	V_MainName = Server_V2
Else
	V_Str = Split(Server_V2,".")(0)
	V_MainName = Trim(Replace(Server_V2,V_Str & ".",""))
End If
If Main_Name <> V_MainName And (Main_Name = "" OR V_MainName = "") Then
	call HTMLEnd("û��Ȩ�ޣ������"&Cookie_Domain,"http://"&Cookie_Domain)
End If
stype = NoHtmlHackInput(NoSqlHack(request.QueryString("type"))) 'NS
Id = CintStr(request.QueryString("Id")) 'Id
SpanId = NoHtmlHackInput(NoSqlHack(request.QueryString("SpanId"))) '
If SpanId="" Then
	Response.End()
Else
	If stype="" Then
		Call HTMLEnd("",SpanId)
	End If

	If Id="" Or Not isnumeric(Id) Then
		Call HTMLEnd("",SpanId)
	End If
End If

Sub HTMLEnd(Info,SpanId)
	response.Write("$('"&SpanId&"').innerHTML='';"&vbNewLine)
	response.End()
End Sub

response.Write("function f_review_"&spanid&"() {new Ajax.Updater('"&SpanId&"', 'http://"&Cookie_Domain&"/ShowReview_Ajax.asp?no-cache='+Math.random() , {method: 'get', parameters: 'type="&stype&"&Id="&Id&"&SpanId="&SpanId&"' })};"&vbNewLine)
response.Write("setTimeout('f_review_"&spanid&"()',100);"&vbNewLine)
Conn.close
Set Conn=Nothing
%>





