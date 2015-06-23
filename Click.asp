<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<!--#include file="FS_Inc/Const.asp" -->
<!--#include file="FS_InterFace/MF_Function.asp" -->
<!--#include file="FS_Inc/Function.asp" -->
<%session.CodePage="936"%>
<%
response.Charset="gb2312"
''前台页面,由JS调用得到 调用该文件必须给定一些参数.
Dim Conn
MF_Default_Conn
Dim stype,SubSys,spanid
Dim Server_Name,Server_V1,Server_V2,Cookie_Domain,TmpArr
Cookie_Domain = Get_MF_Domain()
if Cookie_Domain="" then      
	Cookie_Domain = "localhost"
else
	if left(lcase(Cookie_Domain),len("http://"))="http://" then Cookie_Domain = mid(Cookie_Domain,len("http://")+1)
	if right(Cookie_Domain,1)="/" then Cookie_Domain = mid(Cookie_Domain,1,len(Cookie_Domain) - 1)
end if	


''防盗连
Dim Main_Name,Name_Str1,V_MainName,V_Str
Server_Name = NoHtmlHackInput(NoSqlHack(LCase(Trim(Request.ServerVariables("SERVER_NAME")))))
IF Server_Name <> LCase(Split(Cookie_Domain,"/")(0)) Then
	call HTMLEnd("没有权限，请访问"&Cookie_Domain,"http://"&Cookie_Domain)
End If
Server_V1 = NoHtmlHackInput(NoSqlHack(Trim(Replace(Lcase(Cstr(Request.ServerVariables("HTTP_REFERER"))),"http://",""))))
Server_V1 = Replace(Replace(Server_V1,"//","/"),"///","/")
IF Server_V1 = "" Then
	call HTMLEnd("没有权限，请访问"&Cookie_Domain,"http://"&Cookie_Domain)
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
	call HTMLEnd("没有权限，请访问"&Cookie_Domain,"http://"&Cookie_Domain)
End If

stype = NoHtmlHackInput(NoSqlHack(Trim(request.QueryString("type")))) 'ajax
SubSys = NoHtmlHackInput(NoSqlHack(Trim(request.QueryString("SubSys")))) 'SubSys
spanid = NoHtmlHackInput(NoSqlHack(Trim(request.QueryString("spanid")))) ''NS_id_click_ECB1FJ2E8K82DK9
if stype="" then stype="ajax"
if stype<>"ajax" then call HTMLEnd("Error:type not is ajax!","http://"&Cookie_Domain)
if SubSys="" then call HTMLEnd("Error:SubSys is null!","http://"&Cookie_Domain)
if spanid="" then call HTMLEnd("Error:spanid is null!","http://"&Cookie_Domain)
TmpArr = split(spanid,"_")
if ubound(TmpArr)<3 then call HTMLEnd("Error:spanid's _ is Err!","http://"&Cookie_Domain)
	
Sub HTMLEnd(Info,URL)
	response.Write("$('"&spanid&"').innerHTML='';"&vbNewLine)
	response.End()
End Sub

response.Write("function f_Click_"&spanid&"() {new Ajax.Updater('"&spanid&"', 'http://"&Cookie_Domain&"/Click_Ajax.asp?no-cache='+Math.random() , {method: 'get', parameters: 'type="&stype&"&SubSys="&SubSys&"&Get="&NoSqlHack(request.QueryString("Get"))&"&spanid="&spanid&"'})};"&vbNewLine)
response.Write("setTimeout('f_Click_"&spanid&"()',50);"&vbNewLine)
Conn.close
Set Conn=Nothing
%>