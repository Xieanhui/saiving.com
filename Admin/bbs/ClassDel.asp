<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/ns_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn
MF_Default_Conn
'session判断
MF_Session_TF 
if not MF_Check_Pop_TF("WS002") then Err_Show
%>
<html>
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY>
<%
Dim Act,ID,strShowErr,i
Act=Request.QueryString("Act")
if Act="single" then
ID=CintStr(Request.QueryString("ID"))
if ID="" Then
	strShowErr="<li>你必须选择一项再删除</li>"
	Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
End if
Conn.execute("delete From FS_WS_Class Where ID="&ID&"")
		strShowErr = "<li>操作成功</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
end if
if Act="del" then
	ID=Request.form("ID")
	if ID="" Then
		strShowErr="<li>你必须选择一项再删除</li>"
		Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	End if
	ID=split(ID,",")
	For i=LBound(ID) To UBound(ID) 
		Conn.execute("delete From FS_WS_Class Where ID="&CintStr(ID(i))&"")
	Next
		strShowErr = "<li>操作成功</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
End if
Set Conn=nothing
%>
</body>
</html>






