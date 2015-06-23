<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/ns_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
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
<body>
<%
dim Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
dim ID,i,strShowErr
If Request.querystring("Act")="singledel" then
	ID=Request.querystring("ID")
	if ID="" then
		strShowErr = "<li>参数不足</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=BBS/NewTell.asp")
		Response.end
	else
		Conn.execute("delete From FS_WS_NewsTell Where ID="&CintStr(ID)&" And AddUser = '" & Temp_Admin_Name & "'")
		strShowErr = "<li>删除成功</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=BBS/NewTell.asp")
		Response.end
	end if
end if
If Request.queryString("Act")="del" then
	ID=Request.form("TellID")
	if ID="" then
		strShowErr = "<li>参数不足</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=BBS/NewTell.asp")
		Response.end
	else
		ID=split(ID,",")
		For i=LBound(ID) to UBound(ID)
			Conn.execute("delete From FS_WS_NewsTell Where ID="&CintStr(ID(i))&" And AddUser = '" & Temp_Admin_Name & "'")
		Next
		strShowErr = "<li>删除成功</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=BBS/NewTell.asp")
		Response.end
	end if
end if
Set Conn=nothing
%>
</body>
</html>






