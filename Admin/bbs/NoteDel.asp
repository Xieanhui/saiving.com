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
<TITLE>FoosunCMS留言系统</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<body>
<%
Dim NoteId,i,strShowErr
if Request.QueryString("Act")="single" then
	if Request.querystring("ID")<>"" then
		Conn.execute("Delete From FS_WS_BBS where ID="&CintStr(Request.querystring("ID"))&" and ParentID='0' ")
		Conn.execute("Delete From FS_WS_BBS where ParentID='"&CintStr(Request.querystring("ID"))&"'")
		strShowErr = "<li>操作成功</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=bbs/ClassMessageManager.asp")
		Response.end
	else
		strShowErr = "<li>参数出错!</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	end if
end if
if Request.QueryString("Act")="del" then
 	if Request.form("NoteID")<>""then
		NoteID=Request.form("NoteID")
		NoteID=split(NoteID,",")
		For i=LBound(NoteID) To UBound(NoteID)
			Conn.execute("Delete From FS_WS_BBS where ID="&CintStr(NoteID(i))&"")
		Next
		strShowErr = "<li>操作成功</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=bbs/ClassMessageManager.asp")
		Response.end
	else
		strShowErr = "<li>参数出错!</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	end if
end if
Set Conn=nothing
%>
</body>
</html>






