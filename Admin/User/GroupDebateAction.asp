<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
on error resume next
Dim User_Conn,DebateID,isLock
MF_Default_Conn
MF_User_Conn
MF_Session_TF
DebateID=NoSqlHack(Request.QueryString("DebateID"))
isLock=NoSqlHack(Request.QueryString("value"))
if isLock="true" then
	if isnumeric(DebateID) then 
		User_Conn.execute("update FS_ME_GroupDebate set isLock=1 where DebateID="&DebateID)
	elseif DebateID="all" then
		User_Conn.execute("update FS_ME_GroupDebate set isLock=1")
	end if
elseif isLock="false" then
	if isnumeric(DebateID) then 
		User_Conn.execute("update FS_ME_GroupDebate set isLock=0 where DebateID="&DebateID)
	elseif DebateID="all" then
		User_Conn.execute("update FS_ME_GroupDebate set isLock=0")
	end if
end if
if Request.QueryString("act")="Delete" then
	User_Conn.Execute("Delete from FS_ME_GroupDebate where DebateID in ("&FormatIntArr(Request.QueryString("Debatecheck"))&")")
	if err.number>0 then
		Response.Redirect("../error.asp?ErrCodes="&err.description&"ErrorUrl="&Request.ServerVariables("HTTP_REFERER"))
		Response.End()
	else
		Response.Redirect("../success.asp?ErrCodes=<li>É¾³ý²Ù×÷³É¹¦</li>&ErrorUrl="&Request.ServerVariables("HTTP_REFERER"))
		Response.End()
	end if
end if

User_Conn.close
set User_Conn=nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>FoosunCMS</title>
</head>
<body>

</body>
</html>






