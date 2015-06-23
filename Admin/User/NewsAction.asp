<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
on error resume next
Dim User_Conn,NewsID,isLock,Conn
MF_Default_Conn
MF_User_Conn
MF_Session_TF

NewsID=NoSqlHack(Request("newsid"))
isLock=NoSqlHack(Request("value"))
if isLock="true" then
	if isnumeric(NewsID) then 
		User_Conn.execute("update FS_ME_News set isLock=1 where NewsID="&NewsID)
	elseif NewsID="all" then
		User_Conn.execute("update FS_ME_News set isLock=1")
	end if
elseif isLock="false" then
	if isnumeric(NewsID) then 
		User_Conn.execute("update FS_ME_News set isLock=0 where NewsID="&NewsID)
	elseif NewsID="all" then
		User_Conn.execute("update FS_ME_News set isLock=0")
	end if
end if
if Request("act")="delete" then
if request.Form("DeleteNews")<>"" then
	User_Conn.Execute("Delete from FS_ME_News where NewsID in ("&FormatIntArr(Request("deleteNews"))&")")
	if err.number>0 then
		Response.Redirect("../error.asp?ErrCodes="&err.description&"ErrorUrl="&Request.ServerVariables("HTTP_REFERER"))
		Response.End()
	else
		Response.Redirect("../success.asp?ErrCodes=<li>删除操作成功</li>&ErrorUrl="&Request.ServerVariables("HTTP_REFERER"))
		Response.End()
	end if
else
	Response.Redirect("../error.asp?ErrCodes=<li>你必须至少选择一个记录</li>&ErrorUrl="&Request.ServerVariables("HTTP_REFERER"))
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






