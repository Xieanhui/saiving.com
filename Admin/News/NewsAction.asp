<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->

<%
Dim User_Conn,NewsID,isLock
MF_User_Conn
MF_Session_TF
NewsID=Request.QueryString("newsid")
isLock=NoSqlHack(Request.QueryString("value"))
if isLock="true" then
	if not Get_SubPop_TF("","NS005","NS","news") then Err_Show
	if isnumeric(NewsID) then 
		User_Conn.execute("update FS_ME_News set isLock=1 where NewsID="&NoSqlHack(NewsID))
	elseif NewsID="all" then
		User_Conn.execute("update FS_ME_News set isLock=1")
	end if
elseif isLock="false" then
	if not Get_SubPop_TF("","NS008","NS","news") then Err_Show
	if isnumeric(NewsID) then 
		User_Conn.execute("update FS_ME_News set isLock=0 where NewsID="&NewsID)
	elseif NewsID="all" then
		User_Conn.execute("update FS_ME_News set isLock=0")
	end if
end if
if Request.QueryString("act")="delete" then
	if not Get_SubPop_TF("","NS003","NS","news") then Err_Show
	User_Conn.Execute("Delete from FS_ME_News where NewsID in ("&FormatIntArr(Request.QueryString("deleteNews"))&")")
	User_Conn.Execute("Delete from FS_NS_TodayPic where NewsID in ("&FormatIntArr(Request.QueryString("deleteNews"))&")")
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






