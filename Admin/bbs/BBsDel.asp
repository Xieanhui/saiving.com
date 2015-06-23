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
if not MF_Check_Pop_TF("WS002") then Err_Show%>
<html>
<HEAD>
<TITLE>FoosunCMS留言系统</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<script language="javascript">
function ShowNote(NoteID,ClassName,ClassID)
{
location="ShowNote.asp?NoteID="+NoteID+"&ClassName="+ClassName+"&ClassID="+ClassID;
}
</script>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<body>
<%
Dim NoteID,ClassName,ClassID,strShowErr
if request.QueryString("Act")="SinglDel" then
	if Request.QueryString("BBSID")="" then
		strShowErr = "<li>参数出错!</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	end if
	NoteID=Request.QueryString("NoteID")
	ClassName=Request.querystring("ClassName")
	ClassID=Request.querystring("ClassID")
	Response.write(ClassID)
	Conn.execute("Delete From FS_WS_BBS Where ID="&CintStr(Request.QueryString("BBSID"))&"")
	Response.write("'<script>ShowNote("&NoteID&",'"&ClassName&"','"&ClassID&"');</script>'")
end if
%>
</body>
</html>






