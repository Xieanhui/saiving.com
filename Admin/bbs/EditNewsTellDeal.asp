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
Dim AddUser,Topic,Content,Person,EditRs,ID,strShowErr
Set EditRs=Server.createobject(G_FS_RS)
if NoSqlHack(Request.querystring("Act"))="Edit" then
	ID=Request.form("ID")
	AddUser=NoSqlHack(NoHtmlHackInput(Request.form("AddUser")))
	Topic=NoSqlHack(NoHtmlHackInput(Request.form("Topic")))
	Content=NoHtmlHackInput(NoSqlHack(Replace(Request.form("Content"),vbcrlf,"<br>")))
	Person=NoSqlHack(NoHtmlHackInput(request.form("Person")))
	if AddUser="" then 
		strShowErr = "<li>用户名为空,请检楂用户登陆是否已过期!</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	end if
	if Topic="" then
		strShowErr = "<li>公告标题不能为空!</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	end if
	if Content="" then
		strShowErr = "<li>公告内容不能为空!</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	end if
	EditRs.open "Select Topic,Content,Person,AddUser,AddDate From FS_WS_NewsTell where ID="&CintStr(ID)&"",Conn,1,3
	if not EditRs.eof then
	EditRs("Topic")=Topic
	EditRs("Content")=Content
	EditRs("Person")=Person
	EditRs("AddUser")=AddUser
	EditRs.update
	end if
	Set EditRs=nothing	
		strShowErr = "<li>修改成功</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=BBS/NewTell.asp")
		Response.end
end if
%>





