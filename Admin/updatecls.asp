<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
	Dim Conn,strShowErr
	MF_Default_Conn
	'session判断
	MF_Session_TF
	SubSys_Cookies:MFConfig_Cookies:NSConfig_Cookies:DSConfig_Cookies
	if Request.Cookies("FoosunSUBCookie")("FoosunSUBMS")=1 Then:MSConfig_Cookies:end if
	strShowErr = "<li>更新缓存成功!</li>"
	Response.Redirect("success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
%>