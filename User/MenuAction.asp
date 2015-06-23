<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
Dim menuType,showValue
menuType=request.querystring("type")
showValue=request.querystring("value")
If Trim(""&showValue)="" Then
	showValue="0"
End if
Select Case menuType
	Case "msg" Response.Cookies("FoosunUserCookies")("FoosunMenuMsg")=NoSqlHack(showValue)
	Case "friend" Response.Cookies("FoosunUserCookies")("FoosunMenuFriend")=NoSqlHack(showValue)
	Case "special" Response.Cookies("FoosunUserCookies")("FoosunMenuSpecial")=NoSqlHack(showValue)
End Select
Response.write("menuType="&menuType&" and  showValue="&showValue)
%>
<%
Set Conn=Nothing
Set User_Conn=nothing
Set Fs_User = Nothing
%>






