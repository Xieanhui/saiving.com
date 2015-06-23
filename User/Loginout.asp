<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="../FS_Inc/Md5.asp" -->
<!--#include file="../API/Cls_PassportApi.asp" -->
<%
Dim UrlAdd,forward,Fs_User
forward=Request.QueryString("forward")
'-----------------------------------------------------------------
'系统整合
'-----------------------------------------------------------------
Dim API_Obj,API_SaveCookie,SysKey
If API_Enable Then
	SysKey = Md5(Session("FS_UserName")&API_SysKey,16)
	Set API_Obj = New PassportApi
		API_SaveCookie = API_Obj.SetCookie(SysKey,Session("FS_UserName"),"","")
	Set API_Obj = Nothing
End If
'-----------------------------------------------------------------
Set Fs_User = New Cls_User
Call Fs_User.out
Set Fs_User = Nothing
If forward = "" then
	forward = Request.ServerVariables("HTTP_REFERER")
End If
If forward="" Then
	forward = left(ThisPage,InStrRev(ThisPage,"/"))&"Main.asp"
End If
Response.Write API_SaveCookie
Response.Flush
Response.Write "<script language=""JavaScript"">window.location.href="""&forward&""";</script>"
Response.End
%>





