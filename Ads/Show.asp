<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp"-->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_InterFace/NS_Function.asp" -->
<%
	response.buffer=true	
	Response.CacheControl = "no-cache"
	Dim Conn,int_ID,o_Ad_Rs
	MF_Default_Conn

 If request.QueryString("Location")<>"" and isnull(request.QueryString("Location"))=false then
	Conn.Execute("update FS_AD_Info set AdShowNum=AdShowNum+1 where AdID="&CintStr(request.QueryString("Location"))&"")
 End if
Set Conn=nothing
%>





