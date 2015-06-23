<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<%session.CodePage="936"%>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<%
response.Charset="gb2312"
dim req_number,iik
req_number = request.QueryString("number")
for iik = 1 to req_number
	response.Write(GetRamCode(14)&"|"&GetRamCode(6)&"&nbsp;"&vbcrlf)
next
%>





