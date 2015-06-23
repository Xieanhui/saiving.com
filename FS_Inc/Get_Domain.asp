<% Option Explicit %>
<!--#include file="Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="Function.asp" -->
<%
Dim Conn,Str_MF_Domain
MF_Default_Conn
Str_MF_Domain = GET_MF_Domain()
If Str_MF_Domain="" then
	Response.write "FSDomain="""";"
Else
	Response.write "FSDomain=""http://"&GET_MF_Domain()&""";"
End If
If G_VIRTUAL_ROOT_DIR="" Then
	Response.write "VirtualDir="""";"
Else
	Response.write "VirtualDir=""/"&G_VIRTUAL_ROOT_DIR&""";"
End if
%>





