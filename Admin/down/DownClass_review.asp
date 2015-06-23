<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/Dynamic_Function.asp" -->
<%
Dim conn,Str_NS_NewsClassID,Rs_NS_NewsClass,Str_NS_NewsClass_Templet,StrSql,User_Conn
Dim f_Phy_Templet,f_FSO_OBJ,f_FILE_OBJ,f_FILE_STREAM_OBJ,f_File_Content,f_PAGES_DICT_OBJ,Page
MF_Default_Conn
MF_User_Conn
Str_NS_NewsClassID=NoSqlHack(Trim(Request("id")))
Page=NoSqlHack(Trim(Request("Page")))
If isnumeric(Page) then
	Page=Cint(Page)
Else
	Page=0
End if
If Str_NS_NewsClassID<>"" Then
	StrSql="SELECT Templet FROM FS_DS_Class WHERE ReycleTF=0 AND ClassID='"&NoSqlHack(Str_NS_NewsClassID)&"'"
	'Response.Write StrSql
	Set Rs_NS_NewsClass=conn.Execute(StrSql)
	If Not Rs_NS_NewsClass.eof Then   
		Str_NS_NewsClass_Templet=Rs_NS_NewsClass(0)
	Else
		Response.Write "<script language=""JavaScript"" type=""text/javascript"" src=""../FS_Inc/PublicJs.js""></script>"&vbcrlf
		Response.Write "<SCRIPT LANGUAGE=""JavaScript"">ShowErr(""错误：\n参数错误"");</SCRIPT>"
	End If
Else
	Response.Write "<script language=""JavaScript"" type=""text/javascript"" src=""../FS_Inc/PublicJs.js""></script>"&vbcrlf
	Response.Write "<SCRIPT LANGUAGE=""JavaScript"">ShowErr(""错误：\n参数错误"");</SCRIPT>"
End If
If G_VIRTUAL_ROOT_DIR<>"" Then
	Str_NS_NewsClass_Templet="/"&G_VIRTUAL_ROOT_DIR&Str_NS_NewsClass_Templet
Else
	Str_NS_NewsClass_Templet=Str_NS_NewsClass_Templet
End If
Response.Write Get_Dynamic_Refresh_Content(Str_NS_NewsClass_Templet,Str_NS_NewsClassID,"DS",Page,"down")
%>





