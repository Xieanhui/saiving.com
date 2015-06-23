<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_InterFace/Dynamic_Function.asp" -->
<%
Dim conn,User_Conn,Str_ME_ilogID,Rs_ME_ilog,Str_ME_ilog_Templet,StrSql
Dim f_Phy_Templet,f_FSO_OBJ,f_FILE_OBJ,f_FILE_STREAM_OBJ,f_File_Content,f_PAGES_DICT_OBJ
MF_Default_Conn
MF_User_Conn
MFConfig_Cookies
StrSql="SELECT TempletSavePath FROM FS_ME_InfoiLogTemplet WHERE isDefault=1"
Set Rs_ME_ilog=User_Conn.Execute(StrSql)
If Not Rs_ME_ilog.eof Then
	Str_ME_ilog_Templet=Rs_ME_ilog("TempletSavePath")
Else
	Response.Write "<script language=""JavaScript"" type=""text/javascript"" src=""../FS_Inc/PublicJs.js""></script>"&vbcrlf
	Response.Write "<SCRIPT LANGUAGE=""JavaScript"">ShowErr(""´íÎó£º\n²ÎÊı´íÎó"");</SCRIPT>"
End If
If G_VIRTUAL_ROOT_DIR<>"" Then
	Str_ME_ilog_Templet="/"&G_VIRTUAL_ROOT_DIR&Str_ME_ilog_Templet&"/index.htm"
Else
	Str_ME_ilog_Templet=Str_ME_ilog_Templet&"/index.htm"
End If
Response.Write Get_Dynamic_Refresh_Content(Str_ME_ilog_Templet,Str_ME_ilogID,"ME",0,"")
%>





