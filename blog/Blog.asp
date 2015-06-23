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
'User_Conn
MFConfig_Cookies
Str_ME_ilogID=NoSqlHack(Trim(Request("ID")))
if not isnumeric(Str_ME_ilogID) or Str_ME_ilogID="" then
	response.Write("错误的参数")
	response.End
Else
	StrSql="SELECT FS_ME_InfoiLogParam.TempletID FROM FS_ME_Infoilog,FS_ME_InfoiLogParam WHERE FS_ME_Infoilog.UserNumber=FS_ME_InfoiLogParam.UserNumber AND FS_ME_Infoilog.iLogID="&CintStr(Str_ME_ilogID)
	Set Rs_ME_ilog=User_Conn.Execute(StrSql)
	If Not Rs_ME_ilog.eof Then
		Str_ME_ilog_Templet=Rs_ME_ilog("TempletID")
	Else
		Response.Write "<script language=""JavaScript"" type=""text/javascript"" src=""../FS_Inc/PublicJs.js""></script>"&vbcrlf
		Response.Write "<SCRIPT LANGUAGE=""JavaScript"">ShowErr(""错误：\n参数错误"");</SCRIPT>"
	End If
End If
If G_VIRTUAL_ROOT_DIR<>"" Then
	Str_ME_ilog_Templet="/"&G_VIRTUAL_ROOT_DIR&Str_ME_ilog_Templet&"/page.htm"
Else
	Str_ME_ilog_Templet=Str_ME_ilog_Templet&"/page.htm"
End If
Response.Write Get_Dynamic_Refresh_Content(Str_ME_ilog_Templet,Str_ME_ilogID,"ME",0,"news")
%>





