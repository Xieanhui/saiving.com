<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/Dynamic_Function.asp" -->
<%
Dim conn,Str_DS_DownLoadID,Rs_DS_News,Str_DS_Down_Templet,StrSql,User_Conn
Dim f_Phy_Templet,f_FSO_OBJ,f_FILE_OBJ,f_FILE_STREAM_OBJ,f_File_Content,f_PAGES_DICT_OBJ,Page
MF_Default_Conn
MF_User_Conn
Str_DS_DownLoadID=NoSqlHack(Trim(Request("DownID")))
Page=NoSqlHack(Trim(Request("Page")))
If isnumeric(Page) then
	Page=Cint(Page)
Else
	Page=0
End if
If Str_DS_DownLoadID<>"" Then
	StrSql="SELECT NewsTemplet FROM FS_DS_List WHERE DownLoadID='"&NoSqlHack(Str_DS_DownLoadID)&"'"
	'Response.Write StrSql
	Set Rs_DS_News=conn.Execute(StrSql)
	If Not Rs_DS_News.eof Then
		Str_DS_Down_Templet=Rs_DS_News(0)
	Else
		Response.Write "<script language=""JavaScript"" type=""text/javascript"" src=""../../FS_Inc/PublicJs.js""></script>"&vbcrlf
		Response.Write "<SCRIPT LANGUAGE=""JavaScript"">ShowErr(""错误：\n参数错误1"");</SCRIPT>"
	End If
Else
	Response.Write "<script language=""JavaScript"" type=""text/javascript"" src=""../../FS_Inc/PublicJs.js""></script>"&vbcrlf
	Response.Write "<SCRIPT LANGUAGE=""JavaScript"">ShowErr(""错误：\n参数错误2"");</SCRIPT>"
End If
If G_VIRTUAL_ROOT_DIR<>"" Then
	Str_DS_Down_Templet="/"&G_VIRTUAL_ROOT_DIR&Str_DS_Down_Templet
Else
	Str_DS_Down_Templet=Str_DS_Down_Templet
End If
Response.Write Get_Dynamic_Refresh_Content(Str_DS_Down_Templet,Str_DS_DownLoadID,"DS",Page,"news")
%>