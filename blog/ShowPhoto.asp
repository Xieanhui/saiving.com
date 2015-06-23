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
Str_ME_ilogID=NoSqlHack(Request.QueryString("ID"))
if not isnumeric(Str_ME_ilogID) or Str_ME_ilogID="" then
	response.Write("´íÎóµÄ²ÎÊý")
	response.End
Else
If G_VIRTUAL_ROOT_DIR<>"" Then
	If G_TEMPLETS_DIR<>"" Then
		Str_ME_ilog_Templet="/"&G_VIRTUAL_ROOT_DIR&"/"&G_TEMPLETS_DIR&"/ilog/ShowPhoto.htm"
	Else
		Str_ME_ilog_Templet="/"&G_VIRTUAL_ROOT_DIR&"/ilog/ShowPhoto.htm"
	End If
Else
	If G_TEMPLETS_DIR<>"" Then
		Str_ME_ilog_Templet="/"&G_TEMPLETS_DIR&"/ilog/ShowPhoto.htm"
	Else
		Str_ME_ilog_Templet="/ilog/ShowPhoto.htm"
	End If
End If
End if
Response.Write Get_Dynamic_Refresh_Content(Str_ME_ilog_Templet,Str_ME_ilogID,"ME",0,"photonews")
%>





