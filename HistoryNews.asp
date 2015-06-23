<% Option Explicit %>
<!--#include file="FS_Inc/Const.asp" -->
<!--#include file="FS_Inc/Function.asp" -->
<!--#include file="FS_InterFace/MF_Function.asp" -->
<!--#include file="FS_InterFace/Dynamic_Function.asp" -->
<%
Dim conn,Old_News_Conn,Str_NS_PublicID,Rs_NS_Public,Str_NS_Templet,StrSql,Rs_NS_News,Str_NS_News_Templet
Dim f_Phy_Templet,f_FSO_OBJ,f_FILE_OBJ,f_FILE_STREAM_OBJ,f_File_Content,f_PAGES_DICT_OBJ,pageIndex
MF_Default_Conn
MF_Old_News_Conn
'User_Conn
Str_NS_PublicID=CintStr(Request("ID"))
pageIndex = Request("page")
if pageIndex="" then
	pageIndex=0
else
	pageIndex = cint(pageIndex)
end if
If Str_NS_PublicID="" Then
	Response.Write "<script language=""JavaScript"" type=""text/javascript"" src=""../FS_Inc/PublicJs.js""></script>"&vbcrlf
	Response.Write "<SCRIPT LANGUAGE=""JavaScript"">ShowErr(""´íÎó£º\n²ÎÊý´íÎó"");</SCRIPT>"
End If


Str_NS_Templet = "/templets/OldNews/index.htm"
If G_VIRTUAL_ROOT_DIR<>"" Then
	If G_TEMPLETS_DIR<>"" Then
		Str_NS_Templet="/"&G_VIRTUAL_ROOT_DIR&"/"&Str_NS_Templet
	Else
		Str_NS_Templet="/"&G_VIRTUAL_ROOT_DIR&Str_NS_Templet
	End If
Else
	If G_TEMPLETS_DIR<>"" Then
		Str_NS_Templet=Str_NS_Templet
	Else
		Str_NS_Templet=Str_NS_Templet
	End If
End If
Response.Write Get_Dynamic_Refresh_Content(Str_NS_Templet,Str_NS_PublicID&"R__D","NS",pageIndex,"news")
%>