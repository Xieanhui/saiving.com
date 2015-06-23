<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/Dynamic_Function.asp" -->
<%
if not MF_Check_Pop_TF("NS013") then
	if not MF_Check_Pop_TF("NS001") then Err_Show
end if
Dim conn,Str_NS_NewsID,Rs_NS_News,Str_NS_News_Templet,StrSql,User_Conn,URLAddress
Dim f_Phy_Templet,f_FSO_OBJ,f_FILE_OBJ,f_FILE_STREAM_OBJ,f_File_Content,f_PAGES_DICT_OBJ,Page
MF_Default_Conn
MF_User_Conn
Str_NS_NewsID=NoSqlHack(Trim(Request("NewsID")))
Page=NoSqlHack(Trim(Request("Page")))
If isnumeric(Page) then
	Page=Cint(Page)
Else
	Page=0
End if
If Str_NS_NewsID<>"" Then
	StrSql="SELECT Templet,IsURL,URLAddress FROM FS_NS_News WHERE isDraft =0 and isRecyle=0 and NewsID='"&NoSqlHack(Str_NS_NewsID)&"'"
	Set Rs_NS_News=conn.Execute(StrSql)
	If Not Rs_NS_News.eof Then
		if Rs_NS_News("IsURL") = 1 then
			URLAddress = Rs_NS_News("URLAddress") & ""
			Rs_NS_News.Close
			Set Rs_NS_News = Nothing
			Response.Redirect(URLAddress)
			Response.End
		else
			Str_NS_News_Templet=Rs_NS_News(0)
		end if
	Else
		Rs_NS_News.Close
		Set Rs_NS_News = Nothing
		Response.Write "<script language=""JavaScript"" type=""text/javascript"" src=""../FS_Inc/PublicJs.js""></script>"&vbcrlf
		Response.Write "<SCRIPT LANGUAGE=""JavaScript"">ShowErr(""错误：\n参数错误"");</SCRIPT>"
		Response.End
	End If
	Rs_NS_News.Close
	Set Rs_NS_News = Nothing
Else
	Response.Write "<script language=""JavaScript"" type=""text/javascript"" src=""../FS_Inc/PublicJs.js""></script>"&vbcrlf
	Response.Write "<SCRIPT LANGUAGE=""JavaScript"">ShowErr(""错误：\n参数错误"");</SCRIPT>"
	Response.End
End If
If G_VIRTUAL_ROOT_DIR<>"" Then
	Str_NS_News_Templet="/"&G_VIRTUAL_ROOT_DIR&Str_NS_News_Templet
Else
	Str_NS_News_Templet=Str_NS_News_Templet
End If
Response.Write Get_Dynamic_Refresh_Content(Str_NS_News_Templet,Str_NS_NewsID,"NS",Page,"news")
%>