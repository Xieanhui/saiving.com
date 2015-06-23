<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<%
Dim Conn
Dim Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
MF_Default_Conn
MF_Session_TF 
dim p_Style_num,fromUrl
p_Style_num = Request.QueryString("Style_num")
fromUrl = Request.QueryString("RURL")
if fromUrl = "" then fromUrl = Request.ServerVariables("HTTP_REFERER")
if p_Style_num = "1" then
	Session("Admin_Style_Num") ="1"
Elseif p_Style_num = "2" then
	Session("Admin_Style_Num") ="2"
Elseif p_Style_num = "3" then
	Session("Admin_Style_Num") ="3"
Elseif p_Style_num = "4" then
	Session("Admin_Style_Num") ="4"
Elseif p_Style_num = "5" then
	Session("Admin_Style_Num") ="5"
Else
	Session("Admin_Style_Num") ="1"
End if
Conn.execute("Update FS_MF_Admin set Admin_Style_Num = "& CintStr(Session("Admin_Style_Num")) &" where Admin_Name = '"& NoSqlHack(Temp_Admin_Name) &"' and Admin_Pass_Word='"& NoSqlHack(Session("Admin_Pass_Word"))&"'")
if fromUrl="" or Instr(LCase(fromUrl),"TopFrame.asp?") > 0 Or Instr(LCase(fromUrl),"sessionid=") > 0 then fromUrl="index.asp"
Response.Redirect fromUrl
Response.end
%>





