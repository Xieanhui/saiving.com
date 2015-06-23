<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
p_Style_num = Request.QueryString("Style_num")
if p_Style_num = "1" then
	response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") ="1"
Elseif p_Style_num = "2" then
	response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") ="2"
Elseif p_Style_num = "3" then
	response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") ="3"
Elseif p_Style_num = "4" then
	response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") ="4"
Elseif p_Style_num = "5" then
	response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") ="5"
Else
	response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") ="1"
End if
User_Conn.execute("Update FS_ME_Users set MySkin = "& CintStr(Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")) &" where UserNumber = '"& NoSqlHack(Session("FS_UserNumber")) &"' and UserPassword='"& NoSqlHack(Session("FS_UserPassword"))&"'")
dim ReturnUrl
ReturnUrl = Replace(Request.QueryString("ReturnUrl"),"''","")
if Replace(Trim(ReturnUrl),"?","")<>"" then
	Response.Redirect ReturnUrl
	Response.end
Else
	Response.Redirect"main.asp"
	Response.end
End if
%>





