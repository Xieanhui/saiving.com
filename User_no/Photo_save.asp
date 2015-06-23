<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
dim rs,str_tmppic,i_s,str_tmppicarr,p_FolderObj,p_FSO,p_FileObj
Set p_FSO = Server.CreateObject(G_FS_FSO)
if Request.Form("Action")="add" then
str_tmppic=Replace(NoSqlHack(Request.Form("pic_1"))&"|"&NoSqlHack(Request.Form("pic_2"))&"|"&NoSqlHack(Request.Form("pic_3"))," ","")
str_tmppicarr =split(str_tmppic,"|")
for i_s=0 to Ubound(str_tmppicarr)
	if str_tmppicarr(i_s)<>"" then
		If Left(LCase(str_tmppicarr(i_s)),7) <> "http://" And Left(str_tmppicarr(i_s),1) = "/" then
			p_FolderObj = p_FSO.GetFile(Server.MapPath(Replace(str_tmppicarr(i_s),"/","\"))).size
		Else
			p_FolderObj = 0
		End If	
		set rs= Server.CreateObject(G_FS_RS)
		rs.open "select * From FS_ME_Photo where 1=0",User_conn,1,3
		rs.addnew
		rs("title")=NoSqlHack(Request.Form("title"))
		rs("PicSavePath")=str_tmppicarr(i_s)
		rs("Content")=NoSqlHack(NoHtmlHackInput(Request.Form("Content")))
		rs("Addtime")=now
		if Request.Form("ClassID")<>"" then
			rs("ClassID")=clng(Request.Form("ClassID"))
		else
			rs("ClassID")=0
		end if
		rs("PicSize")=p_FolderObj
		rs("Hits")=0
		rs("UserNumber")=Fs_User.UserNumber
		rs.update
		rs.close:set rs=nothing
	end if
next
	strShowErr="<li>保存相册成功</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../PhotoManage.asp")
	Response.end
elseif Request.Form("Action")="edit" then
		If Left(LCase(Request.Form("pic_1")),7) <> "http://" And Left(Request.Form("pic_1"),1) = "/" then
			p_FolderObj = p_FSO.GetFile(Server.MapPath(Replace(Request.Form("pic_1"),"/","\"))).size
		Else
			p_FolderObj = 0
		End If
		set rs= Server.CreateObject(G_FS_RS)
		rs.open "select * From FS_ME_Photo where id="&CintStr(Request.Form("id")),User_conn,1,3
		rs("title")=NoSqlHack(Request.Form("title"))
		rs("PicSavePath")=NoSqlHack(Request.Form("pic_1"))
		rs("Content")=NoSqlHack(NoHtmlHackInput(Request.Form("Content")))
		if Request.Form("ClassID")<>"" then
			rs("ClassID")=clng(Request.Form("ClassID"))
		else
			rs("ClassID")=0
		end if
		rs("PicSize")=p_FolderObj
		rs.update
		rs.close:set rs=nothing
		strShowErr="<li>修改成功</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../PhotoManage.asp")
		Response.end
end if
set p_FSO=nothing
set Fs_User=nothing
%><!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->





