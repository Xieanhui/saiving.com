<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp"-->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_InterFace/ns_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Func_page.asp" -->
<%
Dim Conn,User_Conn
MF_Default_Conn
Dim Configobj,PageS,sql_1,MSTitle,ShowIP,s_IsUser,Style
Set Configobj=server.CreateObject (G_FS_RS)
sql_1="select ID,Title,IPShow,IsUser,IsAut,PageSize,Style From FS_WS_Config"
configobj.open sql_1,Conn,1,1
if not configobj.eof then
	PageS=configobj("PageSize")
	MSTitle=configobj("Title")
	ShowIP=configobj("IPShow")
	s_IsUser = configobj("IsUser")
	Style = configobj("Style")
	if Style<>"" then
		Style = Style
	else
		Style = "3"
	end if
end if
Response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") = Style
%>
<html>
<HEAD>
<TITLE><%=GetGuestBookTitle%></TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<link href="../<% = G_USER_DIR %>/images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</HEAD>
<%
if NoSqlHack(request.QueryString("action"))="showone" and NoSqlHack(Request.QueryString("boardid"))="0" then
	dim TellRs,Sql,DisID
	DisID = NoSqlHack(request.QueryString("DisID"))
	If DisID = "" Or ISnull(DisID) Or Not IsNUmeric(DisID) Then
		DisID = 1
	Else
		DisID = Clng(DisID)
	End If	
	Set TellRs=server.CreateObject (G_FS_RS)
	Sql="select ID,Topic,Content,Person,IsUse,PV,AddUser,AddDate From FS_WS_NewsTell Where ID =" & NoSqlHack(DisID)
	TellRs.open Sql,Conn,1,1
	if not TellRs.eof then
%>
      <table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr>
    <th valign="top" align="middle" class="xingmu"><%=TellRs("Topic")%></th>
  </tr>
  <tr>
    <td class="hback" style="LEFT: 0px; WIDTH: 100%; WORD-WRAP: break-word" valign="top"><br />
        <blockquote> <%=TellRs("Content")%> </blockquote>
    </td>
  </tr>
  <tr>
    <td class="hback" valign="center">
      <table cellspacing="0" cellpadding="0" width="100%" border="0">
        <tr>
          <td class="tablebody2" align="left">&nbsp;&nbsp;&nbsp;<b>发布人</b>： 
            <%=TellRs("AddUser")%>&nbsp;
            <bgsound src="dfdf" border="0" />
          </td>
          <td class="hback" align="right"><b>发布时间</b>： 
             <%=TellRs("AddDate")%>&nbsp;&nbsp;&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%
	end if
	Set TellRs=nothing
else
Response.Write("请不要直接提交本页!")
Response.end 
end if
Set Conn=nothing
%>
</body>
</html>






