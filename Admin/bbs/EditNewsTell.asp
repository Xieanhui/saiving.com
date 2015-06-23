<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/ns_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn
MF_Default_Conn
'session判断
MF_Session_TF 
if not MF_Check_Pop_TF("WS002") then Err_Show
Dim TellRs,Tellsql,ID,strShowErr
%>
<html>
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY>
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="1"></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr>
    <td align="left" colspan="2" class="xingmu">公告管理&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="../../help?Lable=Message" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
  </tr>
  <tr>
    <td colspan="2" class="hback"><a href="NewTell.asp">返回</a></td>
  </tr>
</table>
<%
if NoSqlHack(Request.QueryString("Act"))="Edits" then
ID=trim(Request.QueryString("ID"))
Set TellRs=Server.createobject(G_FS_RS)
if ID="" then
		strShowErr = "<li>参数错误</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=BBS/NewTell.asp")
		Response.end
end if
TellRs.open "Select ID,Topic,Content,Person,IsUse,PV,AddUser,AddDate From FS_WS_NewsTell Where ID="&CintStr(ID)&"",Conn,1,1
	if TellRs.eof and TellRs.bof then
	Set TellRs=nothing
		strShowErr = "<li>你在修改的公告已不存了,可能已被别人删除!</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=BBS/NewTell.asp")
		Response.end
end if
end if

%>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <form action="EditNewsTellDeal.asp?Act=Edit" method="post" name="Addform">
    <tr>
      <td class="hback" width="10%" align="right">添 加 者</td>
      <td class="hback" width="80%"><input type="text" id="AddUser" name="AddUser" size="40" value="<%=TellRs("AddUser")%>" readonly></td>
    </tr>
    <tr>
      <td class="hback" align="right">公告标题</td>
      <td class="hback"><input name="Topic" type="text" id="Topic" size="40" maxlength="20" value="<%=TellRs("Topic")%>">
        <font color="#FF0000">*必须填写项目</font></td>
    </tr>
    <tr>
      <td class="hback" align="right"> 内  容 </td>
      <td class="hback"><textarea name="Content" rows="8" style="width:80%"><%=TellRs("Content")%></textarea>
        <font color="#FF0000">*必须填写项目</font>
		</td>
    <tr>
      <td class="hback" >&nbsp;
        <input type="hidden" name="ID" ID="ID" value='<%=TellRs("ID")%>'></td>
      <td class="hback">&nbsp;
        <input type="submit" name="submit" value="保  存">
        &nbsp;&nbsp;
        <input type="reset" name="reset" value="重  置">
 		</td>
		</tr>
  </form>
</table>
<%
Set Conn=nothing
%>
</body>
</html>






