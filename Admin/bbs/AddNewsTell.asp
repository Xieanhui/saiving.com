<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/ns_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn
MF_Default_Conn
'session�ж�
MF_Session_TF 
Dim TellRs,Tellsql
if not MF_Check_Pop_TF("WS002") then Err_Show

dim Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
%>
<html>
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<style type="text/css">
<!--
.style1 {font-weight: bold}
.style2 {color: #FF0000}
-->
</style>
</HEAD>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY>
  <table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
    <tr>
      <td align="left" colspan="2" class="xingmu">�������&nbsp;&nbsp;&nbsp;<a href="../../help?Lable=Message" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
    </tr>
	<tr class="hback">
  	<Td class="hback" colspan="4"><a href="NewTell.asp">����</a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
<form action="?Act=Add" method="post" name="Addform">
	<tr>
		<td class="hback" width="10%" align="right">�� �� ��</td>
		<td class="hback" width="80%"><input type="text" id="AddUser" name="AddUser" size="40" value="<%=Temp_Admin_Name%>" readonly></td>
	</tr>
	<tr>
		<td class="hback" align="right">�������</td>
	  <td class="hback"><input name="Topic" type="text" id="Topic" size="40" maxlength="50">
	    <font color="#FF0000">*������д��Ŀ</font></td>
	</tr>
	<tr>
		<td class="hback" align="right"> ��  ��  </td>
	  <td class="hback"><textarea name="Content" cols="70" rows="7"></textarea>
		  <font color="#FF0000">*������д��Ŀ</font>
	<tr>
		<td class="hback" >&nbsp;</td>
	  <td class="hback">&nbsp;
		  <input type="submit" name="submit" value="��  ��">		  &nbsp;&nbsp;
	      <input type="reset" name="reset" value="��  ��">
</form>
	</table>
<%
Dim AddUser,Topic,Content,Person,AddRs,strShowErr
Set AddRs=Server.createobject(G_FS_RS)
if Request.querystring("Act")="Add" then
	AddUser = NoSqlHack(Request.form("AddUser"))
	Topic   = NoSqlHack(Request.form("Topic"))
	Content = NoSqlHack(Replace(Request.form("Content"),vbcrlf,"<br>"))
	Person  = NoSqlHack(request.form("Person"))
	if AddUser="" then 
		strShowErr = "<li>�û���Ϊ��,�����û���½�Ƿ��ѹ���!</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	end if
	if Topic="" then
		strShowErr = "<li>������ⲻ��Ϊ��</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	end if
	if Content="" then
		strShowErr = "<li>�������ݲ���Ϊ��</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	end if	
	AddRs.open "Select Topic,Content,Person,AddUser,AddDate From FS_WS_NewsTell where 1=2",Conn,1,3
	AddRs.Addnew
	AddRs("Topic")  = Topic
	AddRs("Content")= Content
	AddRs("Person") = Person
	AddRs("AddUser")= AddUser
	AddRs("AddDate")=now()
	AddRs.update
	Set AddRs=nothing
		strShowErr = "<li>��ӳɹ�</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=BBS/NewTell.asp")
		Response.end
end if
Set Conn=nothing
%>
</body>
</html>






