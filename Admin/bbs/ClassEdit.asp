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
if not MF_Check_Pop_TF("WS002") then Err_Show
%>
<html>
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="javascript">
function checknum(obj){
if (isNaN(obj.value)){
alert("��ʽ����!");
obj.focus();
}
}
</script>
<BODY>
<%
Dim ClassName,ClassExp,Pid,ID,Rs,strShowErr
Set Rs=Server.createobject(G_FS_RS)
 If NoSqlHack(Request.QueryString("Act"))="Edit" Then
    ID=Request.form("ID")
 	ClassName=NoHtmlHackInput(NoSqlHack(Request.form("ClassName")))
	ClassExp=NoHtmlHackInput(NoSqlHack(Request.form("ClassExp")))
	Pid=Request.Form("pid")
	If ClassName="" Then
		strShowErr = "<li>��Ŀ���Ʋ���Ϊ��</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	End If
		if not Isnumeric(Pid) then
		strShowErr = "<li>��Ŀ������д����</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	else
		Pid=iNT(AbS(Pid))
	end if	
	Conn.execute("Update FS_WS_Class Set ClassName='"&NoSqlHack(ClassName)&"',ClassExp='"&NoSqlHack(ClassExp)&"',Pid='"&NoSqlHack(Pid)&"' Where ID="&CintStr(ID)&"")
		strShowErr = "<li>�޸ĳɹ�</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=BBS/ClassManager.asp")
		Response.end
 End If
ID=Request.QueryString("ID")
If ID="" then
	Response.write("<script>alert('��������');</script>")
	Response.End()
Else
	Rs.open "Select ID,ClassName,ClassExp,Pid From FS_WS_Class where ID="&CintStr(ID)&"",Conn,1,1
	if not Rs.eof then
%>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr>
    <td align="left" colspan="2" class="xingmu">����ϵͳ�������&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="../../help?Lable=Message" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
  </tr>
  <tr>
    <td ><a href="ClassManager.asp">������ҳ</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="ClassAdd.asp">�����Ŀ</a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table"> 
<form name="ClassAdd" action="?Act=Edit" method="post">
<tr>
	<td class="hback" width="10%">��Ŀ����</td>
	<td class="hback" width="90%"><input type="text" name="ClassName" value="<%=Rs("ClassName")%>" size="40">
	<font color="red">*</font></td>
</tr>
<tr>
	<td class="hback">��Ŀ˵��</td>
	<td class="hback"><input type="text" name="ClassExp" value="<%=Rs("ClassExp")%>" size="40">
	<font color="red">*</font></td>
</tr>
<tr>
	<td class="hback">��Ŀ����</td>
	<td class="hback">
	  <input name="pid" type="text" onKeyUp="checknum(this)" value="<%=Rs("Pid")%>" size="40">
      <font color="red">*</font>(����Խ��,��Ŀ����Խ��ǰ)</td>
</tr>
<tr>
	<td class="hback"><input type="hidden"  name="ID" value="<%=Rs("ID")%>"></td>
	<td class="hback"><input type="Submit" name="sumbit" value="����">&nbsp;&nbsp;<input type="Reset" name="reset" value="����"></td>
</tr>
</form></table>

<%
	end if
Set Rs=nothing
End if
 Set Conn=nothing
%>
</body>
</html>






