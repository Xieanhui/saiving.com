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
obj.value="0";
obj.focus();
}
}
</script>
<body>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table"> 
  <tr> 
    <td align="left" colspan="2" class="xingmu">����ϵͳ�������&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="../../help?Lable=Message" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td> 
  </tr>
  <tr>
  	<td ><a href="ClassManager.asp">������ҳ</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="ClassAdd.asp">�����Ŀ</a></td>
  </tr>
 </table>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table"> 
<form name="ClassAdd" action="?Act=Add" method="post">
<tr>
	<td class="hback" width="10%">��Ŀ����</td>
	<td class="hback" width="90%"><input type="text" name="ClassName" size="40">
	<font color="red"> *</font></td>
</tr>
<tr>
	<td class="hback">��Ŀ˵��</td>
	<td class="hback"><input type="text" name="ClassExp"	 size="40">
	<font color="red"> *</font></td>
</tr>
<tr>
	<td class="hback">��Ŀ����</td>
	<td class="hback"><input name="pid" type="text" onKeyUp="checknum(this)" value="0" size="40">
	  <font color="red">*</font>(����Խ��,��Ŀ����Խ��ǰ)</td>
</tr>
<tr>
	<td class="hback">&nbsp;</td>
	<td class="hback"><input type="Submit" name="sumbit" value="���">&nbsp;&nbsp;<input type="Reset" name="reset" value="����"></td>
</tr>
</form>
</table>

<%
Dim ClassName,ClassExp,Pid,strShowErr
 If NoSqlHack(Request.QueryString("Act"))="Add" Then
 	ClassName=NoHtmlHackInput(NoSqlHack(Replace(Request.form("ClassName"),"'","")))
	ClassExp=NoHtmlHackInput(NoSqlHack(Replace(Request.form("ClassExp"),"'","")))
	Pid=trim(Request.Form("pid"))
	If ClassName="" Then
		strShowErr = "<li>��Ŀ���Ʋ���Ϊ��!</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	End If
	if not Isnumeric(Pid) then
		strShowErr = "<li>��Ŀ������д����!</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	else
		Pid=iNT(AbS(Pid))
	end if	
	Conn.execute("Insert into FS_WS_Class(ClassID,ClassName,ClassExp,Pid,Author) values('"&GetRamCode(15)&"','"&NoSqlHack(ClassName)&"','"&NoSqlHack(ClassExp)&"','"&NoSqlHack(Pid)&"','"&session("Admin_Name")&"')")
		strShowErr = "<li>��ӳɹ�</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=BBS/ClassManager.asp")
		Response.end
 End If
 Set Conn=nothing
%>
</body>
</html>






