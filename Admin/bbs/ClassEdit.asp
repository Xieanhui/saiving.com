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
'session判断
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
alert("格式不对!");
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
		strShowErr = "<li>栏目名称不能为空</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	End If
		if not Isnumeric(Pid) then
		strShowErr = "<li>栏目排序填写不对</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"")
		Response.end
	else
		Pid=iNT(AbS(Pid))
	end if	
	Conn.execute("Update FS_WS_Class Set ClassName='"&NoSqlHack(ClassName)&"',ClassExp='"&NoSqlHack(ClassExp)&"',Pid='"&NoSqlHack(Pid)&"' Where ID="&CintStr(ID)&"")
		strShowErr = "<li>修改成功</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=BBS/ClassManager.asp")
		Response.end
 End If
ID=Request.QueryString("ID")
If ID="" then
	Response.write("<script>alert('参数错误');</script>")
	Response.End()
Else
	Rs.open "Select ID,ClassName,ClassExp,Pid From FS_WS_Class where ID="&CintStr(ID)&"",Conn,1,1
	if not Rs.eof then
%>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr>
    <td align="left" colspan="2" class="xingmu">留言系统分类管理&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="../../help?Lable=Message" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
  </tr>
  <tr>
    <td ><a href="ClassManager.asp">管理首页</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="ClassAdd.asp">添加栏目</a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table"> 
<form name="ClassAdd" action="?Act=Edit" method="post">
<tr>
	<td class="hback" width="10%">栏目名称</td>
	<td class="hback" width="90%"><input type="text" name="ClassName" value="<%=Rs("ClassName")%>" size="40">
	<font color="red">*</font></td>
</tr>
<tr>
	<td class="hback">栏目说明</td>
	<td class="hback"><input type="text" name="ClassExp" value="<%=Rs("ClassExp")%>" size="40">
	<font color="red">*</font></td>
</tr>
<tr>
	<td class="hback">栏目排序</td>
	<td class="hback">
	  <input name="pid" type="text" onKeyUp="checknum(this)" value="<%=Rs("Pid")%>" size="40">
      <font color="red">*</font>(数字越大,栏目排序越靠前)</td>
</tr>
<tr>
	<td class="hback"><input type="hidden"  name="ID" value="<%=Rs("ID")%>"></td>
	<td class="hback"><input type="Submit" name="sumbit" value="保存">&nbsp;&nbsp;<input type="Reset" name="reset" value="重置"></td>
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






