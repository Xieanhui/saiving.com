<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->

<%
Dim Conn,User_Conn
MF_Default_Conn
MF_User_Conn
MF_Session_TF
If  Request.Form("action") = "add" then
    Dim P_UserAdd_Obj,P_UserAdd_Sql,P_ChooseME_NameObj,P_MeName_Str
	If NoSqlHack(Request.Form("UserName"))="" or isnull(Request.Form("UserName")) then
		Response.Write("<script>alert(""����д��Ա��¼��"");</script>")
		Response.End
	Else
	End If

	If len(Request.Form("UserName"))>10 then
		Response.Write("<script>alert(""��Ա��¼�������Գ���10���ַ�"");</script>")
		Response.End
	Else
	end if
'by xq
		P_MeName_Str = Replace(Replace(Request.Form("Name"),"""",""),"'","")
	
	Set P_ChooseME_NameObj = User_Conn.Execute("Select UserID from FS_Me_Users where UserName='"&P_MeName_Str&"'")
	If Not P_ChooseME_NameObj.eof then
		Response.Write("<script>alert(""�˻�Ա��¼���Ѿ�����,���޸�"");</script>")
		Response.End
	End If
	P_ChooseME_NameObj.Close
	Set P_ChooseME_NameObj = Nothing
	If Request.Form("Password")="" or isnull("Password") then
		Response.Write("<script>alert(""�������Ա��¼����"");</script>")
		Response.End
	End If
	If Len(Request.Form("Password")) < 6 then
		Response.Write("<script>alert(""��Ա��¼���벻��������λ"");</script>")
		Response.End
	End If
	If Cstr(Request.Form("Password"))<>Cstr(Request.Form("PasswordTF")) then
		Response.Write("<script>alert(""������ȷ�����벻ͬ"");</script>")
		Response.End
	End If
	If Request.Form("Name")="" or isnull(Request.Form("Name")) then
		Response.Write("<script>alert(""����д��Ա��ʵ����"");</script>")
		Response.End
	End If
'-------------------
	Dim Fs_User,NumShopPoint
	Set Fs_User = New Cls_User
	NumShopPoint = Fs_User.getUserConfig(6)
	Randomize 
	Dim RandomFigure
	RandomFigure = CStr(Int((9999 * Rnd) + 1))

'----------------------------------
	Set P_UserAdd_Obj = Server.CreateObject(G_FS_RS)
		P_UserAdd_Sql = "Select * from FS_Me_Users where 1=0"
		P_UserAdd_Obj.Open P_UserAdd_Sql,User_Conn,3,3
		P_UserAdd_Obj.AddNew
		P_UserAdd_Obj("UserName") = NoSqlHack(Replace(Request.Form("UserName"),"""",""))
		P_UserAdd_Obj("Password") = md5(Request.Form("Password"),16)
		P_UserAdd_Obj("GroupID") = NoSqlHack(Request.Form("GroupID"))
		P_UserAdd_Obj("Name") = NoSqlHack(Replace(Request.Form("Name"),"""",""))
		If Request.Form("Lock") = "0" then
			P_UserAdd_Obj("IsLock") = "0"
		Else
			P_UserAdd_Obj("IsLock") = "1"
		End If
		If Request.Form("Sex") = "0" then
			P_UserAdd_Obj("Sex") = "0"
		Else
			P_UserAdd_Obj("Sex") = "1"
		End If
		P_UserAdd_Obj("RegTime") = Now()
		P_UserAdd_Obj("Email")="Foosun@foosun.cn"
		P_UserAdd_Obj("LastLoginIP") = NoSqlHack(Request.ServerVariables("REMOTE_ADDR"))
		P_UserAdd_Obj("LastLoginTime") = Now()
		P_UserAdd_Obj("LoginNum") = 0
		P_UserAdd_Obj("Integral") = 100 '��Ա���� 
		P_UserAdd_Obj("UserNumber") = year(now)&month(now)&day(now)&hour(now)&RandomFigure '��Ա���
		P_UserAdd_Obj.Update
		P_UserAdd_Obj.Close
		Set P_UserAdd_Obj = Nothing
		Response.Redirect("Usermanage.asp")
		Response.End
End If
%>
<html>
<HEAD>
<TITLE>��ӻ�Ա</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">

<body leftmargin="2" topmargin="2">
<form action="" method="post" name="UserAddSForm">
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="table">
  <tr> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="3" cellspacing="1">
        <tr class="hback">
          <td width=35 align="center" alt="����" onClick="document.UserAddSForm.submit();" >����</td>
		  <td width=2 >|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" >����</td>
          <td>&nbsp;<input name="action" type="hidden" id="action" value="add"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%"  border="0" cellpadding="3" cellspacing="1" class="table">
    <tr> 
      <td width="100"> 
        <div align="right">�� Ա ��</div></td>
      <td colspan="3" > 
        <input name="Name" type="text"  id="Name" style="width:100%" value="<%=Request("UserName")%>"></td>
    </tr>
    <tr > 
      <td > 
        <div align="right">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</div></td>
      <td colspan="3"> 
        <input name="Password" type="password" id="Password" style="width:1090%"></td>
    </tr>
    <tr> 
      <td> 
        <div align="right">ȷ������</div></td>
      <td colspan="3"> 
        <input name="PasswordTF" type="password" id="PasswordTF" style="width:100%"></td>
    </tr>
    <tr> 
      <td> 
        <div align="right">�� Ա ��</div></td>
      <td colspan="3"> 
        <select name="GroupID" id="GroupID" style="width:100%">
          <option value="0" <%If Request("GroupID") = "" or  Request("GroupID") = "0" then Response.Write("selected") end if%>> 
          </option>
          <%
		Dim P_Set_GroupObj
		Set P_Set_GroupObj = User_Conn.Execute("Select GroupID,GroupName from FS_ME_Grouporder by GroupNamedesc")
		do while not P_Set_GroupObj.eof 
	%>
          <option value="<%=P_Set_GroupObj("GroupID")%>" <%If Cstr(Request("GroupID"))=Cstr(P_Set_GroupObj("GroupID")) then Response.Write("selected") end if%>><%=P_Set_GroupObj("GroupName")%></option>
          <%
		P_Set_GroupObj.MoveNext
		Loop
		P_Set_GroupObj.Close
		Set P_Set_GroupObj = Nothing
	%>
        </select></td>
    </tr>
    <tr> 
      <td> 
        <div align="right">��ʵ����</div></td>
      <td colspan="3"> 
        <input name="Name" type="text" id="Name" size="20" style="width:100%" value="<%=Request("Name")%>"></td>
    </tr>
    <tr valign="middle"> 
      <td> 
        <div align="right">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</div></td>
      <td> 
        <input type="radio" name="Lock" value="1" <%If Request("IsLock") = "1" then Response.Write("checked") end if%>>
        �� 
        <input name="Lock" type="radio" value="0" <%If Request("IsLock") = "0" or Request("Lock") = "" then Response.Write("checked") end if%>>
        ��</td>
    </tr>
    <tr> 
      <td> 
        <div align="right">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</div></td>
      <td> 
        <input name="Sex" type="radio" value="0" <%If Request("Sex") = "0" or Request("Sex") = "" then Response.Write("checked") end if%>>
        �� 
        <input type="radio" name="Sex" value="1" <%If Request("Sex") = "1" then Response.Write("checked") end if%>>
        Ů</td>
    </tr>
</table>
</form>
</body>
</html>






