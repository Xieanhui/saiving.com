<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Md5.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="../API/Cls_PassportApi.asp" -->
<%
	User_GetParm
	response.Cookies("FoosunUserCookies")("UserLogin_Style_Num")  = p_LoginStyle
	If p_LoginStyle="" Or p_LoginStyle = 0 then
		Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num") = "1"
	End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title>��ӭ�û�<%=Session("FS_UserName")%>����<%=GetUserSystemTitle%>-�һ�����--��Ա��½</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="keywords" content="��Ѷ,��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head>
<body oncontextmenu="//return false;">
<%
If Request.Form("Action")  = "step1" then
	Call step1()
ElseIf Request.Form("Action")  = "step2" then
	Call step2()
ElseIf Request.Form("Action")  = "step3" then
	Call step3()
Else
   Call Main()
End if
%>
<%Sub main()%>
<table width="90%" height="145" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <form action=""  method="post" name="myform" id="myform" >
    <tr class="back"> 
      <td   colspan="2" class="xingmu" height="24">�һ������һ��</td>
    </tr>
    <tr class="back"> 
      <td width="23%" height="72" class="hback"> <div align="right">�û��� </div></td>
      <td width="77%" class="hback"><input name="UserName" type="text" id="UserName" style="width:160px"  /> 
        <input class="button" type="submit" value="�һ������һ��" name="Submit" /> <input name="Action" type="hidden" id="Action" value="step1"> 
      </td>
    </tr>
    <tr class="back"> 
      <td height="26"  colspan="2" class="xingmu"> <div align="left"> 
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td width="53%" class="xingmu"  height="24">FoosunCMS<%=p_Soft_Version %></td>
              <td width="47%" class="xingmu"  height="24">Powered by <a href="http://www.foosun.net" target="_blank" title="��ѶCMS---��վ���ݹ���ר��,Www.foosun.cn">Foosun 
                Inc.</a></td>
            </tr>
          </table>
        </div></td>
    </tr>
  </form>
</table>
<%End Sub%>
<%
Sub step1()
Dim p_UserName,RsStep1Obj
p_UserName = NoSqlHack(Replace(Request.Form("UserName"),"''",""))
if p_UserName = "" then
	strShowErr = "<li>����д�û���</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../GetPassword.asp")
	Response.end
End if
Set RsStep1Obj = server.CreateObject(G_FS_RS)
RsStep1Obj.open "select  UserName,UserID,PassQuestion From FS_ME_Users where UserName = '"& p_UserName &"'",User_Conn,1,1
if RsStep1Obj.eof then
	strShowErr = "<li>�����ڴ��û���</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../GetPassword.asp")
	Response.end
Else
%>
<table width="90%" height="195" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <form action=""  method="post" name="myform" id="myform" >
    <tr class="back"> 
      <td   colspan="2" class="xingmu" height="24">�һ�����ڶ���</td>
    </tr>
    <tr class="back"> 
      <td height="25" class="hback"> <div align="right">������������:</div></td>
      <td class="hback"><input name="PassQuestion" type="text" id="PassQuestion" style="width:160px" value="<% = RsStep1Obj("PassQuestion")%>" readonly /></td>
    </tr>
    <tr class="back"> 
      <td height="27" class="hback"> <div align="right">��д���������</div></td>
      <td class="hback"><input name="PassAnswer" type="text" id="PassAnswer" style="width:160px" /></td>
    </tr>
    <tr class="back"> 
      <td height="31" class="hback"> <div align="right">��ȫ��</div></td>
      <td class="hback"><input name="safeCode" type="text" id="safeCode" style="width:160px"/>
        �����ͨ����վ��̳�򲩿�����ĳ�ʼ�Ǻ������һ��</td>
    </tr>
    <tr class="back">
      <td height="27" class="hback"><div align="right">���ĵ����ʼ�</div></td>
      <td class="hback"><input name="Email" type="text" id="Email" style="width:160px"/>
        ������д</td>
    </tr>
    <tr class="back"> 
      <td width="23%" height="27" class="hback"> <div align="right"></div></td>
      <td width="77%" class="hback"> <input class="button" type="submit" value="�һ�����ڶ���" name="Submit2" /> 
        <input name="Action" type="hidden" id="Action" value="step2">
        <input name="UserName" type="hidden" id="UserName" value="<% = RsStep1Obj("UserName")%>">
        <span class="tx"> ˵��������𰸺Ͱ�ȫ��������дһ���</span></td>
    </tr>
    <tr class="back"> 
      <td height="26"  colspan="2" class="xingmu"> <div align="left"> 
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td width="53%" class="xingmu"  height="24">FoosunCMS<%=p_Soft_Version %></td>
              <td width="47%" class="xingmu"  height="24">Powered by <a href="http://www.foosun.net" target="_blank" title="��ѶCMS---��վ���ݹ���ר��,Www.foosun.cn">Foosun 
                Inc.</a></td>
            </tr>
          </table>
        </div></td>
    </tr>
  </form>
</table>
<%
	End if
RsStep1Obj.close
set RsStep1Obj = nothing
End Sub
%>
<%Sub step2()
Dim p_UserName_str,RsStep2Obj,p_PassAnswer,p_safeCode,p_Email,SQL,p_PassQuestion
p_UserName_str = NoSqlHack(Replace(Trim(Request.Form("UserName")),"''",""))
p_PassAnswer = NoSqlHack(Replace(Trim(Request.Form("PassAnswer")),"''",""))
p_PassQuestion = NoSqlHack(Replace(Trim(Request.Form("PassQuestion")),"''",""))
p_safeCode = NoSqlHack(Replace(Trim(Request.Form("safeCode")),"''",""))
p_Email = NoSqlHack(Replace(Trim(Request.Form("Email")),"''","")) 
if p_PassAnswer = "" and p_safeCode = "" then
	strShowErr = "<li>����𰸻������������дһ��</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
End if
if p_Email = ""  then
	strShowErr = "<li>����д�����ʼ�</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
End if
Set RsStep2Obj = server.CreateObject(G_FS_RS)
if p_PassAnswer <>"" then
	SQL = "select  PassAnswer,safeCode,UserName,UserID,PassQuestion,Email From FS_ME_Users where UserName = '"& NoSqlHack(p_UserName_str) &"' and PassAnswer = '"& NoSqlHack(md5(p_PassAnswer,16))  &"' and Email='"& NoSqlHack(p_Email) &"'"
Else
	SQL =  "select  PassAnswer,safeCode,UserName,UserID,PassQuestion,Email From FS_ME_Users where UserName = '"& NoSqlHack(p_UserName_str) &"' and safeCode = '"& NoSqlHack(md5(p_safeCode,16))  &"'  and Email='"& NoSqlHack(p_Email) &"'"
End if
'Response.Write(SQL)
'Response.end
RsStep2Obj.open SQL,User_Conn,1,1
if RsStep2Obj.eof then
	strShowErr = "<li>�Ҳ�����¼����ȷ����������ȷ</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
Else
%>
<table width="90%" height="137" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <form action=""  method="post" name="myform" id="myform" >
    <tr class="back"> 
      <td   colspan="2" class="xingmu" height="24">�޸�����</td>
    </tr>
    <tr class="back">
      <td height="38" class="hback"> 
        <div align="right">����������������</div></td>
      <td class="hback"><input name="pass_new" type="password" id="pass_new" style="width:160px"  />
        <input type="hidden" name="UserName" value="<% = p_UserName_str %>">
		<input type="hidden" name="PassQuestion" value="<% = p_PassQuestion %>">
		<input type="hidden" name="PassAnswer" value="<% = p_PassAnswer %>">
        <input type="hidden" name="Email" value="<% = p_Email %>"></td>
    </tr>
    <tr class="back"> 
      <td width="23%" height="42" class="hback"> 
        <div align="right">����������������</div></td>
      <td width="77%" class="hback"><input name="confim_pass_new" type="password" id="confim_pass_new" style="width:160px"  /> 
        <input class="button" type="submit" value="�һ�����" name="Submit3" /> 
        <input name="Action" type="hidden" id="Action" value="step3"> </td>
    </tr>
    <tr class="back"> 
      <td height="26"  colspan="2" class="xingmu"> <div align="left"> 
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td width="53%" class="xingmu"  height="24">FoosunCMS<%=p_Soft_Version %></td>
              <td width="47%" class="xingmu"  height="24">Powered by <a href="http://www.foosun.net" target="_blank" title="��ѶCMS---��վ���ݹ���ר��,Www.foosun.cn">Foosun 
                Inc.</a></td>
            </tr>
          </table>
        </div></td>
    </tr>
  </form>
</table>
<%
End if
 RsStep2Obj.close
 set RsStep2Obj = nothing
End Sub%>
<%
Sub step3()
Dim p_pass_new,p_confim_pass_new
p_pass_new = md5(Request.Form("pass_new"),16)
p_confim_pass_new = md5(Request.Form("confim_pass_new"),16)
if NoSqlHack(Replace(Request.Form("pass_new"),"''","")) = ""  then
	strShowErr = "<li>����д��������</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
End if
if NoSqlHack(Replace(Request.Form("pass_new"),"''","")) <> NoSqlHack(Replace(Request.Form("confim_pass_new"),"''",""))  then
	strShowErr = "<li>2�����벻һ��</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
End if
Dim StrUserName,getoldnamepwdrs,oldpassword
StrUserName = NoSqlHack(Replace(Request.Form("UserName"),"''",""))
set getoldnamepwdrs = User_Conn.execute("select userpassword from [FS_ME_Users] where username='"&StrUserName&"'") 
if not getoldnamepwdrs.eof then
	oldpassword = getoldnamepwdrs("userpassword")
end if
getoldnamepwdrs.close

User_Conn.execute("Update FS_ME_Users set UserPassword ='"& NoSqlHack(p_pass_new) &"' where UserName = '"& NoSqlHack(StrUserName) &"' and Email = '"& NoSqlHack(Replace(Request.Form("Email"),"''",""))&"'")

	'-----------------------------------------------------------------
	'ϵͳ����
	'-----------------------------------------------------------------
	Dim API_Obj,API_SaveCookie,SysKey
	If API_Enable Then
		Set API_Obj = New PassportApi
			API_Obj.NodeValue "action","update",0,False
			API_Obj.NodeValue "username",StrUserName,1,False
			API_Obj.NodeValue "question","",1,False
			API_Obj.NodeValue "answer","",1,False
			API_Obj.NodeValue "email","",1,False
			SysKey = Md5(API_Obj.XmlNode("username")&API_SysKey,16)
			API_Obj.NodeValue "syskey",SysKey,0,False
			API_Obj.NodeValue "password",NoSqlHack(Request.Form("pass_new")),0,False
			API_Obj.SendHttpData
			If API_Obj.Status = "1" Then
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(API_Obj.Message)&"&ErrorUrl=")
				Response.end
			End If
		Set API_Obj = Nothing
	End If
	'-----------------------------------------------------------------
	strShowErr = "<li>�޸ĳɹ��������µ�¼��</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../login.asp")
	Response.end
End Sub
%>
</body>
</html>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->
<script language="JavaScript" type="text/javascript">
function SetFocus()
{
if (document.myform.name.value=="")
	document.myform.name.focus();
else
	document.myform.name.select();
}
function CheckForm()
{
	if(document.myform.name.value=="")
	{
		alert("�����������û�����");
		document.myform.name.focus();
		return false;
	}
	if(document.myform.password.value == "")
	{
		alert("�������������룡");
		document.myform.password.focus();
		return false;
	}
}
</script>







