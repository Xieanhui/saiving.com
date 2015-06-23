<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
'==============================================================================
'������ƣ���Ѷ��վ��Ϣ����ϵͳ
'��ǰ�汾��Foosun Content Manager System(FoosunCMS V4.6)
'������ʱ�䣺2008.09.22
'==============================================================================
Dim ObjInstalled,NewsID,Action,rs

Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

Function SendMail(SMTPServer,loginName,LoginPass,NameSendFrom,EmailSendFrom,StrSendTo,StrSubject,StrContent)
	'On error resume next
	Dim ObjJmail,ArrSendTo,i
	If InStr(StrSendTo,",")>0 Then 
		ArrSendTo = Split(StrSendTo,",")
	Else
		ArrSendTo = Array(StrSendTo)
	End If 
	Set ObjJmail = Server.CreateObject(G_JMAIL_MESSAGE) 
	ObjJmail.Silent = True
	ObjJmail.Logging = True
	ObjJmail.Charset = "gb2312" 
	ObjJmail.MailServerUserName = LoginName 
	ObjJmail.MailServerPassword = LoginPass 
	ObjJmail.ContentType = "text/html" 
	ObjJmail.From = EmailSendFrom
	ObjJmail.FromName = NameSendFrom
	ObjJmail.Subject = StrSubject
	For i=LBound(ArrSendTo) To UBound(ArrSendTo)
		ObjJmail.AddRecipient ArrSendTo(i)
	Next 
	ObjJmail.Body = StrContent 
	ObjJmail.Priority = 3  '�ʼ������ȼ������Է�Χ��1��5��Խ������ȼ�Լ��
	ObjJmail.AddHeader "Originating-IP", Request.ServerVariables("REMOTE_ADDR")
	if not ObjJmail.Send(SMTPServer) then
		SendMail = false
'		Response.Write("�ʼ�����ʧ�ܣ������Ƿ�������֧��JMAIL�������ʹ��jmail4.3���ϰ汾��<br>")
	Else
		SendMail = true
'		Response.Write("�ʼ��Ѿ����͵���ע��������У���ע�����<br>")
	End If
	ObjJmail.Close
	Set ObjJmail=nothing   
End Function
'----

function IsValidEmail(email)
	Dim regEx
	Set regEx = New RegExp
	regEx.Pattern = "(\w|-|_|0-9|\.| )+@{1}(\w|0-9|\.|-)+\.[A-Za-z]{2,3}"
	regEx.IgnoreCase = True
	IsValidEmail = regEx.Test(email)
	Set regEx=Nothing
end function

ObjInstalled=IsObjInstalled("JMail.SMTPMail")
Newsid= request("id")
Action=trim(request("Action"))
if Trim(Newsid)="" then
	Response.write"<script>alert(""����Ĳ�����"");history.back();</script>"
    Response.end
end if

	sql="Select * from FS_NS_News where Newsid='"&NoSqlHack(Newsid)&"'"
	set rs=server.createobject(G_FS_RS)
	rs.open sql,conn,1,3
	if rs.bof and rs.eof then
		Response.write"<script>alert(""�Ҳ������ţ�"");history.back();</script>"
		Response.end
	else
		if Action="MailToFriend" then
			call MailToFriend()
		else
			call main()
		end if
	end if
	rs.close
	set rs=nothing
sub main()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%=GetUserSystemTitle%>-���͵����ʼ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">

<body>
<table width="95%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
  <tr> 
    <td bgcolor="#FFFFFF">
<TABLE width="100%" border=0 cellpadding="6">
        <TBODY>
          <TR> 
            <TD width=26><IMG src="images/GroupUser.gif" border=0></TD>
            <TD class=f4>���͵����ʼ�</TD>
          </TR>
        </TBODY>
      </TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
          <TR> 
            <TD bgColor=#ff6633 height=4><IMG height=1 src="" width=1></TD>
          </TR>
        </TBODY>
      </TABLE></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF">
<form name="form1" method="post" action="">
        <table cellpadding=6 cellspacing=1 border=0 width=90% class="border" align=center>
          <tr> 
      <td height="22" colspan=2 align=center valign=middle class="title"> <b>�����ĸ��ߺ���</b></td>
    </tr>
    <tr class="tdbg"> 
      <td width="120" align="right"><strong>������������</strong></td>
      <td><input name="MailtoName" type="text" id="MailtoName" size="60" maxlength="20"> 
        <font color="#FF0000">*</font></td>
    </tr>
    <tr class="tdbg"> 
      <td width="120" align="right"><strong>������Email��ַ��</strong></td>
      <td><input name="MailToAddress" type=text id="MailToAddress" size="60" maxlength="100"> 
        <font color="#FF0000">*</font></td>
    </tr>
    <tr class="tdbg"> 
      <td height="20" align="right"><strong>���������</strong></td>
      <td height="20"> <input name="Username" type=text id="Username" value="<% =Fs_User.UserName%>" size="60" maxlength="100"></td>
    </tr>
    <tr class="tdbg"> 
      <td height="20" align="right"><strong>���Email��ַ��</strong></td>
      <td height="20"><input name="Useremail" type=text id="Useremail" value="<% =Fs_User.Email%>" size="60" maxlength="100"></td>
    </tr>
    <tr class="tdbg"> 
      <td width="120" height="60" align="right"><strong>������Ϣ��</strong></td>
      <td height="60">���ű��⣺<font color="#FF0000"><strong><%= rs("NewsTitle") %></strong></font><br>
        �������ߣ�<%= rs("Author") %> <br>
        ����ʱ�䣺<%= rs("addtime") %> </td>
    </tr>
    <tr class="tdbg"> 
      <td colspan=2 align=center><input name="Action" type="hidden" id="Action" value="MailToFriend"> 
        <input name="filename" type="hidden" id="Newsid" value="<%=request("Newsid")%>"> 
        <input type=submit value=" �� �� " name="Submit" <% If ObjInstalled=false Then response.write "disabled" end if%>> 
      </td>
    </tr>
    <%
If ObjInstalled=false Then
	Response.Write "<tr><td height='40' colspan='2'><b><font color=red>�Բ�����Ϊ��������֧�� JMail���! ���Բ���ʹ�ñ����ܡ�</font></b></td></tr>"
End If
%>
  </table>
</form>
    </td>
  </tr>
  <tr>
    <td bgcolor="#F2F2F2"> 
      <div align="center">
       <!--#include file="Copyright.asp" -->
      </div></td>
  </tr>
</table>
</body>
</html>
<%end sub
sub MailToFriend()
	Dim MailToName,MailToAddress
	'==============================================================================
	'���� FSϵͳ�ʼ�����
	Dim MailCfg,MF_Domain,MF_Site_Name,MF_eMail,MF_Mail_Server,MF_Mail_Name,MF_Mail_Pass_Word
	set MailCfg = Conn.execute("select top 1 MF_Domain,MF_Site_Name,MF_eMail,MF_Mail_Server,MF_Mail_Name,MF_Mail_Pass_Word from FS_MF_Config")
	if MailCfg.eof then
		response.Write "<script>alert('�Ҳ���������Ϣ������ϵͳ����Ա��ϵ.\n����ϵͳ��Ӧ����ϵ����������á�by Foosun.CN');window.history.back();</script>"
		response.end
		MailCfg.close:set MailCfg=nothing
	else
		MF_Domain=MailCfg("MF_Domain")
		MF_Site_Name=MailCfg("MF_Site_Name")
		MF_eMail=MailCfg("MF_eMail")
		MF_Mail_Server=MailCfg("MF_Mail_Server")
		MF_Mail_Name=MailCfg("MF_Mail_Name")
		MF_Mail_Pass_Word=MailCfg("MF_Mail_Pass_Word")
		
		MailCfg.close:set MailCfg=nothing
	end if
	'===============================================================================
	MailToName=trim(request.form("MailToName"))
	MailToAddress=trim(request.form("MailToAddress"))
	if MailToName="" then
		Response.write "<script>alert(""�����˲���Ϊ�գ�"");history.back();</script>"
        Response.end
	end if
	if IsValidEmail(MailToAddress)=False then
   		Response.write "<script>alert(""EMAIL��ַ����"");history.back();</script>"
        Response.end
	end if
				
	Dim t_server,t_Name,t_Pwd,t_From,t_Efrom,t_to,t_ret,Subject,mailbody

	Subject="��������"&request.Form("Username")&"��" & MF_Site_Name & "������������������"

	mailbody=mailbody &"<style>A:visited {	TEXT-DECORATION: none	}"
	mailbody=mailbody &"A:active  {	TEXT-DECORATION: none	}"
	mailbody=mailbody &"A:hover   {	TEXT-DECORATION: underline overline	}"
	mailbody=mailbody &"A:link 	  {	text-decoration: none;}"
	mailbody=mailbody &"A:visited {	text-decoration: none;}"
	mailbody=mailbody &"A:active  {	TEXT-DECORATION: none;}"
	mailbody=mailbody &"A:hover   {	TEXT-DECORATION: underline overline}"
	mailbody=mailbody &"BODY   {	FONT-FAMILY: ����; FONT-SIZE: 9pt;}"
	mailbody=mailbody &"TD	   {	FONT-FAMILY: ����; FONT-SIZE: 9pt	}</style>"
	mailbody=mailbody &"<TABLE border=0 width='95%' align=center><TBODY><TR>"
	mailbody=mailbody &"<TD valign=middle align=top>"
	mailbody=mailbody &"--&nbsp;&nbsp;���ߣ�"&rs("Author")&"<br>"
	mailbody=mailbody &"--&nbsp;&nbsp;����ʱ�䣺"&rs("addtime")&"<br><br>"
	mailbody=mailbody &"--&nbsp;&nbsp;"&rs("NewsTitle")&"<br>"
	mailbody=mailbody &""&rs("Content")&""
	mailbody=mailbody &"</TD></TR></TBODY></TABLE>"
	mailbody=mailbody &"<center><a href='" & "Http://"&MF_Domain & "'>" & MF_Site_Name & ",�����ʼ�"&request.Form("Useremail")&"</a>"

	t_server = MF_Mail_Server
	t_Name =MF_Mail_Name
	t_Pwd = MF_Mail_Pass_Word
	t_From = NoSqlHack(request.Form("Username"))
	t_Efrom = MF_eMail
	t_to = NoSqlHack(request.Form("MailToAddress"))
'	Response.write subject & mailbody :response.End
	t_ret = SendMail(t_server,t_Name,t_Pwd,t_From,t_Efrom,t_to,Subject,mailbody)
	If t_ret=False Then
		response.Write("<script>alert('����ʧ�ܡ�\nϵͳ��������ȷ��');history.back();</script>")
		response.end
	End If 
	if Err then '���
		response.Write("<script>alert('����ʧ��\n"&err.description&"');history.back();</script>")
		Err.clear
		response.end
	else
		response.Write("<script>alert('���ͳɹ�');window.history.back();</script>")
		response.end
	end if

end sub
%>






