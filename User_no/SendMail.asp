<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
'==============================================================================
'软件名称：风讯网站信息管理系统
'当前版本：Foosun Content Manager System(FoosunCMS V4.6)
'最后更新时间：2008.09.22
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
	ObjJmail.Priority = 3  '邮件的优先级，可以范围从1到5。越大的优先级约高
	ObjJmail.AddHeader "Originating-IP", Request.ServerVariables("REMOTE_ADDR")
	if not ObjJmail.Send(SMTPServer) then
		SendMail = false
'		Response.Write("邮件发送失败，可能是服务器不支持JMAIL组件，请使用jmail4.3以上版本！<br>")
	Else
		SendMail = true
'		Response.Write("邮件已经发送到你注册的邮箱中，请注意查收<br>")
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
	Response.write"<script>alert(""错误的参数！"");history.back();</script>"
    Response.end
end if

	sql="Select * from FS_NS_News where Newsid='"&NoSqlHack(Newsid)&"'"
	set rs=server.createobject(G_FS_RS)
	rs.open sql,conn,1,3
	if rs.bof and rs.eof then
		Response.write"<script>alert(""找不到新闻！"");history.back();</script>"
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
<title><%=GetUserSystemTitle%>-发送电子邮件</title>
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
            <TD class=f4>发送电子邮件</TD>
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
      <td height="22" colspan=2 align=center valign=middle class="title"> <b>将本文告诉好友</b></td>
    </tr>
    <tr class="tdbg"> 
      <td width="120" align="right"><strong>收信人姓名：</strong></td>
      <td><input name="MailtoName" type="text" id="MailtoName" size="60" maxlength="20"> 
        <font color="#FF0000">*</font></td>
    </tr>
    <tr class="tdbg"> 
      <td width="120" align="right"><strong>收信人Email地址：</strong></td>
      <td><input name="MailToAddress" type=text id="MailToAddress" size="60" maxlength="100"> 
        <font color="#FF0000">*</font></td>
    </tr>
    <tr class="tdbg"> 
      <td height="20" align="right"><strong>你的姓名：</strong></td>
      <td height="20"> <input name="Username" type=text id="Username" value="<% =Fs_User.UserName%>" size="60" maxlength="100"></td>
    </tr>
    <tr class="tdbg"> 
      <td height="20" align="right"><strong>你的Email地址：</strong></td>
      <td height="20"><input name="Useremail" type=text id="Useremail" value="<% =Fs_User.Email%>" size="60" maxlength="100"></td>
    </tr>
    <tr class="tdbg"> 
      <td width="120" height="60" align="right"><strong>新闻信息：</strong></td>
      <td height="60">新闻标题：<font color="#FF0000"><strong><%= rs("NewsTitle") %></strong></font><br>
        新闻作者：<%= rs("Author") %> <br>
        发布时间：<%= rs("addtime") %> </td>
    </tr>
    <tr class="tdbg"> 
      <td colspan=2 align=center><input name="Action" type="hidden" id="Action" value="MailToFriend"> 
        <input name="filename" type="hidden" id="Newsid" value="<%=request("Newsid")%>"> 
        <input type=submit value=" 发 送 " name="Submit" <% If ObjInstalled=false Then response.write "disabled" end if%>> 
      </td>
    </tr>
    <%
If ObjInstalled=false Then
	Response.Write "<tr><td height='40' colspan='2'><b><font color=red>对不起，因为服务器不支持 JMail组件! 所以不能使用本功能。</font></b></td></tr>"
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
	'加载 FS系统邮件配置
	Dim MailCfg,MF_Domain,MF_Site_Name,MF_eMail,MF_Mail_Server,MF_Mail_Name,MF_Mail_Pass_Word
	set MailCfg = Conn.execute("select top 1 MF_Domain,MF_Site_Name,MF_eMail,MF_Mail_Server,MF_Mail_Name,MF_Mail_Pass_Word from FS_MF_Config")
	if MailCfg.eof then
		response.Write "<script>alert('找不到配置信息，请与系统管理员联系.\n请与系统供应商联系导入参数设置。by Foosun.CN');window.history.back();</script>"
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
		Response.write "<script>alert(""收信人不能为空！"");history.back();</script>"
        Response.end
	end if
	if IsValidEmail(MailToAddress)=False then
   		Response.write "<script>alert(""EMAIL地址有误！"");history.back();</script>"
        Response.end
	end if
				
	Dim t_server,t_Name,t_Pwd,t_From,t_Efrom,t_to,t_ret,Subject,mailbody

	Subject="您的朋友"&request.Form("Username")&"从" & MF_Site_Name & "给您发来的新闻资料"

	mailbody=mailbody &"<style>A:visited {	TEXT-DECORATION: none	}"
	mailbody=mailbody &"A:active  {	TEXT-DECORATION: none	}"
	mailbody=mailbody &"A:hover   {	TEXT-DECORATION: underline overline	}"
	mailbody=mailbody &"A:link 	  {	text-decoration: none;}"
	mailbody=mailbody &"A:visited {	text-decoration: none;}"
	mailbody=mailbody &"A:active  {	TEXT-DECORATION: none;}"
	mailbody=mailbody &"A:hover   {	TEXT-DECORATION: underline overline}"
	mailbody=mailbody &"BODY   {	FONT-FAMILY: 宋体; FONT-SIZE: 9pt;}"
	mailbody=mailbody &"TD	   {	FONT-FAMILY: 宋体; FONT-SIZE: 9pt	}</style>"
	mailbody=mailbody &"<TABLE border=0 width='95%' align=center><TBODY><TR>"
	mailbody=mailbody &"<TD valign=middle align=top>"
	mailbody=mailbody &"--&nbsp;&nbsp;作者："&rs("Author")&"<br>"
	mailbody=mailbody &"--&nbsp;&nbsp;发布时间："&rs("addtime")&"<br><br>"
	mailbody=mailbody &"--&nbsp;&nbsp;"&rs("NewsTitle")&"<br>"
	mailbody=mailbody &""&rs("Content")&""
	mailbody=mailbody &"</TD></TR></TBODY></TABLE>"
	mailbody=mailbody &"<center><a href='" & "Http://"&MF_Domain & "'>" & MF_Site_Name & ",电子邮件"&request.Form("Useremail")&"</a>"

	t_server = MF_Mail_Server
	t_Name =MF_Mail_Name
	t_Pwd = MF_Mail_Pass_Word
	t_From = NoSqlHack(request.Form("Username"))
	t_Efrom = MF_eMail
	t_to = NoSqlHack(request.Form("MailToAddress"))
'	Response.write subject & mailbody :response.End
	t_ret = SendMail(t_server,t_Name,t_Pwd,t_From,t_Efrom,t_to,Subject,mailbody)
	If t_ret=False Then
		response.Write("<script>alert('发送失败。\n系统参数不正确。');history.back();</script>")
		response.end
	End If 
	if Err then '检测
		response.Write("<script>alert('发送失败\n"&err.description&"');history.back();</script>")
		Err.clear
		response.end
	else
		response.Write("<script>alert('发送成功');window.history.back();</script>")
		response.end
	end if

end sub
%>






