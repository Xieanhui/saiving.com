<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<%
User_GetParm
if RegisterTF =false then
	strShowErr = "<li>��ʱ�ر�ע�Ṧ��</li><li>����ϵͳ������ʧ!</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl="&Request.ServerVariables("URL")&"?"&request.QueryString&"")
	Response.end
End if
if Not isnull(DefaultGroupID) then
  if DefaultGroupID = 0 then
	strShowErr = "<li>����Ա��û����Ĭ�ϻ�Ա�顣������ʱ����ע�ᣬ�������Ա��ϵ!</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl="&Request.ServerVariables("URL")&"?"&request.QueryString&"")
	Response.end
  end if
  dim rsGroup
  set rsGroup = User_Conn.execute("select GroupID,GroupName from FS_ME_Group where GroupType=1 and GroupID="&clng(DefaultGroupID))
  if rsGroup.eof then
	strShowErr = "<li>�����쳣!</li><li>����ϵͳ�ṩ�̻�ü���֧��!!</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl="&Request.ServerVariables("URL")&"?"&request.QueryString&"")
	Response.end
  end if
  rsGroup.close:set rsGroup=nothing
else
	strShowErr = "<li>����Ա��û����Ĭ�ϻ�Ա�顣������ʱ����ע�ᣬ�������Ա��ϵ!</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl="&Request.ServerVariables("URL")&"?"&request.QueryString&"")
	Response.End
end if
response.Cookies("FoosunUserCookies")("UserLogin_Style_Num")  = p_LoginStyle
If p_LoginStyle="" Or p_LoginStyle = 0 then
	response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") = "2"
End If
Dim forward
forward = Request.QueryString("forward")
forward = Server.URLEncode(forward)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��Աע��Э��</title>
<link id="skinStyle" href="Images/Skin/css_3/3.css" rel="stylesheet" type="text/css" />
<style type="text/css">
form {
	margin: 0px;
	padding: 0px;
}
</style>
<script language="JavaScript" src="../FS_Inc/CheckJs.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../FS_Inc/Prototype.js" type="text/JavaScript"></script>
<script type="text/javascript">
	var styleCookie = unescape(document.cookie);
	var index = styleCookie.indexOf('UserLogin_Style_Num');
	if (index > -1) {
		var flag = styleCookie.indexOf('&', index);
		var flag2 = styleCookie.indexOf(';', index);
		if (flag > flag2) {
			flag = flag2;
		}
		index = styleCookie.indexOf('=', index) + 1;
		if (flag > -1) {
			styleCookie = styleCookie.substring(index, flag);
		} else {
			styleCookie = styleCookie.substring(index);
		}
	} else {
		styleCookie = '';
	}
	if (styleCookie.length > 0 && !isNaN(styleCookie)) {
		$('skinStyle').href = 'Images/Skin/css_' + styleCookie + '/' + styleCookie + '.css';
	}
</script>
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
	<tbody>
		<tr class="back">
			<td height="20" colspan="3" class="xingmu">����д�����û���<span class="tx">(������Ŀ������д)</span></td>
		</tr>
		<tr class="back">
			<td width="100%">
			<%If RegisterTF = false then%>
				<div align="center" class="tx"><p></p>
				<p>&nbsp;</p>
				<p>����Ա�Ѿ��ر�ע��!</p>
			<%Else%>
				<%=  RegisterNotice %>
			<%End if%>
			</td>
		</tr>
		<tr class="back">
			<td height="39" colspan="3" align="center">
				<input type="submit" name="Submit" value=" �ر� " onclick="window.close();" />
			</td>
		</tr>
		<tr class="back">
			<td height="26" colspan="3" class="xingmu">
				<!--#include file="Copyright.asp" -->
			</td>
		</tr>
	</tbody>
</table>
</body>
</html>




