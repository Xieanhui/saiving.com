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
Dim strUserNumber
strUserNumber = ""
response.Cookies("FoosunUserCookies")("UserLogin_Style_Num")  = p_LoginStyle
If p_LoginStyle="" Or p_LoginStyle = 0 then
	response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") = "2"
End If
Dim forward
forward = NoSqlHack(Request.QueryString("forward"))
forward = Server.URLEncode(forward)
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title>��Աע��step 2 of  4 step</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="keywords" content="��Ѷ,��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../FS_Inc/CheckJs.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../FS_Inc/Prototype.js" type="text/JavaScript"></script>
<head>
<body>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="0">
	<tr>
		<td>
			<table width="100%" height="279" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
				<tr class="back">
					<td   colspan="2" class="xingmu" height="24">��User Register step 2.(��д��Ա��������)</td>
				</tr>
				<tr class="back">
					<td width="15%" valign="top" class="hback"><strong>��ע�Ჽ�衿</strong> <br>
						<br>
						<div align="left"> ��ͬ��ע��Э��<br>
							<br>
							����д��Ա����<br>
							<br>
							����д��ϵ����<br>
							<br>
							��ע��ɹ�</div>
					</td>
					<td width="86%" valign="top" class="hback">
						<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
							<form name="UserForm"  id="UserForm" method="post" action="Reg_contact.asp?forward=<%= forward %>"  onsubmit="return CheckForm();">
								<tr class="back">
									<td height="20" colspan="3" class="xingmu">����д�����û���<span class="tx">(������Ŀ������д)</span></td>
								</tr>
								<tr class="back">
									<td width="11%" height="65">
										<div align="right">�û���</div>
									</td>
									<td width="37%">
										<input name="UserName" type="text" id="UserName" size="20"  onfocus="Do.these('UserName',function(){return isEmpty('UserName','span_UserName')})" onKeyUp="Do.these('UserName',function(){return isEmpty('UserName','span_UserName')})">
										<span id="span_UserName"></span> <a href="javascript:CheckName('lib/CheckName.asp')">����û���</a></td>
									<td width="52%">������a��z��Ӣ����ĸ(�����ִ�Сд)��0��9�����֡��㡢���Ż��»��߼�������ɣ�����Ϊ3��18���ַ���ֻ�������ֻ���ĸ��ͷ�ͽ�β,����:coolls1980��</td>
								</tr>
								<tr class="back">
									<td height="16" colspan="3" class="xingmu">����д��ȫ���ã�����ȫ����������֤�ʺź��һ����룩</td>
								</tr>
								<%If p_isValidate = 0  then%>
								<tr class="back">
									<td height="16">
										<div align="right">����</div>
									</td>
									<td>
										<input name="UserPassword" type="password" id="UserPassword" size="30" maxlength="50"  onfocus="Do.these('UserPassword',function(){return isEmpty('UserPassword','span_UserPassword')})" onKeyUp="Do.these('UserPassword',function(){return isEmpty('UserPassword','span_UserPassword')})">
										<span id="span_UserPassword"></span></td>
									<td rowspan="2">���볤��Ϊ<%=p_LenPassworMin%>��<%=p_LenPassworMax%>λ��������ĸ��Сд����¼�����������ĸ�����֡������ַ���ɡ�</td>
								</tr>
								<tr class="back">
									<td height="24">
										<div align="right">ȷ������</div>
									</td>
									<td>
										<input name="cUserPassword" type="password" id="cUserPassword" size="30" maxlength="50" onFocus="Do.these('cUserPassword',function(){return isEmpty('cUserPassword','span_cUserPassword')})" onKeyUp="Do.these('cUserPassword',function(){return isEmpty('cUserPassword','span_cUserPassword')})">
										<span id="span_cUserPassword"></span></td>
								</tr>
								<%End if%>
								<tr class="back">
									<td height="16">
										<div align="right">������ʾ����</div>
									</td>
									<td>
										<input name="PassQuestion" type="text" id="PassQuestion" size="30" maxlength="30">
									</td>
									<td rowspan="2">������������ʱ���ɴ��һ����롣���磬�����ǡ��ҵĸ����˭��������Ϊ&quot;coolls8&quot;�����ⳤ�Ȳ�����36���ַ���һ������ռ�����ַ����𰸳�����6��30λ֮�䣬���ִ�Сд��</td>
								</tr>
								<tr class="back">
									<td height="16">
										<div align="right">�����</div>
									</td>
									<td>
										<input name="PassAnswer" type="text" id="PassAnswer" size="30" maxlength="50">
									</td>
								</tr>
								<tr class="back">
									<td height="16">
										<div align="right">��ȫ��</div>
									</td>
									<td>
										<input name="SafeCode" type="password" id="SafeCode" size="30" maxlength="30">
									</td>
									<td rowspan="2">ȫ�������һ��������Ҫ;������ȫ�볤��Ϊ6��20λ��������ĸ��Сд������ĸ�����֡������ַ���ɡ�<br>
										<span class="tx">�ر����ѣ���ȫ��һ���趨�������������޸�.</span></td>
								</tr>
								<tr class="back">
									<td height="16">
										<div align="right">ȷ�ϰ�ȫ��</div>
									</td>
									<td>
										<input name="cSafeCode" type="password" id="cSafeCode" size="30" maxlength="30">
									</td>
								</tr>
								<tr class="back">
									<td height="16">
										<div align="right">�����ʼ�</div>
									</td>
									<td>
										<input name="Email" type="text" id="Email" size="30" maxlength="100" onFocus="Do.these('Email',function(){return checkMail('Email','span_Email')})" onKeyUp="Do.these('Email',function(){return checkMail('Email','span_Email')})">
										<span id="span_Email"></span> <br>
										<a href="javascript:CheckEmail('lib/Checkemail.asp')">�Ƿ�ռ��</a> </td>
									<td>����ע������ʼ���<span class="tx">ע��ɹ��󣬽������޸�</span></td>
								</tr>
								<tr class="back">
									<td height="16">
										<div align="right">��Ա����</div>
									</td>
									<td>
										<select name="IsCorporation" id="IsCorporation">
											<option value="0" selected>���˻�Ա</option>
											<option value="1">��ҵ��Ա</option>
										</select>
									</td>
									<td>&nbsp;</td>
								</tr>
								<tr class="back">
									<td height="16">
										<div align="right"></div>
									</td>
									<td>&nbsp;</td>
									<td>&nbsp;</td>
								</tr>
								<tr class="back">
									<td height="39" colspan="3">
										<div align="center">
											<input name="SubSys" type="hidden" id="SubSys" value="<% = Request.QueryString("SubSys")%>">
											<input type="submit" name="Submit" value="�����Ա������Ϣ" style="CURSOR:hand">
											<input class="button" onClick="javascript:location.href='../'" type="button"  style="CURSOR:hand" value="������ҳ" name="Submit1" />
										</div>
									</td>
								</tr>
							</form>
						</table>
					</td>
				</tr>
				<tr class="back">
					<td height="26"  colspan="2" class="xingmu">
						<div align="left">
							<!--#include file="Copyright.asp" -->
						</div>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</body>
</html>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->
<script language="JavaScript" type="text/javascript">
function CheckForm()
{
<%if p_AllowChineseName = 0 then%>
	if( strlen2(UserForm.UserName.value) ) {
	alert("\�����û��������зǷ��ַ�,���������ַ�")
	UserForm.UserName.focus();
	return false;
	}
<%End if%>	
	if(document.UserForm.UserName.value=="")
	{
		alert("�����������û���!");
		document.UserForm.UserName.focus();
		return false;
	}
	if(UserForm.UserName.value.length<<%=p_NumLenMin%> ) {
	alert("�û�������<%=p_NumLenMin%>���ַ�");
	UserForm.UserName.focus();
	return false;
	}
	if(UserForm.UserName.value.length><%=p_NumLenMax%> ) {
		alert("�û������ܳ���<%=p_NumLenMax%>���ַ�");
		UserForm.UserName.focus();
		return false;
		}
<% if p_isValidate = 0 then%>
	if(document.UserForm.UserPassword.value == "")
	{
		alert("��������������");
		document.UserForm.UserPassword.focus();
		return false;
	}
	if(UserForm.UserPassword.value.length><%=p_LenPassworMax%> ) {
		alert("���볤�Ȳ��ܳ���<%=p_LenPassworMax%>���ַ�");
		UserForm.UserPassword.focus();
		return false;
		}
	if(UserForm.UserPassword.value.length<<%=p_LenPassworMin%> ) {
		alert("���볤�Ȳ�������<%=p_LenPassworMin%>���ַ�");
		UserForm.UserPassword.focus();
		return false;
		}
	if(document.UserForm.UserPassword.value !== document.UserForm.cUserPassword.value)
	{
		alert("2�����벻һ��");
		document.UserForm.cUserPassword.focus();
		return false;
	}
<%End if%>
	if(document.UserForm.PassQuestion.value == "")
	{
		alert("����������������ʾ����");
		document.UserForm.PassQuestion.focus();
		return false;
	}
		if(UserForm.PassQuestion.value.length>36 ) {
		alert("������ʾ���ⳤ�Ȳ��ܳ���36���ַ�");
		UserForm.PassQuestion.focus();
		return false;
		}
	if(document.UserForm.PassAnswer.value == "")
	{
		alert("���������������");
		document.UserForm.PassAnswer.focus();
		return false;
	}
	if(document.UserForm.SafeCode.value == "")
	{
		alert("���������İ�ȫ��");
		document.UserForm.SafeCode.focus();
		return false;
	}
	if(document.UserForm.UserPassword.value == document.UserForm.SafeCode.value)
	{
		alert("���벻�ܺͰ�ȫ����ͬ");
		document.UserForm.SafeCode.focus();
		return false;
	}
	if(document.UserForm.SafeCode.value != document.UserForm.cSafeCode.value)
	{
		alert("2�ΰ�ȫ�벻һ��");
		document.UserForm.cSafeCode.focus();
		return false;
	}
	if(UserForm.SafeCode.value.length>20 ) {
		alert("��ȫ�볤�Ȳ��ܳ���20���ַ�");
		UserForm.SafeCode.focus();
		return false;
		}
	if(UserForm.SafeCode.value.length<6 ) {
		alert("��ȫ�볤�Ȳ�������6���ַ�");
		UserForm.SafeCode.focus();
		return false;
		}
	if(UserForm.Email.value.length<8 || UserForm.Email.value.length>64) {
		alert("����������ȷ�������ַ!");
		UserForm.Email.focus();
		return false;
	}
	if(document.UserForm.IsCorporation.value == "")
	{
		alert("��ѡ�����Ļ�Ա����");
		document.UserForm.IsCorporation.focus();
		return false;
	}
<%if p_AllowChineseName = 0 then%>
	function strlen2(str){
		var len;
		var i;
		len = 0;
		for (i=0;i<str.length;i++){
			if (str.charCodeAt(i)>255) return true;
		}
		return false;
	}
	function isSsnString (ssn)
	{
		var re=/^[0-9a-z][\w-.]*[0-9a-z]$/i;
		if(re.test(ssn))
			return true;
		else
			return false;
	}
<%End if%>
}
function CheckName(gotoURL) {
   var ssn=UserForm.UserName.value.toLowerCase();
	   var open_url = gotoURL + "?Username=" + ssn;
	   window.open(open_url,'','status=0,directories=0,resizable=0,toolbar=0,location=0,scrollbars=0,width=150,height=80');
}
function CheckEmail(gotoURL) {
   var ssn1=UserForm.Email.value.toLowerCase();
	   var open_url = gotoURL + "?email=" + ssn1;
	   window.open(open_url,'','status=0,directories=0,resizable=0,toolbar=0,location=0,scrollbars=0,width=150,height=80');
}
</script>
