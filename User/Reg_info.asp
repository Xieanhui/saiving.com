<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<%
User_GetParm
if RegisterTF =false then
	strShowErr = "<li>暂时关闭注册功能</li><li>或者系统参数丢失!</li>"
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
<title>会员注册step 2 of  4 step</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="keywords" content="风讯,风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
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
					<td   colspan="2" class="xingmu" height="24">·User Register step 2.(填写会员基本资料)</td>
				</tr>
				<tr class="back">
					<td width="15%" valign="top" class="hback"><strong>【注册步骤】</strong> <br>
						<br>
						<div align="left"> √同意注册协议<br>
							<br>
							→填写会员资料<br>
							<br>
							×填写联系资料<br>
							<br>
							×注册成功</div>
					</td>
					<td width="86%" valign="top" class="hback">
						<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
							<form name="UserForm"  id="UserForm" method="post" action="Reg_contact.asp?forward=<%= forward %>"  onsubmit="return CheckForm();">
								<tr class="back">
									<td height="20" colspan="3" class="xingmu">请填写您的用户名<span class="tx">(以下项目必须填写)</span></td>
								</tr>
								<tr class="back">
									<td width="11%" height="65">
										<div align="right">用户名</div>
									</td>
									<td width="37%">
										<input name="UserName" type="text" id="UserName" size="20"  onfocus="Do.these('UserName',function(){return isEmpty('UserName','span_UserName')})" onKeyUp="Do.these('UserName',function(){return isEmpty('UserName','span_UserName')})">
										<span id="span_UserName"></span> <a href="javascript:CheckName('lib/CheckName.asp')">检查用户名</a></td>
									<td width="52%">户名由a～z的英文字母(不区分大小写)、0～9的数字、点、减号或下划线及中文组成，长度为3～18个字符，只能以数字或字母开头和结尾,例如:coolls1980。</td>
								</tr>
								<tr class="back">
									<td height="16" colspan="3" class="xingmu">请填写安全设置：（安全设置用于验证帐号和找回密码）</td>
								</tr>
								<%If p_isValidate = 0  then%>
								<tr class="back">
									<td height="16">
										<div align="right">密码</div>
									</td>
									<td>
										<input name="UserPassword" type="password" id="UserPassword" size="30" maxlength="50"  onfocus="Do.these('UserPassword',function(){return isEmpty('UserPassword','span_UserPassword')})" onKeyUp="Do.these('UserPassword',function(){return isEmpty('UserPassword','span_UserPassword')})">
										<span id="span_UserPassword"></span></td>
									<td rowspan="2">密码长度为<%=p_LenPassworMin%>～<%=p_LenPassworMax%>位，区分字母大小写。登录密码可以由字母、数字、特殊字符组成。</td>
								</tr>
								<tr class="back">
									<td height="24">
										<div align="right">确认密码</div>
									</td>
									<td>
										<input name="cUserPassword" type="password" id="cUserPassword" size="30" maxlength="50" onFocus="Do.these('cUserPassword',function(){return isEmpty('cUserPassword','span_cUserPassword')})" onKeyUp="Do.these('cUserPassword',function(){return isEmpty('cUserPassword','span_cUserPassword')})">
										<span id="span_cUserPassword"></span></td>
								</tr>
								<%End if%>
								<tr class="back">
									<td height="16">
										<div align="right">密码提示问题</div>
									</td>
									<td>
										<input name="PassQuestion" type="text" id="PassQuestion" size="30" maxlength="30">
									</td>
									<td rowspan="2">当您忘记密码时可由此找回密码。例如，问题是“我的哥哥是谁？”，答案为&quot;coolls8&quot;。问题长度不大于36个字符，一个汉字占两个字符。答案长度在6～30位之间，区分大小写。</td>
								</tr>
								<tr class="back">
									<td height="16">
										<div align="right">密码答案</div>
									</td>
									<td>
										<input name="PassAnswer" type="text" id="PassAnswer" size="30" maxlength="50">
									</td>
								</tr>
								<tr class="back">
									<td height="16">
										<div align="right">安全码</div>
									</td>
									<td>
										<input name="SafeCode" type="password" id="SafeCode" size="30" maxlength="30">
									</td>
									<td rowspan="2">全码是您找回密码的重要途径，安全码长度为6～20位，区分字母大小写，由字母、数字、特殊字符组成。<br>
										<span class="tx">特别提醒：安全码一旦设定，将不可自行修改.</span></td>
								</tr>
								<tr class="back">
									<td height="16">
										<div align="right">确认安全码</div>
									</td>
									<td>
										<input name="cSafeCode" type="password" id="cSafeCode" size="30" maxlength="30">
									</td>
								</tr>
								<tr class="back">
									<td height="16">
										<div align="right">电子邮件</div>
									</td>
									<td>
										<input name="Email" type="text" id="Email" size="30" maxlength="100" onFocus="Do.these('Email',function(){return checkMail('Email','span_Email')})" onKeyUp="Do.these('Email',function(){return checkMail('Email','span_Email')})">
										<span id="span_Email"></span> <br>
										<a href="javascript:CheckEmail('lib/Checkemail.asp')">是否被占用</a> </td>
									<td>您的注册电子邮件。<span class="tx">注册成功后，将不能修改</span></td>
								</tr>
								<tr class="back">
									<td height="16">
										<div align="right">会员类型</div>
									</td>
									<td>
										<select name="IsCorporation" id="IsCorporation">
											<option value="0" selected>个人会员</option>
											<option value="1">企业会员</option>
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
											<input type="submit" name="Submit" value="保存会员基本信息" style="CURSOR:hand">
											<input class="button" onClick="javascript:location.href='../'" type="button"  style="CURSOR:hand" value="返回首页" name="Submit1" />
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
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->
<script language="JavaScript" type="text/javascript">
function CheckForm()
{
<%if p_AllowChineseName = 0 then%>
	if( strlen2(UserForm.UserName.value) ) {
	alert("\您的用户名不能有非法字符,或者中文字符")
	UserForm.UserName.focus();
	return false;
	}
<%End if%>	
	if(document.UserForm.UserName.value=="")
	{
		alert("请输入您的用户名!");
		document.UserForm.UserName.focus();
		return false;
	}
	if(UserForm.UserName.value.length<<%=p_NumLenMin%> ) {
	alert("用户名至少<%=p_NumLenMin%>个字符");
	UserForm.UserName.focus();
	return false;
	}
	if(UserForm.UserName.value.length><%=p_NumLenMax%> ) {
		alert("用户名不能超过<%=p_NumLenMax%>个字符");
		UserForm.UserName.focus();
		return false;
		}
<% if p_isValidate = 0 then%>
	if(document.UserForm.UserPassword.value == "")
	{
		alert("请输入您的密码");
		document.UserForm.UserPassword.focus();
		return false;
	}
	if(UserForm.UserPassword.value.length><%=p_LenPassworMax%> ) {
		alert("密码长度不能超过<%=p_LenPassworMax%>个字符");
		UserForm.UserPassword.focus();
		return false;
		}
	if(UserForm.UserPassword.value.length<<%=p_LenPassworMin%> ) {
		alert("密码长度不能少于<%=p_LenPassworMin%>个字符");
		UserForm.UserPassword.focus();
		return false;
		}
	if(document.UserForm.UserPassword.value !== document.UserForm.cUserPassword.value)
	{
		alert("2次密码不一致");
		document.UserForm.cUserPassword.focus();
		return false;
	}
<%End if%>
	if(document.UserForm.PassQuestion.value == "")
	{
		alert("请输入您的密码提示问题");
		document.UserForm.PassQuestion.focus();
		return false;
	}
		if(UserForm.PassQuestion.value.length>36 ) {
		alert("密码提示问题长度不能超过36个字符");
		UserForm.PassQuestion.focus();
		return false;
		}
	if(document.UserForm.PassAnswer.value == "")
	{
		alert("请输入您的密码答案");
		document.UserForm.PassAnswer.focus();
		return false;
	}
	if(document.UserForm.SafeCode.value == "")
	{
		alert("请输入您的安全码");
		document.UserForm.SafeCode.focus();
		return false;
	}
	if(document.UserForm.UserPassword.value == document.UserForm.SafeCode.value)
	{
		alert("密码不能和安全码相同");
		document.UserForm.SafeCode.focus();
		return false;
	}
	if(document.UserForm.SafeCode.value != document.UserForm.cSafeCode.value)
	{
		alert("2次安全码不一致");
		document.UserForm.cSafeCode.focus();
		return false;
	}
	if(UserForm.SafeCode.value.length>20 ) {
		alert("安全码长度不能超过20个字符");
		UserForm.SafeCode.focus();
		return false;
		}
	if(UserForm.SafeCode.value.length<6 ) {
		alert("安全码长度不能少于6个字符");
		UserForm.SafeCode.focus();
		return false;
		}
	if(UserForm.Email.value.length<8 || UserForm.Email.value.length>64) {
		alert("请您输入正确的邮箱地址!");
		UserForm.Email.focus();
		return false;
	}
	if(document.UserForm.IsCorporation.value == "")
	{
		alert("请选择您的会员类型");
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
