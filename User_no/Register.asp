<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Md5.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="../API/Cls_PassportApi.asp" -->
<%
Dim Fs_User
User_GetParm
Set Fs_User = New Cls_User

if RegisterTF =false then
	strShowErr = "<li>��ʱ�ر�ע�Ṧ��</li><li>����ϵͳ������ʧ!</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl="&Request.ServerVariables("URL")&"?"&request.QueryString&"")
	Response.end
End if
Dim forward,url
forward = NoSqlHack(Request.QueryString("forward"))
forward = Server.URLEncode(forward)
'���������ݿ�
If Request.Form("Action") = "SaveData" then
	Dim p_UserName_1,p_UserPassword_1,p_PassQuestion_1,p_PassAnswer_1,p_SafeCode_1,p_Email_1,p_IsCorporation_1,p_unMD5Password_1,LimitUserNameArr
	Dim p_NickName,p_RealName,p_sex,p_Year,p_month,p_day,p_Certificate,p_CerTificateCode,p_VerCode,strUserNumber,p_Province,p_city
	Dim p_C_Name,p_C_ShortName,p_C_Province,p_C_City,p_C_Address,p_C_PostCode,p_C_ConactName,p_C_Tel,p_C_Fax,p_C_VocationClassID,p_C_Website,p_C_size,p_C_Capital,p_C_BankName,p_C_BankUserName,dz_UserPassword
	
	
	
	'���������ȷ��
	p_UserName_1 = NoHtmlHackInput(NoSqlHack(Request.Form("UserName")))

	p_PassQuestion_1 = NoHtmlHackInput(Replace(Replace(Request.Form("PassQuestion"),"""",""),"'",""))
	p_PassAnswer_1 = Request.Form("PassAnswer")
	p_SafeCode_1 = Request.Form("SafeCode")
	if p_SafeCode_1<>"" then
		p_SafeCode_1 = MD5(p_SafeCode_1,16)
	end if
	p_Email_1 = NoHtmlHackInput(Replace(Replace(Request.Form("Email"),"""",""),"'",""))
	p_IsCorporation_1 = cint(Request.Form("IsCorporation"))
	If Trim(p_UserName_1)="" or Trim(p_Email_1)="" then
		strShowErr = "<li>����Ĳ����ύ!</li><li>�벻Ҫ���ⲿ�ύ����!</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End if
	if Fs_User.chk_regname(p_LimitUserName,p_UserName_1) = false then
		strShowErr = "<li>�û���������ע��</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End if 
	Dim GetUserTFObj
	Set GetUserTFObj = server.CreateObject(G_FS_RS)
	GetUserTFObj.open "select UserName,Email From FS_ME_Users where UserName = '"& NoSqlHack(p_UserName_1) &"' or Email='"&NoSqlHack(p_Email_1)&"'",User_Conn,1,3
	If  Not GetUserTFObj.eof then
		strShowErr = "<li>����д���û�������ӵ����ʼ��Ѿ���ע�ᣬ��ʹ���û��������ʼ���ַ��������������ע�ᣡ</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
		GetUserTFObj.Close:set GetUserTFObj = nothing  
	End if
	

	if p_isValidate = 1 then
		Dim Randomizecode
		Randomizecode=GetRamCode(8)
		p_UserPassword_1 = MD5(Randomizecode,16)
		p_unMD5Password_1 = NoSqlHack(Randomizecode)
		session("Tmp_p_password")  = p_UserPassword_1
		session("Tmp_password") = p_unMD5Password_1
	Else
		p_unMD5Password_1 = NoSqlHack(Request.Form("UserPassword"))
		p_UserPassword_1 = MD5(p_unMD5Password_1,16)
	End if
	dz_UserPassword = md5(p_unMD5Password_1,32)
	p_NickName = NoSqlHack(Request.Form("NickName"))
	p_RealName = Replace(NoSqlHack(Request.Form("RealName")),"''","")
	p_sex = CintStr(NoSqlHack(Request.Form("sex")))
	p_Year = NoSqlHack(Request.Form("Year"))
	p_month = NoSqlHack(Request.Form("month"))
	p_day = NoSqlHack(Request.Form("day"))
	p_Certificate = CintStr(NoSqlHack(Request.Form("Certificate")))
	p_CerTificateCode = Replace(NoSqlHack(Request.Form("CerTificateCode")),"''","")
	p_VerCode = NoSqlHack(Request.Form("VerCode"))
	p_Province = NoSqlHack(Request.Form("Province"))
	p_City = NoSqlHack(Request.Form("City"))

	p_C_Name = NoSqlHack(Request.Form("C_Name"))
	p_C_ShortName = NoSqlHack(Request.Form("C_ShortName"))
	p_C_Province = NoSqlHack(Request.Form("C_Province"))
	p_C_City = NoSqlHack(Request.Form("C_City"))
	p_C_Address = NoSqlHack(Request.Form("C_Address"))
	p_C_PostCode = NoSqlHack(Request.Form("C_PostCode"))
	p_C_ConactName = NoSqlHack(Request.Form("C_ConactName"))
	p_C_Tel = NoSqlHack(Request.Form("C_Tel"))
	p_C_Fax = NoSqlHack(Request.Form("C_Fax"))
	p_C_VocationClassID = NoSqlHack(Request.Form("C_VocationClassID"))
	If p_C_VocationClassID="" Then
		p_C_VocationClassID=0
	End If
	p_C_Website = NoSqlHack(Request.Form("C_Website"))
	p_C_size = NoSqlHack(Request.Form("C_size"))
	p_C_Capital = NoSqlHack(Request.Form("C_Capital"))
	p_C_BankName = NoSqlHack(Request.Form("C_BankName"))
	p_C_BankUserName = NoSqlHack(Request.Form("C_BankUserName"))
	dim strUserNumberRule
	If Not p_UserNumberRule="" Or Not IsNull(p_UserNumberRule) Then
	strUserNumberRule= Fs_User.strUserNumberRule(p_UserNumberRule)
	Else
		strUserNumberRule=GetRamCode(13)
	End if
	if cstr(Session("GetCode"))<>cstr(lcase(Trim(p_VerCode))) then
			strShowErr = "<li>��֤�벻��ȷ��</li><li>��ʱ��û����,�����֤ˢ��һ��������!</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
	End if
	if Request.Form("IsCorporation") = "1" And p_C_Name<>"" then
		Dim AddCorpDataTFObj
		Set AddCorpDataTFObj = server.CreateObject(G_FS_RS)
		AddCorpDataTFObj.open "select  C_Name From FS_ME_CorpUser where C_Name = '"& p_C_Name &"'",User_Conn,1,3
		If Not AddCorpDataTFObj.eof then
			strShowErr = "<li>���ύ����ҵ�����Ѿ���ע�ᣡ</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		AddCorpDataTFObj.close:set AddCorpDataTFObj = nothing
	End if
	strShowErr = ""
	'-----------------------------------------------------------------
	'ϵͳ����
	'-----------------------------------------------------------------
	Dim API_Obj,API_SaveCookie,SysKey
	If API_Enable Then
		Set API_Obj = New PassportApi
			API_Obj.NodeValue "action","checkname",0,False
			API_Obj.NodeValue "username",p_UserName_1,1,False
			SysKey = Md5(p_UserName_1&API_SysKey,16)
			API_Obj.NodeValue "syskey",SysKey,0,False
			API_Obj.NodeValue "email",p_Email_1,1,False
			API_Obj.SendHttpData
			If API_Obj.Status = "1" Then
				Response.redirect "lib/Error.asp?ErrCodes="&API_Obj.Message&"&ErrorUrl=../Reg_info.asp"
				REsponse.End()
			End If
		Set API_Obj = Nothing
	End If
	'-----------------------------------------------------------------
	Dim AddUserDataTFObj,UserNumberRuleObj,AddUserDataObj
	Set AddUserDataTFObj = server.CreateObject(G_FS_RS)
	AddUserDataTFObj.open "select  UserName,Email From FS_ME_Users where UserName = '"& p_UserName_1 &"'",User_Conn,1,3
	If Not AddUserDataTFObj.eof then
		strShowErr = strShowErr & "<li>���ύ���û������ߵ����ʼ��Ѿ���ע��!</li>"
	End if
	AddUserDataTFObj.close:set AddUserDataTFObj =nothing
	'�ж��û�����Ƿ��ظ�
	Set UserNumberRuleObj = server.CreateObject(G_FS_RS)
	UserNumberRuleObj.open "select UserNumber From FS_ME_Users where UserName='"& p_UserName_1&"'",User_Conn,1,1
	If Not UserNumberRuleObj.eof then
		strShowErr = strShowErr & "<li>���ύ���û���������ظ����ǳ���Ǹ����������д�û����ϡ�!</li>"
	End if
	If strShowErr<>"" Then
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Reg_info.asp")
		Response.End()
	End If
	'-----------------------------------------------------------------
	'ϵͳ����
	'-----------------------------------------------------------------
	If API_Enable Then
		'SysKey = Md5(UserName&APISysKey,16)
		Set API_Obj = New PassportApi
			'API_Obj.NodeValue "syskey",SysKey,0,False
			API_Obj.NodeValue "action","reguser",0,False
			API_Obj.NodeValue "username",p_UserName_1,1,False
			SysKey = Md5(API_Obj.XmlNode("username")&API_SysKey,16)
			API_Obj.NodeValue "syskey",SysKey,0,False
			API_Obj.NodeValue "password",p_unMD5Password_1,0,False
			API_Obj.NodeValue "email",p_Email_1,1,False
			API_Obj.NodeValue "question",p_PassQuestion_1,1,False
			API_Obj.NodeValue "answer",p_PassAnswer_1,1,False
			API_Obj.NodeValue "truename",p_RealName,1,False
			API_Obj.NodeValue "gender",p_Sex,0,False
			API_Obj.NodeValue "birthday",p_year&"-"&p_month&"-"&p_day,0,False
			API_Obj.NodeValue "qq","",1,False
			API_Obj.NodeValue "msn","",1,False
			API_Obj.NodeValue "mobile","",1,False
			API_Obj.NodeValue "homepage","",1,False
			API_Obj.SendHttpData
			If API_Obj.Status = "1" Then
				Response.redirect "lib/Error.asp?ErrCodes="&API_Obj.Message&"&ErrorUrl=../Reg_info.asp"
				REsponse.End()
			Else
				API_SaveCookie = API_Obj.SetCookie(SysKey,p_UserName_1,MD5(p_unMD5Password_1,16),0)
			End If
		Set API_Obj = Nothing
	End If
	'-----------------------------------------------------------------
	'�����û�����

	Set AddUserDataObj = server.CreateObject(G_FS_RS)
	AddUserDataObj.open "select  * From FS_ME_Users where 1=0",User_Conn,1,3
	AddUserDataObj.addNew
	AddUserDataObj("UserNumber") = strUserNumberRule
	AddUserDataObj("UserName") = p_UserName_1
	AddUserDataObj("UserPassword") = p_UserPassword_1
	AddUserDataObj("PassQuestion") = p_PassQuestion_1
	AddUserDataObj("PassAnswer") = MD5(p_PassAnswer_1,16)
	AddUserDataObj("safeCode") = p_safeCode_1
	AddUserDataObj("Email") = p_Email_1
	AddUserDataObj("isMessage") = 0
	AddUserDataObj("HeadPicsize") = "60,60"
	AddUserDataObj("NickName") = p_NickName
	AddUserDataObj("RealName") = p_RealName
	AddUserDataObj("Province") = p_Province
	AddUserDataObj("city") = p_city
	AddUserDataObj("Sex") = p_Sex
	AddUserDataObj("BothYear") = p_year&"-"&p_month&"-"&p_day
	AddUserDataObj("Certificate") = p_Certificate
	AddUserDataObj("CertificateCode") = p_CertificateCode
	AddUserDataObj("IsCorporation") = p_IsCorporation_1
	AddUserDataObj("RegTime") = now
	AddUserDataObj("CloseTime") = "3000-1-1"
	AddUserDataObj("LoginNum") = 0
	AddUserDataObj("Integral") = p_NumGetPoint
	AddUserDataObj("FS_Money") = p_NumGetMoney
	'AddUserDataObj("TempLastLoginTime") = year(now)&"-"&month(now)&"-"&day(now)
	AddUserDataObj("TempLastLoginTime") = now
	AddUserDataObj("TempLastLoginTime_1") = now
	if p_RegisterCheck = 1 then
		AddUserDataObj("isLock") = 1
	Else
		AddUserDataObj("isLock") = 0
	End if
	AddUserDataObj("MySkin") = 2
	AddUserDataObj("OnlyLogin") = 0
	AddUserDataObj("ConNumber") = 0
	AddUserDataObj("ConNumberNews") = 0
	AddUserDataObj("isOpen") = 0
	AddUserDataObj("GroupID") = DefaultGroupID
	AddUserDataObj.Update
	AddUserDataObj.close:set AddUserDataObj = nothing
	'�������ݣ������Ӧ���޻��߽�ң�����
	'˵�������Ϊ0��������
	'��ʼ��������
	Dim rsCreatGroup 
	set rsCreatGroup =User_Conn.execute("select GroupID,GroupPoint,GroupMoney,GroupDate From FS_ME_Group where GroupID="&Clng(DefaultGroupID))
	if not rsCreatGroup.eof then
		if rsCreatGroup("GroupPoint")>0 then
			User_Conn.execute("Update FS_ME_Users Set Integral=Integral+"& rsCreatGroup("GroupPoint")&" where UserNumber='"& NoSqlHack(strUserNumberRule) &"'")
		end if
		if rsCreatGroup("GroupMoney")>0 then
			User_Conn.execute("Update FS_ME_Users Set FS_Money=FS_Money+"& rsCreatGroup("GroupMoney") &" where UserNumber='"& NoSqlHack(strUserNumberRule) &"'")
		end if
		if rsCreatGroup("GroupDate")>0 then
			dim DateCoseed
			DateCoseed = dateAdd("d",rsCreatGroup("GroupDate"),date)
			if G_IS_SQL_User_DB=0 then
				User_Conn.execute("Update FS_ME_Users Set CloseTime=#"& NoSqlHack(DateCoseed) &"# where UserNumber='"& NoSqlHack(strUserNumberRule) &"'")
			else
				User_Conn.execute("Update FS_ME_Users Set CloseTime='"& NoSqlHack(DateCoseed) &"' where UserNumber='"& NoSqlHack(strUserNumberRule) &"'")
			end if
		end if
	end if
	rsCreatGroup.close:set rsCreatGroup = nothing
	if p_IsCorporation_1 = 1 then
		Dim AddCorpDataObj
		Set AddCorpDataObj = server.CreateObject(G_FS_RS)
		AddCorpDataObj.open "select  * From FS_ME_CorpUser where 1=0",User_Conn,1,3
		AddCorpDataObj.addNew
		AddCorpDataObj("UserNumber") = strUserNumberRule
		AddCorpDataObj("C_Name") = p_C_Name
		AddCorpDataObj("C_ShortName") = p_C_ShortName
		AddCorpDataObj("C_Province") = p_C_Province
		AddCorpDataObj("C_City") = p_C_City
		AddCorpDataObj("C_Address") = p_C_Address
		AddCorpDataObj("C_ConactName") = p_C_ConactName
		AddCorpDataObj("C_Tel") = p_C_Tel
		AddCorpDataObj("C_Fax") = p_C_Fax
		AddCorpDataObj("C_VocationClassID") = p_C_VocationClassID
		AddCorpDataObj("C_WebSite") = p_C_WebSite
		AddCorpDataObj("C_size") = p_C_size
		AddCorpDataObj("C_Capital") = p_C_Capital
		AddCorpDataObj("C_BankName") = p_C_BankName
		AddCorpDataObj("C_BankUserName") = p_C_BankUserName
		AddCorpDataObj("C_PostCode") = p_C_PostCode
		AddCorpDataObj("isYellowPage") = 0 
		AddCorpDataObj("isYellowPageCheck") = 0 
		if p_isCheckCorp = 1 then
			AddCorpDataObj("isLockCorp") =1
		Else
			AddCorpDataObj("isLockCorp") =0
		End if
		AddCorpDataObj.update
		AddCorpDataObj.close:set AddCorpDataObj = nothing
	End if
	session("FS_UserName") = p_UserName_1
	session("FS_UserNumber") = strUserNumberRule
	session("FS_NickName") = p_NickName
	session("FS_UserPassword") = p_UserPassword_1
	session("TMP_UserPassword") = p_unMD5Password_1
	session("FS_IsCorp") = p_IsCorporation_1
	session("FS_UserEmail") = p_Email_1
	session("FS_IsLock") = p_RegisterCheck
	Response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") = 1
	Call Fs_User.InsertMyPara(session("FS_UserNumber") )
	Call Fs_User.AddLog("ע��",session("FS_UserNumber"),p_NumGetPoint,p_NumGetMoney,"ע���û���",0) 
	Dim str_isSendMail,FsoObj,Path
	Set FsoObj = Server.CreateObject(G_FS_FSO)  
	Path = strUserNumberRule
	if FsoObj.FolderExists(Server.MapPath("..\"&G_USERFILES_DIR) ) = false then FsoObj.createFolder Server.MapPath("..\"&G_USERFILES_DIR) 
	Path = Server.MapPath("..\"&G_USERFILES_DIR&"\"&Path) 
	if FsoObj.FolderExists(Path) = True then FsoObj.deleteFolder Path
	FsoObj.CreateFolder Path
	str_isSendMail=False
	url = "Reg_Result.asp?SubSys="&NoSqlHack(Request.Form("SubSys"))&""
	Response.Write API_SaveCookie
	Response.Flush
	Response.Write "<script language=""JavaScript"">window.location.href="""&url&""";</script>"
	Response.end
End if
set Fs_User = nothing

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��Աע��</title>
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
<form name="UserForm" id="UserForm" method="post" action="Register.asp?forward=<%= forward %>"  onsubmit="return CheckForm();">
	<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
		<tbody>
			<tr class="back">
				<td height="20" colspan="3" class="xingmu">����д�����û���<span class="tx">(������Ŀ������д)</span></td>
			</tr>
			<tr class="back">
				<td width="11%" align="right">�û���</td>
				<td width="37%">
					<input name="UserName" type="text" id="UserName" size="20"  onfocus="Do.these('UserName',function(){return isEmpty('UserName','span_UserName')})" onkeyup="Do.these('UserName',function(){return isEmpty('UserName','span_UserName')})" />
					<a href="javascript:CheckName('lib/CheckName.asp')">����û���</a> <span id="span_UserName"></span> </td>
				<td width="52%">������a��z��Ӣ����ĸ(�����ִ�Сд)��0��9�����֡��㡢���Ż��»��߼�������ɣ�����Ϊ3��18���ַ���ֻ�������ֻ���ĸ��ͷ�ͽ�β,����:coolls1980��</td>
			</tr>
			<%If p_isValidate = 0 then%>
			<tr class="back">
				<td height="16" align="right">����</td>
				<td>
					<input name="UserPassword" type="password" id="UserPassword" size="30" maxlength="50" onfocus="Do.these('UserPassword',function(){return isEmpty('UserPassword','span_UserPassword')})" onkeyup="Do.these('UserPassword',function(){return isEmpty('UserPassword','span_UserPassword')})" />
					<span id="span_UserPassword"></span></td>
				<td rowspan="2">���볤��Ϊ<%=p_LenPassworMin%>��<%=p_LenPassworMax%>λ��������ĸ��Сд����¼�����������ĸ�����֡������ַ���ɡ�</td>
			</tr>
			<tr class="back">
				<td height="24" align="right">ȷ������</td>
				<td>
					<input name="cUserPassword" type="password" id="cUserPassword" size="30" maxlength="50" onfocus="Do.these('cUserPassword',function(){return isEmpty('cUserPassword','span_cUserPassword')})" onkeyup="Do.these('cUserPassword',function(){return isEmpty('cUserPassword','span_cUserPassword')})" />
					<span id="span_cUserPassword"></span></td>
			</tr>
			<%End if%>
			<tr class="back">
				<td height="16" align="right">�����ʼ�</td>
				<td>
					<input name="Email" type="text" id="Email" size="30" maxlength="100" onfocus="Do.these('Email',function(){return checkMail('Email','span_Email')})" onkeyup="Do.these('Email',function(){return checkMail('Email','span_Email')})" />
					<span id="span_Email"></span> <br />
					<a href="javascript:CheckEmail('lib/Checkemail.asp')">�Ƿ�ռ��</a> </td>
				<td>����ע������ʼ���<span class="tx">ע��ɹ��󣬽������޸�</span></td>
			</tr>
			<tr class="back">
				<td align="right">��Ա����</td>
				<td>
					<select name="IsCorporation" id="IsCorporation">
						<option value="0" selected="selected">���˻�Ա</option>
						<option value="1">��ҵ��Ա</option>
					</select>
				</td>
				<td></td>
			</tr>
			<tr class="back">
				<td align="right">У����</td>
				<td>
					<input name="VerCode" type="text" id="VerCode" size="15" onfocus="Do.these('VerCode',function(){return isEmpty('VerCode','span_VerCode')})" onkeyup="Do.these('VerCode',function(){return isEmpty('VerCode','span_VerCode')})" />
					<img src="../Fs_Inc/vCode.asp?" onclick="this.src+=Math.random()" alt="ͼƬ�����壿������µõ���֤��,ע�⣺�����Сд" style="cursor:hand;" /> <span id="span_VerCode"></span> </td>
				<td></td>
			</tr>
			<tr class="back">
				<td align="right">�߼�ѡ��</td>
				<td>
					<input id="btnAdvOption" type="checkbox" />
					<label id="msgAdvOption" for="btnAdvOption">��ʾ�߼�ѡ��</label>
				</td>
				<td></td>
			</tr>
		</tbody>
		<tbody id="pnlAdvanceOption" style="display:none;">
			<tr class="back">
				<td height="16" colspan="3" class="xingmu">�߼�ѡ��</td>
			</tr>
			<tr class="back">
				<td height="16" align="right">������ʾ����</td>
				<td>
					<input name="PassQuestion" type="text" id="PassQuestion" size="30" 
						maxlength="36" />
				</td>
				<td rowspan="2">������������ʱ���ɴ��һ����롣���磬�����ǡ��ҵĸ����˭��������Ϊ&quot;coolls8&quot;�����ⳤ�Ȳ�����36���ַ���һ������ռ�����ַ����𰸳�����6��30λ֮�䣬���ִ�Сд��</td>
			</tr>
			<tr class="back">
				<td height="16" align="right">�����</td>
				<td>
					<input name="PassAnswer" type="text" id="PassAnswer" size="30" maxlength="30" />
				</td>
			</tr>
			<tr class="back">
				<td height="16" align="right">��ȫ��</td>
				<td>
					<input name="SafeCode" type="password" id="SafeCode" size="30" maxlength="20" />
				</td>
				<td rowspan="3">ȫ�������һ��������Ҫ;������ȫ�볤��Ϊ6��20λ��������ĸ��Сд������ĸ�����֡������ַ���ɡ�<br />
					<span class="tx">�ر����ѣ���ȫ��һ���趨�������������޸�.</span></td>
			</tr>
			<tr class="back">
				<td align="right">ȷ�ϰ�ȫ��</td>
				<td>
					<input name="cSafeCode" type="password" id="cSafeCode" size="30" 
						maxlength="20" />
				</td>
			</tr>
			<tr class="back">
				<td height="20" colspan="3" class="xingmu">����д�������ϣ���������Ϣ����ͨ���ͷ�ȡ���ʺŵ�������ݣ�����ʵ��д�� </td>
			</tr>
			<tr class="back">
				<td height="27" align="right">
					�ǳ�</td>
				<td>
					<input name="NickName" type="text" id="NickName" size="20" maxlength="20" />
				</td>
				<td>����д��������ǳơ�����Ϊ����</td>
			</tr>
			<tr class="back">
				<td width="15%" height="27" align="right">
					����</td>
				<td width="29%">
					<input name="RealName" type="text" id="RealName" size="20" maxlength="20" />
				</td>
				<td width="56%">����д������ʵ������</td>
			</tr>
			<tr class="back">
				<td height="16" align="right">
					�Ա�</td>
				<td>
					<input type="radio" name="sex" value="0" checked="checked" />
					��
					<input type="radio" name="sex" value="1" />
					Ů
				</td>
				<td>����ѡ���Ա�</td>
			</tr>
			<tr class="back">
				<td height="24" align="right">
					��������</td>
				<td>
					<input name="Year" type="text" id="Year" value="19" size="5" maxlength="4" />
					��
					<select name="month" id="month">
						<option value="1" selected="selected">1</option>
						<option value="2">2</option>
						<option value="3">3</option>
						<option value="4">4</option>
						<option value="5">5</option>
						<option value="6">6</option>
						<option value="7">7</option>
						<option value="8">8</option>
						<option value="9">9</option>
						<option value="10">10</option>
						<option value="11">11</option>
						<option value="12">12</option>
					</select>
					��
					<select name="day" id="day">
						<option value="1" selected="selected">1</option>
						<option value="2">2</option>
						<option value="3">3</option>
						<option value="4">4</option>
						<option value="5">5</option>
						<option value="6">6</option>
						<option value="7">7</option>
						<option value="8">8</option>
						<option value="9">9</option>
						<option value="10">10</option>
						<option value="11">11</option>
						<option value="12">12</option>
						<option value="13">13</option>
						<option value="14">14</option>
						<option value="15">15</option>
						<option value="16">16</option>
						<option value="17">17</option>
						<option value="18">18</option>
						<option value="19">19</option>
						<option value="20">20</option>
						<option value="21">21</option>
						<option value="22">22</option>
						<option value="23">23</option>
						<option value="24">24</option>
						<option value="25">25</option>
						<option value="26">26</option>
						<option value="27">27</option>
						<option value="28">28</option>
						<option value="29">29</option>
						<option value="30">30</option>
						<option value="31">31</option>
					</select>
					�� </td>
				<td>����д������ʵ���գ���������ȡ�����롣</td>
			</tr>
			<tr class="back">
				<td height="16" align="right">
					֤�����</td>
				<td>
					<select name="Certificate" id="Certificate">
						<option value="0" selected="selected">���֤</option>
						<option value="2">ѧ��֤</option>
						<option value="1">��ʻ֤</option>
						<option value="3">����֤</option>
						<option value="4">����</option>
					</select>
				</td>
				<td rowspan="2">��Ч֤����Ϊȡ���ʺŵ�����ֶΣ����Ժ�ʵ�ʺŵĺϷ���ݣ����������ʵ��д��<br />
					<span class="tx">�ر����ѣ���Ч֤��һ���趨�����ɸ���</span></td>
			</tr>
			<tr class="back">
				<td height="16" align="right">
					֤������</td>
				<td>
					<input name="CerTificateCode" type="text" id="CerTificateCode" size="30" 
						maxlength="18" />
				</td>
			</tr>
			<tr class="back">
				<td height="16" align="right">
					ʡ��</td>
				<td height="16">
					<select name="Province" id="Province">
						<option value="">��ѡ��</option>
						<option value="�Ĵ�">�Ĵ�</option>
						<option value="����">����</option>
						<option value="�Ϻ�">�Ϻ�</option>
						<option value="���">���</option>
						<option value="����">����</option>
						<option value="����">����</option>
						<option value="����">����</option>
						<option value="�㶫">�㶫</option>
						<option value="����">����</option>
						<option value="����">����</option>
						<option value="����">����</option>
						<option value="����">����</option>
						<option value="�ӱ�">�ӱ�</option>
						<option value="����">����</option>
						<option value="������">������</option>
						<option value="����">����</option>
						<option value="����">����</option>
						<option value="����">����</option>
						<option value="����">����</option>
						<option value="����">����</option>
						<option value="����">����</option>
						<option value="���ɹ�">���ɹ�</option>
						<option value="����">����</option>
						<option value="�ຣ">�ຣ</option>
						<option value="ɽ��">ɽ��</option>
						<option value="ɽ��">ɽ��</option>
						<option value="����">����</option>
						<option value="����">����</option>
						<option value="�½�">�½�</option>
						<option value="����">����</option>
						<option value="�㽭">�㽭</option>
						<option value="�۰�̨">�۰�̨</option>
						<option value="����">����</option>
						<option value="����">����</option>
					</select>
				</td>
				<td height="16" >���������ڵ�ʡ��</td>
			</tr>
			<tr class="back">
				<td height="16" align="right">
					����</td>
				<td height="16">
					<input name="City" type="text" id="City" size="30" maxlength="20" />
				</td>
				<td height="16">���������ڵĳ���</td>
			</tr>
		</tbody>
		<tbody id="pnlCompany" style="display:none;">
			<tr class="back">
				<td height="16" colspan="3" class="xingmu">��д��˾���ϣ� </td>
			</tr>
			<tr class="back">
				<td height="16" align="right">
					��˾����</td>
				<td>
					<input name="C_Name" type="text" id="C_Name" size="30" maxlength="50" />
				</td>
				<td>����д����˾��ȫ��</td>
			</tr>
			<tr class="back">
				<td height="-2" align="right">
					��˾���</td>
				<td>
					<input name="C_ShortName" type="text" id="C_ShortName" size="30" maxlength="30" />
				</td>
				<td>����д����˾�ļ򵥳ƺ�</td>
			</tr>
			<tr class="back">
				<td height="0" align="right">
					��˾����ʡ��</td>
				<td>
					<select name="C_Province" id="C_Province">
						<option value="">��ѡ�񡡡�</option>
						<option value="�Ĵ�">�Ĵ�</option>
						<option value="����">����</option>
						<option value="�Ϻ�">�Ϻ�</option>
						<option value="���">���</option>
						<option value="����">����</option>
						<option value="����">����</option>
						<option value="����">����</option>
						<option value="�㶫">�㶫</option>
						<option value="����">����</option>
						<option value="����">����</option>
						<option value="����">����</option>
						<option value="����">����</option>
						<option value="�ӱ�">�ӱ�</option>
						<option value="����">����</option>
						<option value="������">������</option>
						<option value="����">����</option>
						<option value="����">����</option>
						<option value="����">����</option>
						<option value="����">����</option>
						<option value="����">����</option>
						<option value="����">����</option>
						<option value="���ɹ�">���ɹ�</option>
						<option value="����">����</option>
						<option value="�ຣ">�ຣ</option>
						<option value="ɽ��">ɽ��</option>
						<option value="ɽ��">ɽ��</option>
						<option value="����">����</option>
						<option value="����">����</option>
						<option value="�½�">�½�</option>
						<option value="����">����</option>
						<option value="�㽭">�㽭</option>
						<option value="�۰�̨">�۰�̨</option>
						<option value="����">����</option>
						<option value="����">����</option>
					</select>
				</td>
				<td>����д����˾���ڵ�ʡ��</td>
			</tr>
			<tr class="back">
				<td height="0" align="right">
					��˾���ڳ���</td>
				<td>
					<input name="C_City" type="text" id="C_City" size="30" maxlength="20" />
				</td>
				<td>����д����˾���ڵĳ���</td>
			</tr>
			<tr class="back">
				<td height="0" align="right">
					��˾��ַ</td>
				<td>
					<input name="C_Address" type="text" id="C_Address" size="30" maxlength="100" />
				</td>
				<td>���Ĺ�˾��ַ</td>
			</tr>
			<tr class="back">
				<td height="0" align="right">
					��������</td>
				<td>
					<input name="C_PostCode" type="text" id="C_PostCode" size="30" maxlength="20" />
				</td>
				<td>����˾����������</td>
			</tr>
			<tr class="back">
				<td height="0" align="right">
					��˾��ϵ��</td>
				<td>
					<input name="C_ConactName" type="text" id="C_ConactName" size="30" maxlength="20" />
				</td>
				<td>��˾��ϵ��</td>
			</tr>
			<tr class="back">
				<td height="0" align="right">
					��˾��ϵ�绰</td>
				<td>
					<input name="C_Tel" type="text" id="C_Tel" size="30" maxlength="20" />
				</td>
				<td>��˾��ϵ�绰���зֻ�����&quot;-&quot;�ֿ����磺028-85098980-606</td>
			</tr>
			<tr class="back">
				<td height="1" align="right">
					��˾����</td>
				<td>
					<input name="C_Fax" type="text" id="C_Fax" size="30" maxlength="20" />
				</td>
				<td>��˾���档�зֻ�����&quot;-&quot;�ֿ����磺028-85098980-606</td>
			</tr>
			<tr class="back">
				<td height="3" align="right">
					��ҵ</td>
				<td>
					<input name="C_VocationClassName" type="text" id="C_VocationClassName" size="30" readonly="readonly" />
					<input type="hidden" name="C_VocationClassID" id="C_VocationClassID" />
				</td>
				<td>
					<input type="button" name="Submit3" value="ѡ����ҵ" onclick="SelectClass();" />
					��˾���ڵ���ҵ</td>
			</tr>
			<tr class="back">
				<td height="8" align="right">
					��˾��վ</td>
				<td>
					<input name="C_Website" type="text" id="C_Website" size="30" maxlength="200" />
				</td>
				<td>��˾���ڵ���ҵվ��</td>
			</tr>
			<tr class="back">
				<td height="16" align="right">
					��˾��ģ</td>
				<td>
					<select name="C_size" id="C_size">
						<option value="1-20��" selected="selected">1-20��</option>
						<option value="21-50��">21-50��</option>
						<option value="51-100��">51-100��</option>
						<option value="101-200��">101-200��</option>
						<option value="201-500��">201-500��</option>
						<option value="501-1000��">501-1000��</option>
						<option value="1000������">1000������</option>
					</select>
				</td>
				<td>&nbsp;</td>
			</tr>
			<tr class="back">
				<td height="1" align="right">
					��˾ע���ʱ�</td>
				<td>
					<select name="C_Capital" id="C_Capital">
						<option value="10������">10������</option>
						<option value="10��-19��">10��-19��</option>
						<option value="20��-49��">20��-49��</option>
						<option value="50��-99��" selected="selected">50��-99��</option>
						<option value="100��-199��">100��-199��</option>
						<option value="200��-499��">200��-499��</option>
						<option value="500��-999��">500��-999��</option>
						<option value="1000������">1000������</option>
					</select>
				</td>
				<td>&nbsp;</td>
			</tr>
			<tr class="back">
				<td height="3" align="right">
					��������</td>
				<td>
					<input name="C_BankName" type="text" id="C_BankName" size="30" maxlength="50" />
				</td>
				<td rowspan="2">
					<p>��˾�����ʻ����Է������������ϵ�����С�<br />
						�����������ӣ��й��������гɶ�����˫骷���<br />
						�����ʻ�����</p>
				</td>
			</tr>
			<tr class="back">
				<td height="8" align="right">
					�����ʺż��ʻ���</td>
				<td>
					<textarea name="C_BankUserName" cols="30" rows="4" id="C_BankUserName"></textarea>
				</td>
			</tr>
		</tbody>
		<tbody>
			<tr class="back">
				<td height="39" colspan="3" align="center">
					<input name="Action" type="hidden" id="Action" value="SaveData" />
					<input name="SubSys" type="hidden" id="SubSys" value="<% = Request.QueryString("SubSys")%>" />
					<input id="chbAgreement" type="checkbox" checked="checked" /><a href="Agreement.asp" target="_blank">ͬ�Ⲣ����ע��Э��</a>
					<input type="submit" name="Submit" value="ע��" style="CURSOR:pointer" />
					<input onclick="javascript:location.href='../'" type="button"  style="CURSOR:hand" value="������ҳ" name="Submit1" />
				</td>
			</tr>
			<tr class="back">
				<td height="26" colspan="3" class="xingmu">
					<!--#include file="Copyright.asp" -->
				</td>
			</tr>
		</tbody>
	</table>
</form>
<%
Response.Write("<script type=""text/javascript"">var p_AllowChineseName = "&p_AllowChineseName&";var p_isValidate = "&p_isValidate&";var p_NumLenMin = "&p_NumLenMin&";var p_NumLenMax = "&p_NumLenMax&";var p_LenPassworMax = "&p_LenPassworMax&";var p_LenPassworMin = "&p_LenPassworMin&";</script>")
%>
<script type="text/javascript">
	$('IsCorporation').onchange = function() {
	if ($('btnAdvOption').checked && this.value == '1') {
			$('pnlCompany').style.display = '';
		} else {
			$('pnlCompany').style.display = 'none';
		}
	};
	$('btnAdvOption').onclick = function() {
		if (this.checked) {
			$('pnlAdvanceOption').style.display = '';
			$('msgAdvOption').innerHTML = '�رո߼�ѡ��';
		} else {
			$('pnlAdvanceOption').style.display = 'none';
			$('msgAdvOption').innerHTML = '��ʾ�߼�ѡ��';
		}
		if (this.checked && $('IsCorporation').value == '1') {
			$('pnlCompany').style.display = '';
		} else {
			$('pnlCompany').style.display = 'none';
		}
	};
	function CheckForm() {
		if (p_AllowChineseName == 0 && strlen2(UserForm.UserName.value)) {
			alert("\�����û��������зǷ��ַ�,���������ַ�")
			UserForm.UserName.focus();
			return false;
		}
		if (document.UserForm.UserName.value == "") {
			alert("�����������û���!");
			document.UserForm.UserName.focus();
			return false;
		}
		if (UserForm.UserName.value.length < p_NumLenMin) {
			alert("�û�������" + p_NumLenMin + "���ַ�");
			UserForm.UserName.focus();
			return false;
		}
		if (UserForm.UserName.value.length > p_NumLenMax) {
			alert("�û������ܳ���" + p_NumLenMax + "���ַ�");
			UserForm.UserName.focus();
			return false;
		}
		if (p_isValidate == 0) {
			if (document.UserForm.UserPassword.value == "") {
				alert("��������������");
				document.UserForm.UserPassword.focus();
				return false;
			}
			if (UserForm.UserPassword.value.length > p_LenPassworMax) {
				alert("���볤�Ȳ��ܳ���" + p_LenPassworMax + "���ַ�");
				UserForm.UserPassword.focus();
				return false;
			}
			if (UserForm.UserPassword.value.length < p_LenPassworMin) {
				alert("���볤�Ȳ�������" + p_LenPassworMin + "���ַ�");
				UserForm.UserPassword.focus();
				return false;
			}
			if (document.UserForm.UserPassword.value !== document.UserForm.cUserPassword.value) {
				alert("2�����벻һ��");
				document.UserForm.cUserPassword.focus();
				return false;
			}
		}
		if (UserForm.Email.value.length < 8 || UserForm.Email.value.length > 64) {
			alert("����������ȷ�������ַ!");
			UserForm.Email.focus();
			return false;
		}
		if (document.UserForm.IsCorporation.value == "") {
			alert("��ѡ�����Ļ�Ա����");
			document.UserForm.IsCorporation.focus();
			return false;
		}
		if (document.UserForm.VerCode.value == "") {
			alert("��������֤��!");
			document.UserForm.VerCode.focus();
			return false;
		}	

		if (document.UserForm.UserPassword.value == document.UserForm.SafeCode.value) {
			alert("���벻�ܺͰ�ȫ����ͬ");
			document.UserForm.SafeCode.focus();
			return false;
		}
		if (document.UserForm.SafeCode.value != document.UserForm.cSafeCode.value) {
			alert("2�ΰ�ȫ�벻һ��");
			document.UserForm.cSafeCode.focus();
			return false;
		}
		if (UserForm.SafeCode.value.length > 20) {
			alert("��ȫ�볤�Ȳ��ܳ���20���ַ�");
			UserForm.SafeCode.focus();
			return false;
		}

		if (!UserForm.chbAgreement.checked) {
			alert("���Ķ�������ע��Э�飡");
			UserForm.chbAgreement.focus();
			return false;
		}
		
		function strlen2(str) {
			var len;
			var i;
			len = 0;
			for (i = 0; i < str.length; i++) {
				if (str.charCodeAt(i) > 255) return true;
			}
			return false;
		}
		function isSsnString(ssn) {
			var re = /^[0-9a-z][\w-.]*[0-9a-z]$/i;
			if (re.test(ssn))
				return true;
			else
				return false;
		}
	}
	function CheckName(gotoURL) {
		var ssn = $('UserName').value.toLowerCase();
		var open_url = gotoURL + "?Username=" + ssn;
		window.open(open_url, '', 'status=0,directories=0,resizable=0,toolbar=0,location=0,scrollbars=0,width=150,height=80');
	}
	function CheckEmail(gotoURL) {
		var ssn1 = $('Email').value.toLowerCase();
		var open_url = gotoURL + "?email=" + ssn1;
		window.open(open_url, '', 'status=0,directories=0,resizable=0,toolbar=0,location=0,scrollbars=0,width=150,height=80');
	}
	function SelectClass() {
		var ReturnValue = '', TempArray = new Array();
		ReturnValue = OpenWindow('lib/SelectClassFrame.asp', 400, 300, window);
		try {
			document.getElementById('C_VocationClassID').value = ReturnValue[0][0];
			document.getElementById('C_VocationClassName').value = ReturnValue[1][0];
		}
		catch (ex) { }
	}
	function OpenWindow(Url, Width, Height, WindowObj) {
		var ReturnStr = showModalDialog(Url, WindowObj, 'dialogWidth:' + Width + 'pt;dialogHeight:' + Height + 'pt;status:no;help:no;scroll:no;');
		return ReturnStr;
	}
</script>
</body>
</html>
