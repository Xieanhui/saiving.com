<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Md5.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="../API/Cls_PassportApi.asp" -->
<%
Dim Fs_User,url
User_GetParm
Set Fs_User = New Cls_User
if RegisterTF =false then
	strShowErr = "<li>��ʱ�ر�ע�Ṧ��</li><li>����ϵͳ������ʧ!</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
End If
Dim forward
forward = Request.QueryString("forward")
forward = Server.URLEncode(forward)
'���������ݿ�
If Request.Form("Action") = "SaveData" then
	Dim p_UserName_1,p_UserPassword_1,p_PassQuestion_1,p_PassAnswer_1,p_SafeCode_1,p_Email_1,p_IsCorporation_1,p_unMD5Password_1
	Dim p_NickName,p_RealName,p_sex,p_Year,p_month,p_day,p_Certificate,p_CerTificateCode,p_VerCode,strUserNumber,p_Province,p_city
	Dim p_C_Name,p_C_ShortName,p_C_Province,p_C_City,p_C_Address,p_C_PostCode,p_C_ConactName,p_C_Tel,p_C_Fax,p_C_VocationClassID,p_C_Website,p_C_size,p_C_Capital,p_C_BankName,p_C_BankUserName,dz_UserPassword
	p_UserName_1 = Replace(NoSqlHack(Request.Form("UserName")),"''","")
	p_PassQuestion_1 = Replace(NoSqlHack(Request.Form("PassQuestion")),"''","")
	p_PassAnswer_1 = NoSqlHack(Request.Form("PassAnswer"))
	p_SafeCode_1 = MD5(NoSqlHack(Request.Form("SafeCode")),16)
	p_Email_1 = Replace(NoSqlHack(Request.Form("Email")),"''","")
	p_IsCorporation_1 = CintStr(Request.Form("IsCorporation"))
	if p_isValidate = 1 then
		p_UserPassword_1 = session("Tmp_p_password")
		p_unMD5Password_1 = session("Tmp_password")
	Else
		p_UserPassword_1 =NoSqlHack(Request.Form("UserPassword"))
		p_unMD5Password_1 = NoSqlHack(Request.Form("unMD5Password"))
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
	if Request.Form("IsCorporation") = "1" then
		If Trim(p_C_Name)="" Or Trim(p_C_Province)="" Or Trim(p_C_City)="" Or Trim(p_C_PostCode)=""  Or Trim(p_C_ConactName)="" Or Trim(p_C_Tel)="" or trim(p_C_VocationClassID)="" then
				strShowErr = "<li>��ҵע����Ҫ��д�������ϣ�</li><li>����д����,��ҵ����,����ʡ��,����,�ʱ�,��ϵ��,��ϵ�绰!,��ҵ</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
		End if
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
'���ô���������
Dim p_UserName,p_UserPassword,p_PassQuestion,p_PassAnswer,p_SafeCode,p_Email,p_IsCorporation,p_unMD5Password,LimitUserNameArr
p_UserName = NoHtmlHackInput(NoSqlHack(Request.Form("UserName")))
'if FiltBad1(p_UserName)=true then
'    strShowErr = "<li>�û����ؼ��ֲ�����</li>"
'    Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
'    Response.end
'end if

if p_isValidate = 1 then 
	Dim Randomizecode
	Randomizecode=GetRamCode(8)
	session("Tmp_p_password")  = MD5(Randomizecode,16)
	session("Tmp_password") = NoSqlHack(Randomizecode)
Else
	p_UserPassword = MD5(NoSqlHack(Request.Form("UserPassword")),16)
	p_unMD5Password = NoSqlHack(Request.Form("UserPassword"))
End if
p_PassQuestion = NoHtmlHackInput(Replace(Replace(Request.Form("PassQuestion"),"""",""),"'",""))
p_PassAnswer = Request.Form("PassAnswer")
p_SafeCode = MD5(Request.Form("SafeCode"),16)
p_Email = NoHtmlHackInput(Replace(Replace(Request.Form("Email"),"""",""),"'",""))
p_IsCorporation = cint(Request.Form("IsCorporation"))
If Trim(p_UserName)="" or Trim(p_Email)="" then
	strShowErr = "<li>����Ĳ����ύ!</li><li>�벻Ҫ���ⲿ�ύ����!</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
End if
if Fs_User.chk_regname(p_LimitUserName,p_UserName) = false then
	strShowErr = "<li>�û���������ע��</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
End if 
Dim GetUserTFObj
Set GetUserTFObj = server.CreateObject(G_FS_RS)
GetUserTFObj.open "select UserName,Email From FS_ME_Users where UserName = '"& NoSqlHack(p_UserName) &"' or Email='"&NoSqlHack(p_Email)&"'",User_Conn,1,3
If  Not GetUserTFObj.eof then
	strShowErr = "<li>����д���û�������ӵ����ʼ��Ѿ���ע�ᣬ��ʹ���û��������ʼ���ַ��������������ע�ᣡ</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
	GetUserTFObj.Close:set GetUserTFObj = nothing  
Else
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title>��Աע��step 3 of  4 step</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="keywords" content="��Ѷ,��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head>
<body oncontextmenu="return false;">
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="0">
  <tr> 
    <td><table width="100%" height="279" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
        <tr class="back"> 
          <td   colspan="2" class="xingmu" height="24">��User Register step 3.(��д��Ա��ϵ����)</td>
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
              <table width="96%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
              <form name="UserForm"  id="UserForm" method="post" action="Reg_contact.asp?forward=<%= forward %>"   onsubmit="return CheckForm();">
                <tr class="back"> 
                  <td height="20" colspan="3" class="xingmu">����д�������ϣ���������Ϣ����ͨ���ͷ�ȡ���ʺŵ�������ݣ�����ʵ��д�� 
                  </td>
                </tr>
                <tr class="back"> 
                  <td height="27"><div align="right"><span class="tx">*</span>�ǳ�</div></td>
                  <td><input name="NickName" type="text" id="NickName" size="20" maxlength="20"></td>
                  <td>����д��������ǳơ�����Ϊ����</td>
                </tr>
                <tr class="back"> 
                  <td width="15%" height="27"> <div align="right">����</div></td>
                  <td width="29%"><input name="RealName" type="text" id="RealName" size="20" maxlength="20"> 
                  </td>
                  <td width="56%">����д������ʵ������</td>
                </tr>
                <tr class="back"> 
                  <td height="16"><div align="right"><span class="tx">*</span>�Ա�</div></td>
                  <td><input type="radio" name="sex" value="0" checked="checked">
                    �� 
                    <input type="radio" name="sex" value="1">
                    Ů 
                    <input name="sex1" type="hidden" id="sex1"></td>
                  <td>����ѡ���Ա�</td>
                </tr>
                <tr class="back"> 
                  <td height="24"> <div align="right">��������</div></td>
                  <td><input name="Year" type="text" id="Year" value="19" size="5" maxlength="4">
                    �� 
                    <select name="month" id="month">
                      <option value="1" selected>1</option>
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
                      <option value="1" selected>1</option>
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
                  <td height="16"><div align="right"><span class="tx">*</span>֤�����</div></td>
                  <td> <select name=Certificate  id="Certificate">
                      <option value="0" selected>���֤</option>
                      <option value="2">ѧ��֤</option>
                      <option value="1">��ʻ֤</option>
                      <option value="3">����֤</option>
                      <option value="4">����</option>
                    </select> </td>
                  <td rowspan="2">��Ч֤����Ϊȡ���ʺŵ�����ֶΣ����Ժ�ʵ�ʺŵĺϷ���ݣ����������ʵ��д��<br> <span class="tx">�ر����ѣ���Ч֤��һ���趨�����ɸ���</span></td>
                </tr>
                <tr class="back"> 
                  <td height="16"><div align="right"><span class="tx">*</span>֤������</div></td>
                  <td><input name="CerTificateCode" type="text" id="CerTificateCode" size="30" maxlength="20"></td>
                </tr>
                <tr class="back"> 
                  <td height="16"><div align="right"><span class="tx">*</span>ʡ��</div></td>
                  <td height="16"><select name="Province" size=1 id="Province">
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
                    </select></td>
                  <td height="16" >���������ڵ�ʡ��</td>
                </tr>
                <tr class="back"> 
                  <td height="16"><div align="right"><span class="tx">*</span>����</div></td>
                  <td height="16"><input name="City" type="text" id="City" size="30" maxlength="20"></td>
                  <td height="16">���������ڵĳ���</td>
                </tr>
                <tr class="back"> 
                  <td height="16" colspan="3" class="xingmu">��дУ���룺</td>
                </tr>
                <tr class="back"> 
                  <td height="16"><div align="right"><span class="tx">*</span>У����</div></td>
                  <td><input name="VerCode" type="text" id="VerCode" size="15"> 
                    <img src="../Fs_Inc/vCode.asp?" onClick="this.src+=Math.random()" alt="ͼƬ�����壿������µõ���֤��,ע�⣺�����Сд" style="cursor:hand;"> </td>
                  <td>�뽫ͼ�������������������У��ò��������ڷ�ֹע�����<br> <span class="tx">�������ʱ��û�н��в����������֤�룬���»����֤���롣</span></td>
                </tr>
                <%if  p_IsCorporation =1 then%>
                <tr class="back"> 
                  <td height="16" colspan="3" class="xingmu">��д��˾���ϣ� </td>
                </tr>
                <tr class="back"> 
                  <td height="16"><div align="right"><span class="tx">*</span>��˾����</div></td>
                  <td><input name="C_Name" type="text" id="C_Name" size="30" maxlength="50"></td>
                  <td>����д����˾��ȫ��</td>
                </tr>
                <tr class="back"> 
                  <td height="-2"><div align="right">��˾���</div></td>
                  <td><input name="C_ShortName" type="text" id="C_ShortName" size="30" maxlength="30"></td>
                  <td>����д����˾�ļ򵥳ƺ�</td>
                </tr>
                <tr class="back"> 
                  <td height="0"><div align="right"><span class="tx">*</span>��˾����ʡ��</div></td>
                  <td> <select name="C_Province" size=1 id="C_Province">
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
                    </select> </td>
                  <td>����д����˾���ڵ�ʡ��</td>
                </tr>
                <tr class="back"> 
                  <td height="0"><div align="right"><span class="tx">*</span>��˾���ڳ���</div></td>
                  <td><input name="C_City" type="text" id="C_City" size="30" maxlength="20"></td>
                  <td>����д����˾���ڵĳ���</td>
                </tr>
                <tr class="back"> 
                  <td height="0"><div align="right"><span class="tx">*</span>��˾��ַ</div></td>
                  <td><input name="C_Address" type="text" id="C_Address" size="30" maxlength="100"></td>
                  <td>���Ĺ�˾��ַ</td>
                </tr>
                <tr class="back"> 
                  <td height="0"><div align="right"><span class="tx">*</span>��������</div></td>
                  <td><input name="C_PostCode" type="text" id="C_PostCode" size="30" maxlength="20"></td>
                  <td>����˾����������</td>
                </tr>
                <tr class="back"> 
                  <td height="0"><div align="right"><span class="tx">*</span>��˾��ϵ��</div></td>
                  <td><input name="C_ConactName" type="text" id="C_ConactName" size="30" maxlength="20"></td>
                  <td>��˾��ϵ��</td>
                </tr>
                <tr class="back"> 
                  <td height="0"><div align="right"><span class="tx">*</span>��˾��ϵ�绰</div></td>
                  <td><input name="C_Tel" type="text" id="C_Tel" size="30" maxlength="20"></td>
                  <td>��˾��ϵ�绰���зֻ�����&quot;-&quot;�ֿ����磺028-85098980-606</td>
                </tr>
                <tr class="back"> 
                  <td height="1"><div align="right">��˾����</div></td>
                  <td><input name="C_Fax" type="text" id="C_Fax" size="30" maxlength="20"></td>
                  <td>��˾���档�зֻ�����&quot;-&quot;�ֿ����磺028-85098980-606</td>
                </tr>
                <tr class="back"> 
                  <td height="3"><div align="right"><span class="tx">*</span>��ҵ</div></td>
                  <td><input name="C_VocationClassName" type="text" id="C_VocationClassName" size="30" readonly>
				  <input type="hidden" name="C_VocationClassID" id="C_VocationClassID">
				  </td>
                  <td><input type="button" name="Submit3" value="ѡ����ҵ" onClick="SelectClass();">
                    ��˾���ڵ���ҵ</td>
                </tr>
                <tr class="back"> 
                  <td height="8"><div align="right">��˾��վ</div></td>
                  <td><input name="C_Website" type="text" id="C_Website" size="30" maxlength="200"></td>
                  <td>��˾���ڵ���ҵվ��</td>
                </tr>
                <tr class="back"> 
                  <td height="16"><div align="right">��˾��ģ</div></td>
                  <td><select name="C_size" id="C_size">
                      <option value="1-20��" selected>1-20��</option>
                      <option value="21-50��">21-50��</option>
                      <option value="51-100��">51-100��</option>
                      <option value="101-200��">101-200��</option>
                      <option value="201-500��">201-500��</option>
                      <option value="501-1000��">501-1000��</option>
                      <option value="1000������">1000������</option>
                    </select></td>
                  <td>&nbsp;</td>
                </tr>
                <tr class="back"> 
                  <td height="1"><div align="right">��˾ע���ʱ�</div></td>
                  <td><select name="C_Capital" id="C_Capital">
                      <option value="10������">10������</option>
                      <option value="10��-19��">10��-19��</option>
                      <option value="20��-49��">20��-49��</option>
                      <option value="50��-99��" selected>50��-99��</option>
                      <option value="100��-199��">100��-199��</option>
                      <option value="200��-499��">200��-499��</option>
                      <option value="500��-999��">500��-999��</option>
                      <option value="1000������">1000������</option>
                    </select></td>
                  <td>&nbsp;</td>
                </tr>
                <tr class="back"> 
                  <td height="3"><div align="right">��������</div></td>
                  <td><input name="C_BankName" type="text" id="C_BankName" size="30" maxlength="50"></td>
                  <td rowspan="2"><p>��˾�����ʻ����Է������������ϵ�����С�<br>
                      �����������ӣ��й��������гɶ�����˫骷���<br>
                      �����ʻ�����</p></td>
                </tr>
                <tr class="back"> 
                  <td height="8"><div align="right">�����ʺż��ʻ���</div></td>
                  <td><textarea name="C_BankUserName" cols="30" rows="4" id="C_BankUserName"></textarea></td>
                </tr>
                <%End if%>
                <tr class="back"> 
                  <td height="39" colspan="3"> <div align="center"> 
                      <input name="Action" type="hidden" id="Action" value="SaveData">
                      <input name="UserName" type="hidden" id="UserName" value="<% = p_UserName %>">
                      <input name="UserPassword" type="hidden" id="UserPassword" value="<% = p_UserPassword %>">
                      <input name="PassQuestion" type="hidden" id="PassQuestion" value="<% = p_PassQuestion %>">
                      <input name="PassAnswer" type="hidden" id="PassAnswer" value="<% = p_PassAnswer %>">
                      <input name="SafeCode" type="hidden" id="SafeCode" value="<% = p_SafeCode %>">
                      <input name="Email" type="hidden" id="Email" value="<% = p_Email %>">
                      <input name="IsCorporation" type="hidden" id="IsCorporation" value="<% = p_IsCorporation %>">
                      <input name="unMD5Password" type="hidden" id="unMD5Password" value="<% = p_unMD5Password %>">
                      <input name="SubSys" type="hidden" id="SubSys" value="<% = Request.Form("SubSys")%>">
					  <input type="submit" name="Submit" value="��������,��ʼע��" style="CURSOR:hand">
                      �� 
                      <input type="reset" name="Submit2" value="����">
                      �� 
                      <input class="button" onClick="javascript:location.href='../'" type="button"  style="CURSOR:hand" value="������ҳ" name="Submit1" />
                    </div></td>
                </tr>
              </form>
            </table>
            </td>
        </tr>
        <tr class="back"> 
          <td height="26"  colspan="2" class="xingmu"> <div align="left"> 
              <!--#include file="Copyright.asp" -->
            </div></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
<%
End if
set Fs_User = nothing
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->
<script language="JavaScript" type="text/javascript">
function CheckForm()
{
	if(document.UserForm.NickName.value=="")
	{
		alert("�����������ǳ�!");
		document.UserForm.NickName.focus();
		return false;
	}
	if(document.UserForm.Province.value=="")
	{
		alert("������ʡ��!");
		document.UserForm.Province.focus();
		return false;
	}
	if(document.UserForm.City.value=="")
	{
		alert("���������!");
		document.UserForm.City.focus();
		return false;
	}
	if( !(UserForm.sex[0].checked || UserForm.sex[1].checked)) {
	alert("��ѡ���Ա� !");
	return false;
	}
	if(document.UserForm.VerCode.value=="")
	{
		alert("��������֤��!");
		document.UserForm.VerCode.focus();
		return false;
	}
}
	
function SelectClass()
{
	var ReturnValue='',TempArray=new Array();
	ReturnValue = OpenWindow('lib/SelectClassFrame.asp',400,300,window);
	if (ReturnValue.indexOf('***')!=-1)
	{
		TempArray = ReturnValue.split('***');
		document.all.C_VocationClassID.value=TempArray[0]
		document.all.C_VocationClassName.value=TempArray[1]
	}
}
	
</script>
<script language="JavaScript" type="text/JavaScript">
	function OpenWindow(Url,Width,Height,WindowObj)
	{
		var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:no;');
		return ReturnStr;
	}
</script>