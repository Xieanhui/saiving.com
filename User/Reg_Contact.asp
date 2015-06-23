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
	strShowErr = "<li>暂时关闭注册功能</li><li>或者系统参数丢失!</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
End If
Dim forward
forward = Request.QueryString("forward")
forward = Server.URLEncode(forward)
'保存入数据库
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
			strShowErr = "<li>验证码不正确！</li><li>长时间没动作,请点验证刷新一次再输入!</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
	End if
	if Request.Form("IsCorporation") = "1" then
		If Trim(p_C_Name)="" Or Trim(p_C_Province)="" Or Trim(p_C_City)="" Or Trim(p_C_PostCode)=""  Or Trim(p_C_ConactName)="" Or Trim(p_C_Tel)="" or trim(p_C_VocationClassID)="" then
				strShowErr = "<li>企业注册需要填写完整资料！</li><li>请填写完整,企业名称,所在省份,城市,邮编,联系人,联系电话!,行业</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
		End if
		Dim AddCorpDataTFObj
		Set AddCorpDataTFObj = server.CreateObject(G_FS_RS)
		AddCorpDataTFObj.open "select  C_Name From FS_ME_CorpUser where C_Name = '"& p_C_Name &"'",User_Conn,1,3
		If Not AddCorpDataTFObj.eof then
					strShowErr = "<li>您提交的企业名称已经被注册！</li>"
					Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
					Response.end
		End if
		AddCorpDataTFObj.close:set AddCorpDataTFObj = nothing
	End if
	strShowErr = ""
	'-----------------------------------------------------------------
	'系统整合
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
		strShowErr = strShowErr & "<li>您提交的用户名或者电子邮件已经被注册!</li>"
	End if
	AddUserDataTFObj.close:set AddUserDataTFObj =nothing
	'判断用户编号是否重复
	Set UserNumberRuleObj = server.CreateObject(G_FS_RS)
	UserNumberRuleObj.open "select UserNumber From FS_ME_Users where UserName='"& p_UserName_1&"'",User_Conn,1,1
	If Not UserNumberRuleObj.eof then
		strShowErr = strShowErr & "<li>您提交的用户编号意外重复，非常抱歉，请重新填写用户资料。!</li>"
	End if
	If strShowErr<>"" Then
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Reg_info.asp")
		Response.End()
	End If
	'-----------------------------------------------------------------
	'系统整合
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
	'插入用户数据

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
	'更新数据，获得相应期限或者金币，积分
	'说明：如果为0，则不限制
	'开始建立对象
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
	Call Fs_User.AddLog("注册",session("FS_UserNumber"),p_NumGetPoint,p_NumGetMoney,"注册获得积分",0) 
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
'调用传来的数据
Dim p_UserName,p_UserPassword,p_PassQuestion,p_PassAnswer,p_SafeCode,p_Email,p_IsCorporation,p_unMD5Password,LimitUserNameArr
p_UserName = NoHtmlHackInput(NoSqlHack(Request.Form("UserName")))
'if FiltBad1(p_UserName)=true then
'    strShowErr = "<li>用户名关键字不允许</li>"
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
	strShowErr = "<li>错误的参数提交!</li><li>请不要从外部提交数据!</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
End if
if Fs_User.chk_regname(p_LimitUserName,p_UserName) = false then
	strShowErr = "<li>用户名不允许注册</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
End if 
Dim GetUserTFObj
Set GetUserTFObj = server.CreateObject(G_FS_RS)
GetUserTFObj.open "select UserName,Email From FS_ME_Users where UserName = '"& NoSqlHack(p_UserName) &"' or Email='"&NoSqlHack(p_Email)&"'",User_Conn,1,3
If  Not GetUserTFObj.eof then
	strShowErr = "<li>您填写的用户名或电子电子邮件已经被注册，请使用用户名检查和邮件地址检查来辅助您完成注册！</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
	GetUserTFObj.Close:set GetUserTFObj = nothing  
Else
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title>会员注册step 3 of  4 step</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="keywords" content="风讯,风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head>
<body oncontextmenu="return false;">
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="0">
  <tr> 
    <td><table width="100%" height="279" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
        <tr class="back"> 
          <td   colspan="2" class="xingmu" height="24">·User Register step 3.(填写会员联系资料)</td>
        </tr>
        <tr class="back"> 
          <td width="15%" valign="top" class="hback"><strong>【注册步骤】</strong> <br>
            <br>
            <div align="left"> √同意注册协议<br>
              <br>
              √填写会员资料<br>
              <br>
              →填写联系资料<br>
              <br>
              ×注册成功</div>
            </td>
          <td width="86%" valign="top" class="hback"> 
              <table width="96%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
              <form name="UserForm"  id="UserForm" method="post" action="Reg_contact.asp?forward=<%= forward %>"   onsubmit="return CheckForm();">
                <tr class="back"> 
                  <td height="20" colspan="3" class="xingmu">请填写个人资料：（以下信息是您通过客服取回帐号的最后依据，请如实填写） 
                  </td>
                </tr>
                <tr class="back"> 
                  <td height="27"><div align="right"><span class="tx">*</span>昵称</div></td>
                  <td><input name="NickName" type="text" id="NickName" size="20" maxlength="20"></td>
                  <td>请填写您对外的昵称。可以为中文</td>
                </tr>
                <tr class="back"> 
                  <td width="15%" height="27"> <div align="right">姓名</div></td>
                  <td width="29%"><input name="RealName" type="text" id="RealName" size="20" maxlength="20"> 
                  </td>
                  <td width="56%">请填写您的真实姓名。</td>
                </tr>
                <tr class="back"> 
                  <td height="16"><div align="right"><span class="tx">*</span>性别</div></td>
                  <td><input type="radio" name="sex" value="0" checked="checked">
                    男 
                    <input type="radio" name="sex" value="1">
                    女 
                    <input name="sex1" type="hidden" id="sex1"></td>
                  <td>请您选择性别。</td>
                </tr>
                <tr class="back"> 
                  <td height="24"> <div align="right">出生日期</div></td>
                  <td><input name="Year" type="text" id="Year" value="19" size="5" maxlength="4">
                    年 
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
                    月 
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
                    日 </td>
                  <td>请填写您的真实生日，该项用于取回密码。</td>
                </tr>
                <tr class="back"> 
                  <td height="16"><div align="right"><span class="tx">*</span>证件类别</div></td>
                  <td> <select name=Certificate  id="Certificate">
                      <option value="0" selected>身份证</option>
                      <option value="2">学生证</option>
                      <option value="1">驾驶证</option>
                      <option value="3">军人证</option>
                      <option value="4">护照</option>
                    </select> </td>
                  <td rowspan="2">有效证件作为取回帐号的最后手段，用以核实帐号的合法身份，请您务必如实填写。<br> <span class="tx">特别提醒：有效证件一旦设定，不可更改</span></td>
                </tr>
                <tr class="back"> 
                  <td height="16"><div align="right"><span class="tx">*</span>证件号码</div></td>
                  <td><input name="CerTificateCode" type="text" id="CerTificateCode" size="30" maxlength="20"></td>
                </tr>
                <tr class="back"> 
                  <td height="16"><div align="right"><span class="tx">*</span>省份</div></td>
                  <td height="16"><select name="Province" size=1 id="Province">
                      <option value="">请选择　　</option>
                      <option value="四川">四川</option>
                      <option value="北京">北京</option>
                      <option value="上海">上海</option>
                      <option value="天津">天津</option>
                      <option value="重庆">重庆</option>
                      <option value="安徽">安徽</option>
                      <option value="甘肃">甘肃</option>
                      <option value="广东">广东</option>
                      <option value="广西">广西</option>
                      <option value="贵州">贵州</option>
                      <option value="福建">福建</option>
                      <option value="海南">海南</option>
                      <option value="河北">河北</option>
                      <option value="河南">河南</option>
                      <option value="黑龙江">黑龙江</option>
                      <option value="湖北">湖北</option>
                      <option value="湖南">湖南</option>
                      <option value="吉林">吉林</option>
                      <option value="江苏">江苏</option>
                      <option value="江西">江西</option>
                      <option value="辽宁">辽宁</option>
                      <option value="内蒙古">内蒙古</option>
                      <option value="宁夏">宁夏</option>
                      <option value="青海">青海</option>
                      <option value="山东">山东</option>
                      <option value="山西">山西</option>
                      <option value="陕西">陕西</option>
                      <option value="西藏">西藏</option>
                      <option value="新疆">新疆</option>
                      <option value="云南">云南</option>
                      <option value="浙江">浙江</option>
                      <option value="港澳台">港澳台</option>
                      <option value="海外">海外</option>
                      <option value="其它">其它</option>
                    </select></td>
                  <td height="16" >您现在所在的省份</td>
                </tr>
                <tr class="back"> 
                  <td height="16"><div align="right"><span class="tx">*</span>城市</div></td>
                  <td height="16"><input name="City" type="text" id="City" size="30" maxlength="20"></td>
                  <td height="16">您现在所在的城市</td>
                </tr>
                <tr class="back"> 
                  <td height="16" colspan="3" class="xingmu">填写校验码：</td>
                </tr>
                <tr class="back"> 
                  <td height="16"><div align="right"><span class="tx">*</span>校验码</div></td>
                  <td><input name="VerCode" type="text" id="VerCode" size="15"> 
                    <img src="../Fs_Inc/vCode.asp?" onClick="this.src+=Math.random()" alt="图片看不清？点击重新得到验证码,注意：区别大小写" style="cursor:hand;"> </td>
                  <td>请将图中数字填入左边输入框中，该步骤有利于防止注册机。<br> <span class="tx">如果您长时间没有进行操作，请点验证码，重新获得验证代码。</span></td>
                </tr>
                <%if  p_IsCorporation =1 then%>
                <tr class="back"> 
                  <td height="16" colspan="3" class="xingmu">填写公司资料： </td>
                </tr>
                <tr class="back"> 
                  <td height="16"><div align="right"><span class="tx">*</span>公司名称</div></td>
                  <td><input name="C_Name" type="text" id="C_Name" size="30" maxlength="50"></td>
                  <td>请填写您公司的全称</td>
                </tr>
                <tr class="back"> 
                  <td height="-2"><div align="right">公司简称</div></td>
                  <td><input name="C_ShortName" type="text" id="C_ShortName" size="30" maxlength="30"></td>
                  <td>请填写您公司的简单称呼</td>
                </tr>
                <tr class="back"> 
                  <td height="0"><div align="right"><span class="tx">*</span>公司所在省份</div></td>
                  <td> <select name="C_Province" size=1 id="C_Province">
                      <option value="">请选择　　</option>
                      <option value="四川">四川</option>
                      <option value="北京">北京</option>
                      <option value="上海">上海</option>
                      <option value="天津">天津</option>
                      <option value="重庆">重庆</option>
                      <option value="安徽">安徽</option>
                      <option value="甘肃">甘肃</option>
                      <option value="广东">广东</option>
                      <option value="广西">广西</option>
                      <option value="贵州">贵州</option>
                      <option value="福建">福建</option>
                      <option value="海南">海南</option>
                      <option value="河北">河北</option>
                      <option value="河南">河南</option>
                      <option value="黑龙江">黑龙江</option>
                      <option value="湖北">湖北</option>
                      <option value="湖南">湖南</option>
                      <option value="吉林">吉林</option>
                      <option value="江苏">江苏</option>
                      <option value="江西">江西</option>
                      <option value="辽宁">辽宁</option>
                      <option value="内蒙古">内蒙古</option>
                      <option value="宁夏">宁夏</option>
                      <option value="青海">青海</option>
                      <option value="山东">山东</option>
                      <option value="山西">山西</option>
                      <option value="陕西">陕西</option>
                      <option value="西藏">西藏</option>
                      <option value="新疆">新疆</option>
                      <option value="云南">云南</option>
                      <option value="浙江">浙江</option>
                      <option value="港澳台">港澳台</option>
                      <option value="海外">海外</option>
                      <option value="其它">其它</option>
                    </select> </td>
                  <td>请填写您公司所在的省份</td>
                </tr>
                <tr class="back"> 
                  <td height="0"><div align="right"><span class="tx">*</span>公司所在城市</div></td>
                  <td><input name="C_City" type="text" id="C_City" size="30" maxlength="20"></td>
                  <td>请填写您公司所在的城市</td>
                </tr>
                <tr class="back"> 
                  <td height="0"><div align="right"><span class="tx">*</span>公司地址</div></td>
                  <td><input name="C_Address" type="text" id="C_Address" size="30" maxlength="100"></td>
                  <td>您的公司地址</td>
                </tr>
                <tr class="back"> 
                  <td height="0"><div align="right"><span class="tx">*</span>邮政编码</div></td>
                  <td><input name="C_PostCode" type="text" id="C_PostCode" size="30" maxlength="20"></td>
                  <td>您公司的邮政编码</td>
                </tr>
                <tr class="back"> 
                  <td height="0"><div align="right"><span class="tx">*</span>公司联系人</div></td>
                  <td><input name="C_ConactName" type="text" id="C_ConactName" size="30" maxlength="20"></td>
                  <td>公司联系人</td>
                </tr>
                <tr class="back"> 
                  <td height="0"><div align="right"><span class="tx">*</span>公司联系电话</div></td>
                  <td><input name="C_Tel" type="text" id="C_Tel" size="30" maxlength="20"></td>
                  <td>公司联系电话。有分机请用&quot;-&quot;分开，如：028-85098980-606</td>
                </tr>
                <tr class="back"> 
                  <td height="1"><div align="right">公司传真</div></td>
                  <td><input name="C_Fax" type="text" id="C_Fax" size="30" maxlength="20"></td>
                  <td>公司传真。有分机请用&quot;-&quot;分开，如：028-85098980-606</td>
                </tr>
                <tr class="back"> 
                  <td height="3"><div align="right"><span class="tx">*</span>行业</div></td>
                  <td><input name="C_VocationClassName" type="text" id="C_VocationClassName" size="30" readonly>
				  <input type="hidden" name="C_VocationClassID" id="C_VocationClassID">
				  </td>
                  <td><input type="button" name="Submit3" value="选择行业" onClick="SelectClass();">
                    公司所在的行业</td>
                </tr>
                <tr class="back"> 
                  <td height="8"><div align="right">公司网站</div></td>
                  <td><input name="C_Website" type="text" id="C_Website" size="30" maxlength="200"></td>
                  <td>公司所在的企业站点</td>
                </tr>
                <tr class="back"> 
                  <td height="16"><div align="right">公司规模</div></td>
                  <td><select name="C_size" id="C_size">
                      <option value="1-20人" selected>1-20人</option>
                      <option value="21-50人">21-50人</option>
                      <option value="51-100人">51-100人</option>
                      <option value="101-200人">101-200人</option>
                      <option value="201-500人">201-500人</option>
                      <option value="501-1000人">501-1000人</option>
                      <option value="1000人以上">1000人以上</option>
                    </select></td>
                  <td>&nbsp;</td>
                </tr>
                <tr class="back"> 
                  <td height="1"><div align="right">公司注册资本</div></td>
                  <td><select name="C_Capital" id="C_Capital">
                      <option value="10万以下">10万以下</option>
                      <option value="10万-19万">10万-19万</option>
                      <option value="20万-49万">20万-49万</option>
                      <option value="50万-99万" selected>50万-99万</option>
                      <option value="100万-199万">100万-199万</option>
                      <option value="200万-499万">200万-499万</option>
                      <option value="500万-999万">500万-999万</option>
                      <option value="1000万以上">1000万以上</option>
                    </select></td>
                  <td>&nbsp;</td>
                </tr>
                <tr class="back"> 
                  <td height="3"><div align="right">开户银行</div></td>
                  <td><input name="C_BankName" type="text" id="C_BankName" size="30" maxlength="50"></td>
                  <td rowspan="2"><p>公司银行帐户，以方便放在您的联系资料中。<br>
                      开户银行例子：中国工商银行成都分行双楠分理处<br>
                      银行帐户名：</p></td>
                </tr>
                <tr class="back"> 
                  <td height="8"><div align="right">银行帐号及帐户名</div></td>
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
					  <input type="submit" name="Submit" value="保存数据,开始注册" style="CURSOR:hand">
                      　 
                      <input type="reset" name="Submit2" value="重置">
                      　 
                      <input class="button" onClick="javascript:location.href='../'" type="button"  style="CURSOR:hand" value="返回首页" name="Submit1" />
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
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->
<script language="JavaScript" type="text/javascript">
function CheckForm()
{
	if(document.UserForm.NickName.value=="")
	{
		alert("请输入您的昵称!");
		document.UserForm.NickName.focus();
		return false;
	}
	if(document.UserForm.Province.value=="")
	{
		alert("请输入省份!");
		document.UserForm.Province.focus();
		return false;
	}
	if(document.UserForm.City.value=="")
	{
		alert("请输入城市!");
		document.UserForm.City.focus();
		return false;
	}
	if( !(UserForm.sex[0].checked || UserForm.sex[1].checked)) {
	alert("请选择性别 !");
	return false;
	}
	if(document.UserForm.VerCode.value=="")
	{
		alert("请输入验证码!");
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