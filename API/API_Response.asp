<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Md5.asp" -->
<!--#include file="Api_Config.asp"-->

<%
Dim XMLDom,XmlDoc,Node,Status,Messenge,UserTrueIP
Dim UserName,Act,appid
Status = 1
Messenge = ""
Dim Conn,User_Conn
MF_Default_Conn
MF_User_Conn

If Request.QueryString<>"" Then
	SaveUserCookie()
Else
	Set XmlDoc = Server.CreateObject(G_MSXML2_DOCUMENT & MsxmlVersion)
	XmlDoc.ASYNC = False
	If Not XmlDoc.LOAD(Request) Then
		Status = 1
		Messenge = "数据非法，操作中止！"
		appid = "未知"
	Else
		If Not (XmlDoc.documentElement.selectSingleNode("userip") is nothing) Then
			UserTrueIP = XmlDoc.documentElement.selectSingleNode("userip").text
		End If
		If CheckPost() Then
			Select Case Act
				Case "checkname"
					Checkname()
				Case "reguser"
					Reguser()
				Case "login"
					UesrLogin()
				Case "logout"
					LogoutUser()
				Case "update"
					UpdateUser()
				Case "delete"
					Deleteuser()
				Case "lock"
					Lockuser()
				Case "getinfo"
					GetUserinfo()
			End Select
		End If
	End If
	ReponseData()
	Set XmlDoc = Nothing
End If
Set FoosunCMS = Nothing

Sub ReponseData()
	'XmlDoc.loadxml "<root><appid>FoosunCMS</appid><status>0</status><body><message/><email/><question/><answer/><savecookie/><truename/><gender/><birthday/><qq/><msn/><mobile/><telephone/><address/><zipcode/><homepage/><userip/><jointime/><experience/><ticket/><valuation/><balance/><posts/><userstatus/></body></root>"
	If Act <> "getinfo" Then
		XmlDoc.loadxml "<root><appid>FoosunCMS</appid><status>0</status><body><message/></body></root>"
	End If
	XmlDoc.documentElement.selectSingleNode("appid").text = "FoosunCMS"
	XmlDoc.documentElement.selectSingleNode("status").text = status
	XmlDoc.documentElement.selectSingleNode("body/message").text = ""
	Set Node = XmlDoc.createCDATASection(Replace(Messenge,"]]>","]]&gt;"))
	XmlDoc.documentElement.selectSingleNode("body/message").appendChild(Node)
	Response.Clear
	Response.ContentType="text/xml"
	Response.CharSet="gb2312"
	Response.Write "<?xml version=""1.0"" encoding=""gb2312""?>"&vbNewLine
	Response.Write XmlDoc.documentElement.XML
End Sub


Function CheckPost()
	CheckPost = False
	Dim Syskey
	If XmlDoc.documentElement.selectSingleNode("action") is Nothing or XmlDoc.documentElement.selectSingleNode("syskey") is Nothing or XmlDoc.documentElement.selectSingleNode("username")  is Nothing Then
		Status = 1
		Messenge = Messenge & "<li>非法请求。"
		Exit Function
	End If
	UserName =NoSqlHack(XmlDoc.documentElement.selectSingleNode("username").text)
	Syskey = NoSqlHack(XmlDoc.documentElement.selectSingleNode("syskey").text)
	Act = NoSqlHack(XmlDoc.documentElement.selectSingleNode("action").text)
	Appid = NoSqlHack(XmlDoc.documentElement.selectSingleNode("appid").text)
	
	Dim NewMd5,OldMd5
	NewMd5 = Md5(UserName&API_SysKey,16)
	OldMd5 = Md5(UserName&API_SysKey,16)

	If Syskey=NewMd5 or Syskey=OldMd5 Then
		CheckPost = True
	Else
		Status = 1
		Messenge = Messenge & "<li>请求数据验证不通过，请与管理员联系。"
	End If
End Function


Sub GetUserinfo()
	Dim Rs,Sql
	Dim Userinfo,UserIM
	
	XmlDoc.loadxml "<root><appid>FoosunCMS</appid><status>0</status><body><message/><email/><question/><answer/><savecookie/><truename/><gender/><birthday/><qq/><msn/><mobile/><telephone/><address/><zipcode/><homepage/><userip/><jointime/><experience/><ticket/><valuation/><balance/><posts/><userstatus/></body></root>"
	
	Sql = "Select Top 1 * From FS_ME_Users Where UserName='"&NoSqlHack(UserName)&"'"
	Set Rs = User_Conn.Execute(Sql)
	If Not Rs.Eof And Not Rs.Bof Then
		XmlDoc.documentElement.selectSingleNode("body/email").text = Rs("Email")&""
		XmlDoc.documentElement.selectSingleNode("body/question").text = Rs("PassQuestion")&""
		XmlDoc.documentElement.selectSingleNode("body/answer").text = Rs("PassAnswer")&""
		XmlDoc.documentElement.selectSingleNode("body/gender").text = Rs("Sex")&""
		XmlDoc.documentElement.selectSingleNode("body/birthday").text = Rs("BothYear")&""
		XmlDoc.documentElement.selectSingleNode("body/mobile").text = Rs("Mobile")&""
		XmlDoc.documentElement.selectSingleNode("body/userip").text = Rs("LastLoginIP")&""
		XmlDoc.documentElement.selectSingleNode("body/jointime").text = Rs("RegTime")&""
		XmlDoc.documentElement.selectSingleNode("body/zipcode").text = Rs("PostCode")&""
		'XmlDoc.documentElement.selectSingleNode("body/experience").text = Rs("userEP")&""
		XmlDoc.documentElement.selectSingleNode("body/ticket").text = Rs("Integral")&""
		'XmlDoc.documentElement.selectSingleNode("body/valuation").text = Rs("userCP")&""
		XmlDoc.documentElement.selectSingleNode("body/balance").text = Rs("FS_Money")&""
		XmlDoc.documentElement.selectSingleNode("body/posts").text = Rs("PostCode")&""
		XmlDoc.documentElement.selectSingleNode("body/userstatus").text = Rs("isLock")&""
		XmlDoc.documentElement.selectSingleNode("body/homepage").text = Rs("HomePage")&""
		XmlDoc.documentElement.selectSingleNode("body/qq").text = Rs("QQ")&""
		XmlDoc.documentElement.selectSingleNode("body/msn").text = Rs("MSN")&""
		XmlDoc.documentElement.selectSingleNode("body/truename").text = Rs("RealName")&""
		XmlDoc.documentElement.selectSingleNode("body/telephone").text = Rs("tel")&""
		XmlDoc.documentElement.selectSingleNode("body/address").text = Rs("Address")&""
		Status = 0
		Messenge = Messenge & "<li>读取用户资料成功。"
	Else
		Status = 1
		Messenge = Messenge & "<li>该用户不存在。"
	End If
	Rs.Close
	Set Rs = Nothing
End Sub


Sub Deleteuser()
	Dim D_Users,i
	Dim Rs
	D_Users = Split(UserName,",")
	Messenge = UserName
	For i=0 To UBound(D_Users)
		Set Rs=User_Conn.Execute("Select UserName,UserNumber from [FS_ME_Users] where UserName='"&NOSQLHACK(D_Users(i))&"'")
		If not (rs.eof and rs.bof) then
			User_Conn.Execute("Update FS_ME_Message set M_FromUserNumber=0 where M_FromUserNumber='"&rs(1)&"'")
			User_Conn.Execute("Update FS_ME_Message set M_ReadUserNumber=0 where M_ReadUserNumber='"&rs(1)&"'")
			User_Conn.Execute("Delete from FS_ME_BuyBag where UserNumber='"&rs(1)&"'")
			User_Conn.Execute("Update FS_ME_Card set UserNumber=0 where UserNumber='"&rs(1)&"'")
			User_Conn.Execute("Delete from FS_ME_CertFile where UserNumber='"&rs(1)&"'")
			User_Conn.Execute("Delete from FS_ME_CorpUser where UserNumber='"&rs(1)&"'")
			User_Conn.Execute("Delete from FS_ME_Favorite where UserNumber='"&rs(1)&"'")
			User_Conn.Execute("Delete from FS_ME_FavoriteClass where UserNumber='"&rs(1)&"'")
			User_Conn.Execute("Delete from FS_ME_Friends where UserNumber='"&rs(1)&"'")
			User_Conn.Execute("Update FS_ME_Friends set F_UserNumber = 0  where F_UserNumber='"&rs(1)&"'")
			User_Conn.Execute("Delete from FS_ME_GroupDebate where UserNumber='"&rs(1)&"'")
			User_Conn.Execute("Delete from FS_ME_InfoClass where UserNumber='"&rs(1)&"'")
			User_Conn.Execute("Delete from FS_ME_InfoContribution where UserNumber='"&rs(1)&"'")
			User_Conn.Execute("Delete from FS_ME_InfoDown where UserNumber='"&rs(1)&"'")
			User_Conn.Execute("Delete from FS_ME_Infoilog where UserNumber='"&rs(1)&"'")
			User_Conn.Execute("Delete from FS_ME_InfoProduct where UserNumber='"&rs(1)&"'")
			User_Conn.Execute("Delete from FS_ME_Log where UserNumber='"&rs(1)&"'")
			User_Conn.Execute("Delete from FS_ME_MyInfo where UserNumber='"&rs(1)&"'")
			User_Conn.Execute("Delete from FS_ME_MySysPara where UserNumber='"&rs(1)&"'")
			User_Conn.Execute("Delete from FS_ME_Order where UserNumber='"&rs(1)&"'")
			User_Conn.Execute("Delete from FS_ME_Review where UserNumber='"&rs(1)&"'")
			Conn.Execute("Delete from FS_SD_News where UserNumber='"&rs(1)&"'")
			User_Conn.Execute("Delete from FS_ME_Users where UserName='"&D_Users(i)&"'")
			Messenge = Messenge & "<li>用户（"&D_Users(i)&"）删除成功。"
		End If
		rs.close
	Next
	Status = 0
End Sub


Sub SaveUserCookie()
	Dim S_syskey,Password,SaveCookie,TruePassWord,userclass,Userhidden
	S_syskey = NoSqlHack(Request.QueryString("syskey"))
	UserName = NoSqlHack(Request.QueryString("UserName"))
	Password = Request.QueryString("Password")
	SaveCookie = Request.QueryString("savecookie")
	If UserName="" or S_syskey="" Then Exit Sub
	Dim NewMd5,OldMd5
	NewMd5 = Md5(UserName&API_SysKey,16)
	OldMd5 = Md5(UserName&API_SysKey,16)
	If Not (S_syskey=NewMd5 or S_syskey=OldMd5) Then
		Exit Sub
	End If
	If SaveCookie="" or Not IsNumeric(SaveCookie) Then SaveCookie = 0
	'用户退出
	If Password = "" Then
		Session("FS_UserName") = ""
		Session("FS_UserNumber") = ""
		Session("FS_UserPassword") = ""
		Session("FS_Group") = ""
		Session("FS_IsCorp") = ""
		Session("FS_NickName") = ""
		response.Cookies("FoosunUserCookies")("UserLogin_Style_Num")  = ""
		Session("UserLoginCode") = ""
		Exit Sub
	End If

	'用户登陆
	'Password = Md5(Password,16)
	Dim f_RsLoginobj,Sql
	If Not IsObject(User_Conn) Then MF_User_Conn
	Set f_RsLoginobj = Server.CreateObject(G_FS_RS)
	Sql = "Select Top 1 UserName,UserNumber,NickName,UserPassword,Email,isLock,IsCorporation,GroupID,MySkin,TempLastLoginTime From FS_ME_Users Where UserName = '"& NoSqlHack(UserName) &"'"
	f_RsLoginobj.Open Sql,User_Conn,1,3
	If Not f_RsLoginobj.Eof And Not f_RsLoginobj.Bof Then
		If f_RsLoginobj(3)<>Password Then
			Exit Sub
		End If
		session("FS_UserName") = f_RsLoginobj(0)
		session("FS_UserNumber") = f_RsLoginobj(1)
		session("FS_NickName") = f_RsLoginobj(2)
		session("FS_UserPassword") = f_RsLoginobj(3)
		session("FS_UserEmail")  = f_RsLoginobj(4)'改为Cookies
		session("FS_UserGroupID") = f_RsLoginobj(7)
		response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") = f_RsLoginobj(8)'改为Cookies
	Else
		Exit Sub
	End If
	f_RsLoginobj.Close
	Set f_RsLoginobj = Nothing
	'Response.Write "document.write(""OK"");"
End Sub


Sub Checkname()
	Dim UserEmail
	Dim Temp_tr,i,Rs,Sql
	UserEmail = Trim(XmlDoc.documentElement.selectSingleNode("email").text)
	If Messenge<>"" Then
		'输出错误信息
		Status = 1
		Exit Sub
	End If
	Sql="select UserName,Email From FS_ME_Users where UserName = '"& NoSqlHack(UserName) &"'"
	Set Rs = User_Conn.Execute(Sql)
	If Not Rs.Eof And Not Rs.Bof Then
		Messenge = "您填写的用户名已经被注册。"
		Status = 1
		Exit Sub
	Else
		Status = 0
		Messenge = "验证通过。"
	End If
	Rs.Close
	Set Rs = Nothing
End Sub

'用户注册
Sub Reguser()
	Dim UserPass,UserEmail,Question,Answer,usercookies,truename,gender,birthday
	Dim Temp_tr,i
	UserPass = NoSqlHack(XmlDoc.documentElement.selectSingleNode("password").text)
	UserEmail = NoSqlHack(XmlDoc.documentElement.selectSingleNode("email").text)
	Question = NoSqlHack(XmlDoc.documentElement.selectSingleNode("question").text)
	Answer = NoSqlHack(XmlDoc.documentElement.selectSingleNode("answer").text)
	truename = NoSqlHack(XmlDoc.documentElement.selectSingleNode("truename").text)
	'gender = XmlDoc.documentElement.selectSingleNode("gender").text
	birthday = NoSqlHack(XmlDoc.documentElement.selectSingleNode("birthday").text)
	gender = 0
	usercookies = 1
	If UserName="" or UserPass="" or Question="" or Answer = "" Then
		Status = 1
		Messenge = Messenge & "<li>请填写用户名或密码。"
		Exit Sub
	End If
	UserPass = Md5(UserPass,16)
	Answer = Md5(Answer,16)
	strUserNumberRule=GetRamCode(13)'无需过滤
	'信息验证
	If Messenge<>"" Then
		'输出错误信息
		Status = 1
		Exit Sub
	End If

	Dim Rs,Sql,AllowRegister,RegisterAudit,DefaultGroupID,RegMoney,RegPoint
	Set Rs = User_Conn.Execute("Select top 1 RegisterTF,RegPointmoney,RegisterCheck,DefaultGroupID From FS_ME_SysPara")
	If Not Rs.Eof And Not Rs.Bof Then
		if not isnull(Rs("RegisterTF")) then AllowRegister = Rs("RegisterTF")
		if not isnull(Rs("RegisterCheck")) then RegisterAudit = Rs("RegisterCheck")
		if not isnull(Rs("DefaultGroupID")) then DefaultGroupID = Rs("DefaultGroupID")
		if not isnull(Rs("RegPointmoney")) then
			if instr(Rs("RegPointmoney"),",") then
				RegPoint=split(Rs("RegPointmoney"),",")(0)
				RegMoney=split(Rs("RegPointmoney"),",")(1)
			End If
		End If
	Else
		Messenge = "系统未初始化。"
		Status = 1
		Exit Sub
	End If
	If AllowRegister<>1 Then
		Messenge = "注册功能已关闭。"
		Status = 1
		Exit Sub
	End If
	
	Set Rs = User_Conn.Execute("Select UserName From FS_ME_Users where UserNumber = '"& strUserNumberRule &"'")
	If Not Rs.Eof And Not Rs.Bof Then
		Messenge = "用户编号意外重复，请重试。"
		Status = 1
		Exit Sub
	End If
	Set Rs = Server.CreateObject(G_FS_RS)
	Sql="select * From FS_ME_Users where UserName = '"& NoSqlHack(UserName) &"'"
	Rs.Open Sql,User_Conn,1,3
	If Not Rs.Eof And Not Rs.Bof Then
		Messenge = "您填写的用户名已经被注册。"
		Status = 1
		Exit Sub
	Else
		Status = 0
		Rs.AddNew
		Rs("UserNumber") = NoSqlHack(strUserNumberRule)
		Rs("UserName") = NoSqlHack(UserName)
		Rs("UserPassword") = NoSqlHack(UserPass)
		Rs("Email") = NoSqlHack(UserEmail)
		Rs("PassQuestion") = NoSqlHack(Question)
		Rs("PassAnswer") = NoSqlHack(Answer)
		Rs("RealName") = NoSqlHack(truename)
		Rs("Sex") = NoSqlHack(gender)
		If birthday="" Then
			Rs("BothYear") = Null
		Else
			Rs("BothYear") = NoSqlHack(birthday)
		End If
		Rs("safeCode") = NoSqlHack(Answer)
		Rs("isMessage") = 0
		Rs("HeadPicsize") = "60,60"
		Rs("NickName") = NoSqlHack(UserName)
		Rs("Province") = ""
		Rs("city") = ""
		Rs("Certificate") = 0
		Rs("CertificateCode") = ""
		Rs("IsCorporation") = 0
		Rs("RegTime") = now
		Rs("CloseTime") = "3000-1-1"
		Rs("LoginNum") = 0
		Rs("Integral") = NoSqlHack(RegPoint)
		Rs("FS_Money") = NoSqlHack(RegMoney)
		Rs("TempLastLoginTime") = now
		Rs("TempLastLoginTime_1") = now
		If RegisterAudit = 1 Then
			Rs("isLock") = 1
		Else
			Rs("isLock") = 0
		End If
		Rs("MySkin") = 2
		Rs("OnlyLogin") = 0
		Rs("ConNumber") = 0
		Rs("ConNumberNews") = 0
		Rs("isOpen") = 0
		Rs("GroupID") = NoSqlHack(DefaultGroupID)
		Rs.Update
	End If
	Rs.Close
	Set Rs = Nothing
	If Status = 0 Then
		set Rs =User_Conn.execute("select GroupID,GroupPoint,GroupMoney,GroupDate From FS_ME_Group where GroupID="&Clng(DefaultGroupID))
		if not Rs.eof then
			if Rs("GroupPoint")>0 then
				User_Conn.execute("Update FS_ME_Users Set Integral=Integral+"& Rs("GroupPoint")&" where UserNumber='"& NoSqlHack(strUserNumberRule) &"'")
			end if
			if Rs("GroupMoney")>0 then
				User_Conn.execute("Update FS_ME_Users Set FS_Money=FS_Money+"& Rs("GroupMoney") &" where UserNumber='"& NoSqlHack(strUserNumberRule) &"'")
			end if
			if Rs("GroupDate")>0 then
				dim DateCoseed
				DateCoseed = dateAdd("d",Rs("GroupDate"),date)
				if G_IS_SQL_User_DB=0 then
					User_Conn.execute("Update FS_ME_Users Set CloseTime=#"& DateCoseed &"# where UserNumber='"& NoSqlHack(strUserNumberRule) &"'")
				else
					User_Conn.execute("Update FS_ME_Users Set CloseTime='"& DateCoseed &"' where UserNumber='"& NoSqlHack(strUserNumberRule) &"'")
				end if
			end if
		end if
		Rs.Close
		Set Rs = Nothing
		Set Rs = Server.CreateObject(G_FS_RS) 
		Rs.open "select  * From FS_ME_MySysPara where 1=0",User_Conn,1,3
		Rs.addnew
		Rs("DownFileRule") = ",,,,"
		Rs("NewsFileRule") = ",,,,"
		Rs("ProductFileRule") = ",,,,"
		Rs("ilogFileRule") = ",,,,"
		Rs("mysiteName") = "我的个人空间"
		Rs("UserNumber") = NoSqlHack(f_strUserNumber)
		Rs("Keywords") = "风讯,CMS,Foosun"
		Rs("Description") = "风讯,CMS,Foosun"
		Rs("NaviPic") = ""
		Rs("isHtml") = 0
		Rs.update
		Rs.close:Set Rs = nothing
		
		Set Rs = Server.CreateObject(G_FS_RS)
		Rs.open "select  * From FS_ME_Log where 1=0",User_Conn,1,3
		Rs.addnew
		Rs("LogType")="注册"
		Rs("UserNumber")=NoSqlHack(f_strUserNumber)
		Rs("points")=NoSqlHack(RegPoint)
		Rs("moneys")=NoSqlHack(RegMoney)
		Rs("LogTime")=Now
		Rs("LogContent")=NoSqlHack(Appid)&"注册获得积分"
		Rs("Logstyle")=0
		Rs.update
		Rs.close
		set Rs = nothing
		Dim FsoObj,Path,UserFpath
		Set FsoObj = Server.CreateObject(G_FS_FSO)  
		Path = strUserNumberRule
		If G_VIRTUAL_ROOT_DIR<>"" Then
			UserFpath = "/"&G_VIRTUAL_ROOT_DIR&"/"&G_USERFILES_DIR
		Else
			UserFpath = "/"&G_USERFILES_DIR
		End If
		if FsoObj.FolderExists(Server.MapPath(UserFpath) ) = false then FsoObj.createFolder Server.MapPath(UserFpath) 
		Path = Server.MapPath(UserFpath&"/"&Path) 
		if FsoObj.FolderExists(Path) = True then FsoObj.deleteFolder Path
		FsoObj.CreateFolder Path
		session("FS_UserName") = NoSqlHack(UserName)
		session("FS_UserNumber") = NoSqlHack(strUserNumberRule)
		session("FS_NickName") = NoSqlHack(UserName)
		session("FS_UserPassword") = NoSqlHack(UserPass)
		session("FS_IsCorp") = 0
		session("FS_UserEmail") = NoSqlHack(UserEmail)
		session("FS_IsLock") = NoSqlHack(RegisterAudit)
		Response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") = 1
		Messenge = "注册成功。"
	End If
End Sub

'更新用户状态
Sub Lockuser()
	Dim UserStatus,Rs,Sql,locktype
	If XmlDoc.documentElement.selectSingleNode("userstatus") is Nothing Then
		Messenge = "<li>参数非法，中止请求。"
		Status = 1
		Exit Sub
	ElseIf Not IsNumeric(XmlDoc.documentElement.selectSingleNode("userstatus").text) Then
		Messenge = "<li>参数非法，中止请求。"
		Status = 1
		Exit Sub
	Else
		UserStatus = Clng(XmlDoc.documentElement.selectSingleNode("userstatus").text)
	End If
	Select Case UserStatus
		Case 1
			locktype="锁定"
		Case 2
			locktype="屏蔽"
		Case Else
			locktype="解锁"
	End Select
	If Not IsObject(User_Conn) Then MF_User_Conn
	Set Rs = Server.CreateObject(G_FS_RS)
	Sql  = "Select isLock From [FS_ME_Users] Where Username='"& NoSqlHack(UserName) &"'"
	Rs.Open Sql,User_Conn,1,3
	If Not Rs.Eof And Not Rs.Bof Then
		Status = 0
		Messenge = "<li>"&locktype&"成功。"
		Rs("isLock") = UserStatus
		Rs.Update
	End If
	Rs.close
	Set Rs = Nothing
End Sub


'用户信息修改
Sub UpdateUser()
	Dim Rs,Sql
	Dim UserPass,UserEmail,Question,Answer
	UserPass = NoSqlHack(XmlDoc.documentElement.selectSingleNode("password").text)
	UserEmail = NoSqlHack(XmlDoc.documentElement.selectSingleNode("email").text)
	Question = NoSqlHack(XmlDoc.documentElement.selectSingleNode("question").text)
	Answer = NoSqlHack(XmlDoc.documentElement.selectSingleNode("answer").text)
	If UserPass<>"" Then
		UserPass = Md5(UserPass,16)
	End If
	If Answer<>"" THen
		Answer = Md5(Answer,16)
	End If
	Set Rs = Server.CreateObject(G_FS_RS)
	Sql="Select Top 1 * From [FS_ME_Users] Where UserName='"&NoSqlHack(UserName)&"'"
	If Not IsObject(User_Conn) Then MF_User_Conn
	Rs.Open Sql,User_Conn,1,3
	If Not Rs.Eof And Not Rs.Bof Then
		If UserPass<>"" Then Rs("UserPassword") = NoSqlHack(UserPass)
		If Answer<>"" Then Rs("PassAnswer") = NoSqlHack(Answer)
		If UserEmail<>"" Then Rs("Email") = NoSqlHack(UserEmail)
		If Question<>"" Then Rs("PassQuestion") = NoSqlHack(Question)
		Rs.update
		Status = 0
		Messenge = "<li>基本资料修改成功。"
	Else
		Status = 1
		Messenge = "<li>该用户不存在，修改资料失败。"
	End If
	Rs.Close
	Set Rs = Nothing
End Sub

'用户退出
Sub LogoutUser()
	Session("FS_UserName") = ""
	Session("FS_UserNumber") = ""
	Session("FS_UserPassword") = ""
	Session("FS_Group") = ""
	Session("FS_IsCorp") = ""
	Session("FS_NickName") = ""
	response.Cookies("FoosunUserCookies")("UserLogin_Style_Num")  = ""
	Session("UserLoginCode") = ""
	Status = 0
	Messenge = "退出成功。"
End Sub


'用户登录
Sub UesrLogin()
	Dim UserPass
	Dim i
	UserPass = XmlDoc.documentElement.selectSingleNode("password").text
	If UserName="" or UserPass="" Then
		Status = 1
		Messenge = Messenge & "<li>请填写用户名或密码。"
		Exit Sub
	End If
	UserPass = Md5(UserPass,16)
	'判断用户是否登录

	dim f_RsLoginobj,f_RsLoginSQL
	f_RsLoginSQL = "Select UserName,UserNumber,NickName,UserPassword,Email,isLock,IsCorporation,GroupID,MySkin,TempLastLoginTime From FS_ME_Users Where UserName = '"& NoSqlHack(UserName) &"' and UserPassword = '"& NoSqlHack(UserPass) &"'"
	
	Set f_RsLoginobj = server.CreateObject(G_FS_RS)
	f_RsLoginobj.open f_RsLoginSQL,User_Conn,1,1
	If Not f_RsLoginobj.eof then 
		If f_RsLoginobj(5)<>0 then
			f_RsLoginobj.close
			set f_RsLoginobj = nothing
			Status = 1
			Messenge = Messenge & "<li>该用户已经被锁定或没有审核.</li>"
		Else
			'更新数据
			Dim f_RsUpdateobj,f_RsUpdateSQL
			Set f_RsUpdateobj = server.CreateObject(G_FS_RS)
			f_RsUpdateSQL = "Select UserNumber,LoginNum,LastLoginTime,FS_Money,Integral,UserLoginCode,TempLastLoginTime,TempLastLoginTime_1,LastLoginIP  From FS_ME_Users Where UserName = '"& NoSqlHack(UserName) &"' and UserPassword = '"& NoSqlHack(UserPass) &"'"
			
			f_RsUpdateobj.open f_RsUpdateSQL,User_Conn,3,3
			f_RsUpdateobj("LoginNum")=f_RsUpdateobj("LoginNum")+1
			f_RsUpdateobj("LastLoginTime")=now
			f_RsUpdateobj("LastLoginIP")=Request.ServerVariables("Remote_Addr")
			Dim f_DateArr,f_DateArryear,f_DateArrmonth,f_DateArrday,f_DateArr_1,f_strmonth,f_strday,f_strhour,f_strminute
			Dim f_Randchars,f_Randchararrs,f_RandLens,f_Randomizecodes,f_iRs
			dim f_strmonth_DateArr_1,f_strday_DateArr_1,f_strhour_DateArr_1,f_strminute_DateArr_1,f_strmonth_DateArr,f_strday_DateArr
			f_DateArr=f_RsUpdateobj("TempLastLoginTime")
			f_DateArr_1=f_RsUpdateobj("TempLastLoginTime_1")
			f_Randomizecodes=GetRamCode(8)
			f_RsUpdateobj("UserLoginCode") = f_Randomizecodes
			'■■■■■■■■■■■■■■■■■■■■■■■■■■■■更新积分和金钱，待考虑
			if p_LoginGetMoney <> 0 And Not IsNull(f_DateArr) then
				if clng(date-dateValue(f_DateArr))>=p_LoginGetMoney then
					f_RsUpdateobj("FS_Money")=f_RsUpdateobj("FS_Money")+p_LoginPointmoneyarr_2
					f_RsUpdateobj("TempLastLoginTime")=now
				End if
			Else
					f_RsUpdateobj("TempLastLoginTime")=now
			End if
			if p_LoginGetIntegral <>0 And Not IsNull(f_DateArr_1) then
				if DateDiff("h",f_DateArr_1,now)>=p_LoginGetIntegral  Or DateDiff("d",now,f_DateArr_1)<>0  then
					f_RsUpdateobj("Integral")=f_RsUpdateobj("Integral")+NoSqlHack(p_LoginPointmoneyarr_1)
					f_RsUpdateobj("TempLastLoginTime_1")=now
				End if 
			Else
					f_RsUpdateobj("TempLastLoginTime_1")=now
			End if
			'■■■■■■■■■■■■■■■■■■■■■■■■■■■■更新积分和金钱，结束
			f_RsUpdateobj.Update  
			f_RsUpdateobj.close:set f_RsUpdateobj=nothing 
		End if
		Status = 0
		Messenge = Messenge & "<li>登录成功.</li>"
	Else
		Status = 1
		Messenge = Messenge & "<li>用户名或密码错误.</li>"
	End if
End Sub
%>