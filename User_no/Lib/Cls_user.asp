<%
Class Cls_User
	Private m_StrEmail,m_NumSex,m_NumIntegral,m_StrQuesion,m_StrAnswer,m_NumLoginNum,m_Paper,m_OnlyLogin,m_Papercode,m_NumConNumber
	Private m_StrRegTime,m_StrLastLoginTime,m_StrLastLoginIP,m_StrUserNumber,m_NumFS_Money,m_PassAnswer
	Private m_StrUserName,m_RealName,m_StrRealName,m_StrNickName,m_StrPWD,m_NumID,m_RsUser,m_UserID,m_StrHomePage,m_StrBothYear
	Private m_StrTel,m_StrMSN,m_StrQQ,m_StrCorner,m_StrProvince,m_StrCity,m_StrAddress,m_StrPostCode,m_isCorporation
	Private m_PassQuestion,m_isOpen,m_OpenInfoTF,m_Vocation,m_NumGroupID,m_UserLoginCode
	Private m_HeadPic,m_SelfIntro,m_UserFavor,m_IsMarray
	Public Mobile,m_CloseTime,isMessage,m_HeadPicsize,safeCode
	Public Property Let Name(ByVal StrValue)
		m_StrUserName = StrValue
		m_RsUser.open "select UserID,UserNumber,UserName,UserPassword,HeadPic,HeadPicSize,PassQuestion,PassAnswer,safeCode,tel,Mobile,isMessage,Email,HomePage,QQ,MSN,Corner,Province,City,Address,PostCode,NickName,RealName,Vocation,Sex,BothYear,Certificate,CertificateCode,IsCorporation,PopList,Integral,FS_Money,RegTime,CloseTime,LoginNum,LastLoginTime,TempLastLoginTime,TempLastLoginTime_1,IsMarray,SelfIntro,isOpen,GroupID,LastLoginIP,ConNumber,ConNumberNews,isLock,UserFavor,MySkin,UserLoginCode,OnlyLogin,hits from FS_ME_Users where UserName='"&NoSqlHack(m_StrUserName)&"'",User_Conn,1,1
		If m_RsUser.EOF=False Then 
			m_RealName = m_RsUser("UserName")
			m_NumGroupID = m_RsUser("GroupID")
			m_NumIntegral = m_RsUser("Integral")
			m_NumLoginNum = m_RsUser("LoginNum")
			m_StrRegTime = m_RsUser("RegTime")
			m_StrLastLoginTime = m_RsUser("LastLoginTime")
			m_StrLastLoginIP = m_RsUser("LastLoginIP")
			m_StrUserNumber = m_RsUser("UserNumber")
			m_NumFS_Money = m_RsUser("FS_Money")
			m_Paper = m_RsUser("Certificate")
			m_Papercode  = m_RsUser("Certificatecode")
			m_OnlyLogin = m_RsUser("OnlyLogin") 
			m_PassAnswer = m_RsUser("PassAnswer")
			m_IsMarray = m_RsUser("IsMarray")
			m_NumConNumber = m_RsUser("ConNumber")
			m_UserID = m_RsUser("UserID")
			m_StrHomePage = m_RsUser("HomePage")
			m_StrBothYear = m_RsUser("BothYear")
			m_StrTel = m_RsUser("Tel")
			m_StrMSN = m_RsUser("MSN")
			m_StrQQ = m_RsUser("QQ")
			m_StrCorner = m_RsUser("Corner")
			m_StrProvince = m_RsUser("Province")
			m_StrCity = m_RsUser("City")
			m_StrAddress = m_RsUser("Address")
			m_StrPostCode = m_RsUser("PostCode")
			m_PassQuestion = m_RsUser("PassQuestion")
			m_SelfIntro = m_RsUser("SelfIntro")
			m_UserFavor = m_RsUser("UserFavor")
			m_isOpen = m_RsUser("isopen")
			m_Vocation = m_RsUser("Vocation")
			m_HeadPic = m_RsUser("HeadPic")
			m_HeadPicsize = m_RsUser("HeadPicsize")
			m_StrRealName = m_RsUser("RealName")
			m_StrNickName = m_RsUser("NickName")
			Mobile = m_RsUser("Mobile")
			m_CloseTime = m_RsUser("CloseTime")
			m_IsCorporation = m_RsUser("IsCorporation")
			isMessage = m_RsUser("isMessage")
			m_StrEmail = m_RsUser("Email")
			m_NumSex = m_RsUser("sex")
			safeCode = m_RsUser("safeCode")
		End If 
		m_RsUser.close
	End Property 

	Public Property Let ID(ByVal StrValue)
		m_NumID = StrValue
		m_RsUser.open "select isLock,UserName,RealName,GroupID,Integral,LoginNum,RegTime, LastLoginTime,LastLoginIP,UserNumber,FS_Money,ConNumber,UserID,HomePage,BothYear,Tel,MSN,QQ,Corner,Province,City,Address,PostCode,PassQuestion,SelfIntro,isOpen,Certificate,CertificateCode,Vocation,HeadPic,NickName,Mobile,CloseTime,IsCorporation,isMessage,Email,sex,safeCode,UserLoginCode,HeadPicsize,OnlyLogin,UserFavor,IsMarray,PassAnswer from FS_ME_Users where ID="&CintStr(m_NumID),User_Conn,1,1
		If m_RsUser.EOF=False Then 
			m_StrUserName = m_RsUser("UserName")
			m_NumIntegral = m_RsUser("Integral")
			m_NumLoginNum = m_RsUser("LoginNum")
			m_StrRegTime = m_RsUser("RegTime")
			m_StrLastLoginTime = m_RsUser("LastLoginTime")
			m_StrLastLoginIP = m_RsUser("LastLoginIP")
			m_StrUserNumber = m_RsUser("UserNumber")
			m_Paper = m_RsUser("Certificate")
			m_Papercode = m_RsUser("Certificatecode")
			m_OnlyLogin = m_RsUser("OnlyLogin")
			m_PassAnswer = m_RsUser("PassAnswer")
			m_IsMarray = m_RsUser("IsMarray")
			m_NumConNumber = m_RsUser("ConNumber")
			m_UserID = m_RsUser("UserID")
			m_NumGroupID =m_RsUser("GroupID")
			m_StrHomePage = m_RsUser("HomePage")
			m_StrBothYear = m_RsUser("BothYear")
			m_StrTel = m_RsUser("Tel")
			m_StrMSN = m_RsUser("MSN")
			m_StrQQ = m_RsUser("QQ")
			m_StrCorner = m_RsUser("Corner")
			m_StrProvince = m_RsUser("Province")
			m_StrCity = m_RsUser("City")
			m_StrAddress = m_RsUser("Address")
			m_StrPostCode = m_RsUser("PostCode")
			m_PassQuestion = m_RsUser("PassQuestion")
			m_SelfIntro = m_RsUser("SelfIntro")
			m_UserFavor = m_RsUser("UserFavor")
			m_isOpen = m_RsUser("isopen")
			m_Vocation = m_RsUser("Vocation")
			m_HeadPic = m_RsUser("HeadPic")
			m_StrNickName = m_RsUser("NickName")
			m_StrRealName = m_RsUser("RealName")
			Mobile = m_RsUser("Mobile")
			CloseTime = m_RsUser("CloseTime")
			m_IsCorporation = m_RsUser("IsCorporation")
			isMessage = m_RsUser("isMessage")
			m_HeadPicsize = m_RsUser("HeadPicsize")
			m_StrEmail = m_RsUser("Email")
			m_NumSex = m_RsUser("sex")
			safeCode = m_RsUser("safeCode")
			m_UserLoginCode =  m_RsUser("UserLoginCode")
		End If 
		m_RsUser.close
	End Property 
	
	Public Property Get UserID()				'用户ID
		UserID = m_UserID
	End Property 
	Public Property Get NumConNumber()				'投稿数量
		NumConNumber = m_NumConNumber
	End Property 

	Public Property Get PaperType()				'证件类型
		PaperType = m_Paper 
	End Property 
	
	Public Property Get PaperTypecode()				'证件号码
		PaperTypecode = m_Papercode  
	End Property 
	
	Public Property Get OnlyLogin ()				'多人登陆? 
		OnlyLogin  = m_OnlyLogin 
	End Property 
	
	Public Property Get PassAnswer ()				'密码答案? 
		PassAnswer  = m_PassAnswer
	End Property 
	
	Public Property Get IsMarray  ()				'是否结婚? 
		IsMarray   = m_IsMarray 
	End Property 
	
	Public Property Get NumFS_Money()				'可用金币
		NumFS_Money = m_NumFS_Money
	End Property 

	Public Property Get isCorp()			'会员类型
		isCorp = m_isCorpOration
	End Property 
	
	Public Property Get NumLoginNum()			'登陆次数
		NumLoginNum = m_NumLoginNum
	End Property 

	Public Property Get UserNumber()			'用户编号
		UserNumber = m_StrUserNumber 
	End Property 
	
	Public Property Get CloseTime()			'用户编号
		CloseTime = m_CloseTime 
	End Property 
	
	Public Property Get NumGroupID()			'用户群权限
		NumGroupID = m_NumGroupID
	End Property 

	Public Property Get RegTime()			'注册时间
		RegTime = m_StrRegTime
	End Property 
	
	Public Property Get LastLoginTime()		'最后登陆时间
		LastLoginTime = m_StrLastLoginTime
	End Property 
	
	Public Property Get LastLoginIP()		'最后登陆IP
		LastLoginIP = m_StrLastLoginIP
	End Property 
	
	Public Property Get NumIntegral()				'积分
		NumIntegral = m_NumIntegral
	End Property 
	
	Public Property Get Sex()				'性别
		Sex = m_NumSex
	End Property 

	Public Property Get Email()				'邮件
		Email = m_StrEmail
	End Property 
	Public Property Get Tel()				'电话
		Tel = m_StrTel
	End Property 
	Public Property Get MSN()				'MSN
		MSN = m_StrMSN
	End Property 
	Public Property Get QQ()				'QQ
		QQ = m_StrQQ
	End Property
	
	Public Property Get Corner()				'地区
		Corner = m_StrCorner
	End Property
	Public Property Get UserLoginCode()				'地区
		UserLoginCode = m_UserLoginCode
	End Property
	Public Property Get Province()				'省份
		Province = m_StrProvince
	End Property 

	Public Property Get City()				'城市
		City = m_StrCity
	End Property 

	Public Property Get Address()				'地址
		Address = m_StrAddress
	End Property 

	Public Property Get PostCode()				'邮编
		PostCode = m_StrPostCode
	End Property 

	Public Property Get HomePage()				'网站地址
		HomePage = m_StrHomePage
	End Property 

	Public Property Get BothYear()				'生日
		BothYear = m_StrBothYear
	End Property 

	Public Property Get PassQuestion()				'密码问题
		PassQuestion = m_PassQuestion
	End Property 

	Public Property Get SelfIntro()				'个性签名
		SelfIntro = m_SelfIntro
	End Property 
	
	Public Property Get UserFavor()				'爱好
		UserFavor = m_UserFavor
	End Property 

	Public Property Get isOpen()				'是否开放资料
		isOpen = m_isOpen
	End Property 

	Public Property Get OpenInfoTF()				'是否开放资料
		OpenInfoTF = m_OpenInfoTF
	End Property 

	Public Property Get Vocation()				'职业
		Vocation = m_Vocation
	End Property 

	Public Property Get HeadPic()				'头像
		HeadPic = m_HeadPic
	End Property 
		
	Public Property Get HeadPicsize()				'头像
		HeadPicsize = m_HeadPicsize
	End Property 	
	
	Public Property Get UserName()				'用户名
			UserName = m_StrUserName
	End Property 
	
	Public Property Get RealName() '真实姓名
		RealName = m_StrRealName
	End ProPerty
	
	Public Property Get EName()				'英文名字
		EName = m_StrUserName
	End Property 
	
	Public Property Get NickName()				'昵称
		NickName = m_StrNickName
	End Property 

	Private Sub Class_Initialize()
		Set m_RsUser = server.CreateObject(G_FS_RS)
	End Sub

	Private Sub Class_Terminate()
		Set m_RsUser = Nothing 
	End Sub
	Public Function UserGroups(f_strfield,f_strvalue)
		Dim f_RsUG,f_StrUG,f_StrSelected
		UserGroups = ""
		f_StrSelected = ""
		Set f_RsUG = User_Conn.Execute("Select "&f_strfield&",Name from FS_MemGroup")
		Do While Not f_RsUG.EOF 
			f_StrUG = f_RsUG(0)
			If f_StrUG = f_strvalue Then f_StrSelected = "selected"
			UserGroups = UserGroups & "<option value="""&f_StrUG&""" "&f_StrSelected&">"&f_RsUG(1)&"</option>" & vbcrlf
			f_StrSelected = ""
			f_RsUG.MoveNext
		Loop 
		Set f_RsUG = Nothing 
	End Function 

	Public Function DelUser(f_StrNumName,f_StrPWD)
		f_StrNumName = NoSqlHack(f_StrNumName)
		DelUser = True
		Dim StrNumName
		If f_StrNumName="" Or f_StrPWD="" Or IsNull(f_StrNumName) Or IsNull(f_StrPWD) Then
			DelUser = False
		Else
			Dim f_RsMemObj
			Set f_RsMemObj = User_Conn.Execute("select UserNumber from FS_ME_Users where (UserNumber='"&f_StrNumName&"' or UserName='"&f_StrNumName&"') and UserPassword='"&f_StrPWD&"'")
			If f_RsMemObj.EOF Then 
				DelUser = False
				Set f_RsMemObj = Nothing 
			Else
				On Error Resume Next
				StrNumName = f_RsMemObj(0)
				'User_Conn.Execute("Delete from FS_ME_Message where M_ReadUserNumber='"&StrNumName&"'")
				StrNumName = NoSqlHack(StrNumName)
				User_Conn.Execute("Update FS_ME_Message  set M_FromUserNumber=0 where M_FromUserNumber='"&StrNumName&"'")
				User_Conn.Execute("Update FS_ME_Message  set M_ReadUserNumber=0 where M_ReadUserNumber='"&StrNumName&"'")
				User_Conn.Execute("Delete from FS_ME_BuyBag where UserNumber='"&StrNumName&"'")
				User_Conn.Execute("Update FS_ME_Card set UserNumber=0 where UserNumber='"&StrNumName&"'")
				User_Conn.Execute("Delete from FS_ME_CertFile where UserNumber='"&StrNumName&"'")
				User_Conn.Execute("Delete from FS_ME_CorpUser where UserNumber='"&StrNumName&"'")
				User_Conn.Execute("Delete from FS_ME_Favorite where UserNumber='"&StrNumName&"'")
				User_Conn.Execute("Delete from FS_ME_FavoriteClass where UserNumber='"&StrNumName&"'")
				User_Conn.Execute("Delete from FS_ME_Friends where UserNumber='"&StrNumName&"'")
				User_Conn.Execute("Update FS_ME_Friends set F_UserNumber = 0  where F_UserNumber='"&StrNumName&"'")
				User_Conn.Execute("Delete from FS_ME_GroupDebate where UserNumber='"&StrNumName&"'")
				User_Conn.Execute("Delete from FS_ME_GroupDebateClass where UserNumber='"&StrNumName&"'")
				User_Conn.Execute("Delete from FS_ME_InfoClass where UserNumber='"&StrNumName&"'")
				User_Conn.Execute("Delete from FS_ME_InfoContribution where UserNumber='"&StrNumName&"'")
				User_Conn.Execute("Delete from FS_ME_InfoDown where UserNumber='"&StrNumName&"'")
				User_Conn.Execute("Delete from FS_ME_Infoilog where UserNumber='"&StrNumName&"'")
				User_Conn.Execute("Delete from FS_ME_InfoProduct where UserNumber='"&StrNumName&"'")
				User_Conn.Execute("Delete from FS_ME_Log where UserNumber='"&StrNumName&"'")
				User_Conn.Execute("Delete from FS_ME_MyInfo where UserNumber='"&StrNumName&"'")
				User_Conn.Execute("Delete from FS_ME_MySysPara where UserNumber='"&StrNumName&"'")
				User_Conn.Execute("Delete from FS_ME_Order where UserNumber='"&StrNumName&"'")
				User_Conn.Execute("Delete from FS_ME_Review where UserNumber='"&StrNumName&"'")
				User_Conn.Execute("Delete from FS_SD_InfoSupply where UserNumber='"&StrNumName&"'")
				'删除静态目录、其他目录
				'暂时保留 
				Set f_RsMemObj = Nothing 
				User_Conn.Execute("Delete from FS_ME_Users where UserNumber='"&StrNumName&"'")
				If Err Then 
					Err.clear
					DelUser = False
				Else
					DelUser = True 
				End If 
			End If 
		End If 
	End Function 
	
	Public Function strUserNumberRule(str)
		strUserNumberRule = ""
		Dim f_strUserNumberarr,f_str0,f_str1,f_str2,f_str3,f_str4,Getstr
		if instr(str,",")=0 then strUserNumberRule=str : Exit Function
		f_strUserNumberarr = split(str,",")
		If Not IsArray(f_strUserNumberarr) then Exit Function
		f_str0 = f_strUserNumberarr(0)
		f_str1 = f_strUserNumberarr(1)
		f_str2 = f_strUserNumberarr(2)
		f_str3 = f_strUserNumberarr(3)
		f_str4 = f_strUserNumberarr(4)
		strUserNumberRule = strUserNumberRule & f_strUserNumberarr(0)
		If Instr(1,f_strUserNumberarr(1),"y",1)<>0 then
			if Len(Trim(Cstr(f_strUserNumberarr(3))))<>0 then
				strUserNumberRule = strUserNumberRule & right(year(now),2)&f_strUserNumberarr(3)
			Else
				strUserNumberRule = strUserNumberRule & right(year(now),2)
			End if
		End if
		If Instr(1,f_strUserNumberarr(1),"m",1)<>0 then
				If Len(Cstr(Month(Now()))) < 2 then
						if Len(Trim(Cstr(f_strUserNumberarr(3))))<>0 then
							strUserNumberRule = strUserNumberRule & "0"&month(now)&f_strUserNumberarr(3)
						Else
							strUserNumberRule = strUserNumberRule& "0"&month(now)&f_strUserNumberarr(3)
						End if
				Else
						if Len(Trim(Cstr(f_strUserNumberarr(3))))<>0 then
							strUserNumberRule = strUserNumberRule & month(now)&f_strUserNumberarr(3)
						Else
							strUserNumberRule = strUserNumberRule& month(now)&f_strUserNumberarr(3)
						End if
				End if
		End if
		If Instr(1,f_strUserNumberarr(1),"d",1)<>0 then
				If Len(Cstr(day(Now))) < 2 then
						if Len(Trim(Cstr(f_strUserNumberarr(3))))<>0 then
							strUserNumberRule = strUserNumberRule & "0"&day(now)&f_strUserNumberarr(3)
						Else
							strUserNumberRule = strUserNumberRule& "0"&day(now)&f_strUserNumberarr(3)
						End if
				Else
						if Len(Trim(Cstr(f_strUserNumberarr(3))))<>0 then
							strUserNumberRule = strUserNumberRule & day(now)&f_strUserNumberarr(3)
						Else
							strUserNumberRule = strUserNumberRule& day(now)&f_strUserNumberarr(3)
						End if
				End if
		End if
		If Instr(1,f_strUserNumberarr(1),"h",1)<>0 then
				If Len(Cstr(hour(Now))) < 2 then
						if Len(Trim(Cstr(f_strUserNumberarr(3))))<>0 then
							strUserNumberRule = strUserNumberRule & "0"&hour(now)&f_strUserNumberarr(3)
						Else
							strUserNumberRule = strUserNumberRule& "0"&hour(now)&f_strUserNumberarr(3)
						End if
				Else
						if Len(Trim(Cstr(f_strUserNumberarr(3))))<>0 then
							strUserNumberRule = strUserNumberRule & hour(now)&f_strUserNumberarr(3)
						Else
							strUserNumberRule = strUserNumberRule& hour(now)&f_strUserNumberarr(3)
						End if
				End if
		End if
		If Instr(1,f_strUserNumberarr(1),"i",1)<>0 then
				If Len(Cstr(minute(Now))) < 2 then
						if Len(Trim(Cstr(f_strUserNumberarr(3))))<>0 then
							strUserNumberRule = strUserNumberRule & "0"&minute(now)&f_strUserNumberarr(3)
						Else
							strUserNumberRule = strUserNumberRule& "0"&minute(now)&f_strUserNumberarr(3)
						End if
				Else
						if Len(Trim(Cstr(f_strUserNumberarr(3))))<>0 then
							strUserNumberRule = strUserNumberRule & minute(now)&f_strUserNumberarr(3)
						Else
							strUserNumberRule = strUserNumberRule& minute(now)&f_strUserNumberarr(3)
						End if
				End if
		End if
		If Instr(1,f_strUserNumberarr(1),"s",1)<>0 then
				If Len(Cstr(second(Now))) < 2 then
						if Len(Trim(Cstr(f_strUserNumberarr(3))))<>0 then
							strUserNumberRule = strUserNumberRule & "0"&second(now)&f_strUserNumberarr(3)
						Else
							strUserNumberRule = strUserNumberRule& "0"&second(now)&f_strUserNumberarr(3)
						End if
				Else
						if Len(Trim(Cstr(f_strUserNumberarr(3))))<>0 then
							strUserNumberRule = strUserNumberRule & second(now)&f_strUserNumberarr(3)
						Else
							strUserNumberRule = strUserNumberRule& second(now)&f_strUserNumberarr(3)
						End if
				End if
		End if
		Randomize
		Dim f_Randchar,f_Randchararr,f_RandLen,f_iR,f_Randomizecode
		f_Randchar="0,1,2,3,4,5,6,7,8,9,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z"
		f_Randchararr=split(f_Randchar,",") 
		If f_strUserNumberarr(2)="2" then
			if f_strUserNumberarr(4)="w" then
				f_RandLen=2 
				for f_iR=1 to f_RandLen
				f_Randomizecode=f_Randomizecode&f_Randchararr(Int((21*Rnd)))
				next 
				strUserNumberRule = strUserNumberRule &  f_Randomizecode
			Else
				strUserNumberRule = strUserNumberRule &  CStr(Int((99 * Rnd) + 1))
			End if
		Elseif f_strUserNumberarr(2)="3" then
			if f_strUserNumberarr(4)="w" then
				f_RandLen=3 
				for f_iR=1 to f_RandLen
				f_Randomizecode=f_Randomizecode&f_Randchararr(Int((21*Rnd)))
				next 
				strUserNumberRule = strUserNumberRule &  f_Randomizecode
			Else
				strUserNumberRule = strUserNumberRule &  CStr(Int((999* Rnd) + 1))
			End if
		Elseif f_strUserNumberarr(2)="4" then
			if f_strUserNumberarr(4)="w" then
				f_RandLen=4 
				for f_iR=1 to f_RandLen
				f_Randomizecode=f_Randomizecode&f_Randchararr(Int((21*Rnd)))
				next 
				strUserNumberRule = strUserNumberRule &  f_Randomizecode
			Else
				strUserNumberRule = strUserNumberRule &  CStr(Int((9999* Rnd) + 1))
			End if
		Elseif f_strUserNumberarr(2)="5" then
			if f_strUserNumberarr(4)="w" then
				f_RandLen=5 
				for f_iR=1 to f_RandLen
				f_Randomizecode=f_Randomizecode&f_Randchararr(Int((21*Rnd)))
				next 
				strUserNumberRule = strUserNumberRule &  f_Randomizecode
			Else
				strUserNumberRule = strUserNumberRule &  CStr(Int((99999* Rnd) + 1))
			End if
		End if
		strUserNumberRule = strUserNumberRule
	End Function
		
	Public Function Login(f_StrName,f_StrPWD,f_Logintye,p_vercode)
		Login = True
		'Response.Cookies("FoosunMFCookies")("FoosunMFDomain")
		if request.Cookies("FoosunUserlCookies")("FS_User_Login_Number")="" or request.Cookies("FoosunUserlCookies")("FS_User_Login_Number")=0 then
			p_LoginLockNum = 0
		Else
			p_LoginLockNum = cint(request.Cookies("FoosunUserlCookies")("FS_User_Login_Number"))
		End if
		if LoginLockNum<>0 then
			if p_LoginLockNum > LoginLockNum then
					Login = false 
					strShowErr = "<li>您已经连续登陆了"& p_LoginLockNum -1 &"次</li><li> 此帐户已经临时被锁定，今天不能登陆了!!!</li>"
					Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
					Response.end
			End if
		End if
		if p_UserName = "" or  p_UserPassword = "" then
			Login = false 
			strShowErr = "<li>请填写您的用户名</li><li> 请填写您的密码!</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		Else
			dim f_RsLoginobj,f_RsLoginSQL
			if f_Logintye = "0" then
				 f_RsLoginSQL = "Select UserName,UserNumber,NickName,UserPassword,Email,isLock,IsCorporation,GroupID,MySkin,TempLastLoginTime From FS_ME_Users Where UserName = '"& NoSqlHack(f_StrName) &"' and UserPassword = '"& NoSqlHack(f_StrPWD) &"'"
			 Elseif f_Logintye = "1" then
				 f_RsLoginSQL = "Select UserName,UserNumber,NickName,UserPassword,Email,isLock,IsCorporation,GroupID,MySkin,TempLastLoginTime From FS_ME_Users Where UserNumber = '"& NoSqlHack(f_StrName) &"' and UserPassword = '"& NoSqlHack(f_StrPWD) &"'"
			 Elseif f_Logintye = "2" then
				 f_RsLoginSQL = "Select UserName,UserNumber,NickName,UserPassword,Email,isLock,IsCorporation,GroupID,MySkin,TempLastLoginTime From FS_ME_Users Where Email = '"& NoSqlHack(f_StrName) &"' and UserPassword = '"& NoSqlHack(f_StrPWD) &"'"
			 Else
				Login = false 
				strShowErr = "<li>错误的参数</li><li> 请选择登陆方式!</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			 End if
		'判断用户是否处于过期的群组Crazy
		Dim GroupExpired,Groupsql
				Groupsql = "Select FS_ME_Users.Regtime,FS_ME_Users.GroupID,FS_ME_Users.Email,FS_ME_Users.UserNumber,FS_ME_Group.GroupID,FS_ME_Group.GroupDate,FS_ME_Users.UserName,FS_ME_Users.UserPassword From FS_ME_Users,FS_ME_Group Where FS_ME_Users.GroupID = FS_ME_Group.GroupID And "
		Select Case f_Logintye
			Case "0"
					Groupsql = Groupsql&"FS_ME_Users.UserName='"& NoSqlHack(f_StrName) &"' And FS_ME_Users.UserPassword='"& NoSqlHack(f_StrPWD) &"'"	
			Case "1"
					Groupsql = Groupsql&"FS_ME_Users.UserNumber='"& NoSqlHack(f_StrName) &"' And FS_ME_Users.UserPassword='"& NoSqlHack(f_StrPWD) &"'"	
			Case"2"
					Groupsql = Groupsql&"FS_ME_Users.Email='"& NoSqlHack(f_StrName) &"' And FS_ME_Users.UserPassword='"& NoSqlHack(f_StrPWD) &"'"	
		End Select
		set GroupExpired = server.CreateObject(G_FS_RS)
		GroupExpired.open Groupsql,User_Conn,1,1
		if not GroupExpired.eof Then
			If GroupExpired("GroupDate") <> "0" Then
				if dateadd("d",GroupExpired("Regtime"),GroupExpired("GroupDate")) < Now() Then
					GroupExpired.close
					Set GroupExpired = nothing
					strShowErr = "<li>用户已过期</li>"
					Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
					Response.end
			End if
			End if
		End if
		GroupExpired.close
		Set GroupExpired = nothing
		'结束
			Set f_RsLoginobj = server.CreateObject(G_FS_RS)
			f_RsLoginobj.open f_RsLoginSQL,User_Conn,1,1
			If Not f_RsLoginobj.eof then 
				If f_RsLoginobj(5)<>0 then
					f_RsLoginobj.close
					set f_RsLoginobj = nothing
					Login = false 
					strShowErr = "<li>用户已经被锁定</li><li> 此用户注册没有审核!</li>"
					Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
					Response.end
				Else
					'更新数据
					Dim f_RsUpdateobj,f_RsUpdateSQL
					Set f_RsUpdateobj = server.CreateObject(G_FS_RS)
						if f_Logintye = "0" then
							 f_RsUpdateSQL = "Select UserNumber,LoginNum,LastLoginTime,FS_Money,Integral,UserLoginCode,TempLastLoginTime,TempLastLoginTime_1,LastLoginIP,GroupID  From FS_ME_Users Where UserName = '"& NoSqlHack(f_StrName) &"' and UserPassword = '"& NoSqlHack(f_StrPWD) &"'"
						 Elseif f_Logintye = "1" then
							 f_RsUpdateSQL = "Select UserNumber,LoginNum,LastLoginTime,FS_Money,Integral,UserLoginCode,TempLastLoginTime,TempLastLoginTime_1,LastLoginIP,GroupID From FS_ME_Users Where UserNumber = '"& NoSqlHack(f_StrName) &"' and UserPassword = '"& NoSqlHack(f_StrPWD) &"'"
						 Elseif f_Logintye = "2" then
							 f_RsUpdateSQL = "Select UserNumber,LoginNum,LastLoginTime,FS_Money,Integral,UserLoginCode,TempLastLoginTime,TempLastLoginTime_1,LastLoginIP,GroupID From FS_ME_Users Where Email = '"& NoSqlHack(f_StrName) &"' and UserPassword = '"& NoSqlHack(f_StrPWD) &"'"
						 End if
					f_RsUpdateobj.open  f_RsUpdateSQL,User_Conn,3,3
					f_RsUpdateobj("LoginNum")=f_RsUpdateobj("LoginNum")+1
					f_RsUpdateobj("LastLoginTime")=now
					if isnull(f_RsUpdateobj("GroupID")) then
						f_RsUpdateobj("GroupID")=NoSqlHack(DefaultGroupID)
					end if
					f_RsUpdateobj("LastLoginIP")=NoSqlHack(Request.ServerVariables("Remote_Addr"))
					Dim f_DateArr,f_DateArryear,f_DateArrmonth,f_DateArrday,f_DateArr_1,f_strmonth,f_strday,f_strhour,f_strminute
					Dim f_Randchars,f_Randchararrs,f_RandLens,f_Randomizecodes,f_iRs
					dim f_strmonth_DateArr_1,f_strday_DateArr_1,f_strhour_DateArr_1,f_strminute_DateArr_1,f_strmonth_DateArr,f_strday_DateArr
					f_DateArr=f_RsUpdateobj("TempLastLoginTime")
					f_DateArr_1=f_RsUpdateobj("TempLastLoginTime_1")
					f_Randomizecodes=GetRamCode(8)
					f_RsUpdateobj("UserLoginCode") = f_Randomizecodes
					if p_LoginGetMoney <> 0 Then
						If  Not IsNull(f_DateArr) then
							if clng(date-dateValue(f_DateArr))>=p_LoginGetMoney then
								f_RsUpdateobj("FS_Money")=f_RsUpdateobj("FS_Money")+p_LoginPointmoneyarr_2
								f_RsUpdateobj("TempLastLoginTime")=now
							End If
						End if
					Else
							f_RsUpdateobj("TempLastLoginTime")=now
					End if
					if p_LoginGetIntegral <>0 then
						if DateDiff("h",f_DateArr_1,now)>=p_LoginGetIntegral  Or DateDiff("d",now,f_DateArr_1)<>0  then
							f_RsUpdateobj("Integral")=f_RsUpdateobj("Integral")+p_LoginPointmoneyarr_1
							f_RsUpdateobj("TempLastLoginTime_1")=now
						End if 
					Else
							f_RsUpdateobj("TempLastLoginTime_1")=now
					End if
					f_RsUpdateobj.Update  
					f_RsUpdateobj.close:set f_RsUpdateobj=nothing 
					session("FS_UserName") = f_RsLoginobj(0)
					session("FS_UserNumber") = f_RsLoginobj(1)
					session("FS_UserPassword") = f_RsLoginobj(3)
					session("FS_UserEmail")  = f_RsLoginobj(4)'改为Cookies
					If Not IsNull(f_RsLoginobj(8)) then
						Response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") = f_RsLoginobj(8)'改为Cookies
					Else
						Response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") =1'改为Cookies
					End if
					session("UserLoginCode") = f_Randomizecodes'改为Cookies
					f_RsLoginobj.close:	set f_RsLoginobj = nothing
					If CBool(Request.Form("AutoGet")) or Request.Form("AutoGet")<>"" Then
						Response.Cookies("FoosunUserCookie")("FS_UserName")=Session("FS_UserName")
						Response.Cookies("FoosunUserCookie").Expires=Date()+365
					Else
						Response.Cookies("FoosunUserCookie")("FS_UserName")=""
						Response.Cookies("FoosunUserCookie").Expires=Date()-1
					End If
					Response.Cookies("FoosunUserlCookies")("FS_User_Login_Number")=0
					Login = True 
				End if
			Else
					Response.Cookies("FoosunUserlCookies")("FS_User_Login_Number")=p_LoginLockNum+1
					Response.Cookies("FoosunUserlCookies").Expires =Date()+1
					Login = false 
			End if
		End if 
	End Function
	
	Public Function checkName(f_StrName)
		Dim CheckNameTF,CheckArr,CheckNum,InstrNum,Str_Limit
		If Instr(f_StrName,".") > 0 Then
			checkName = False
			Exit Function
		End If	
		m_RsUser.open "select UserName from FS_ME_Users where UserName ='"&NoSqlHack(f_StrName)&"'",User_Conn,1,1
		If m_RsUser.EOF Then 
			Set CheckNameTF = User_Conn.ExeCute("Select Top 1 LimitUserName From FS_ME_SysPara Where SysID > 0 Order By SysID")
			If CheckNameTF.Eof Then
				checkName = True
			Else
				Str_Limit = CheckNameTF("LimitUserName")
				If Str_Limit = "" Or Isnull(Str_Limit) Then
					checkName = True
				Else
					If Instr(Str_Limit,",") > 0 Then
						CheckArr = Split(Str_Limit,",")
						checkName = True
						For CheckNum = LBound(CheckArr) To UBound(CheckArr)
							If CheckArr(CheckNum) <> "" Then
								If f_StrName = CheckArr(CheckNum) Or Instr(f_StrName,CheckArr(CheckNum)) > 0 Then
									checkName = False
									Exit For
								End If
							End If
						Next
					Else
						If f_StrName <> Str_Limit And Instr(f_StrName,Str_Limit) = 0 Then
							checkName = True
						Else
							checkName = False
						End if	
					End If		
				End If	
			End If
			CheckNameTF.Close : Set CheckNameTF = Nothing	
		Else
			checkName = False
		End If
		m_RsUser.close
	End Function 

	Public Function checkEmail(f_StrEmail)
		Dim CheckRs
		m_RsUser.open "select UserID from FS_Me_Users where Email ='"&NoSqlHack(f_StrEmail)&"'",User_Conn,1,1
		If m_RsUser.EOF Then 
			Set CheckRs = Server.CreateObject (G_FS_RS)
			checkEmail = True
		Else
			checkEmail = False
		End If
		m_RsUser.close
	End Function 

   Public Function chk_regname(regname,strUserName)
        Dim regbadstr, i
        regbadstr = Split(regname, ",")
        chk_regname = True
        For i = 0 To UBound(regbadstr)
            If Trim(regbadstr(i)) <> "" Then
                If Trim(strUserName) = Trim(regbadstr(i)) Then
                    chk_regname = false
                End If
            End If
            If chk_regname = false Then Exit For
        Next
    End Function
	
	Public Function UserExist(f_StrName)
		UserExist = True
		If f_StrName="" Or IsNull(f_StrName) Then
			UserExist = False
		Else
			Dim f_RsMemObj
			set f_RsMemObj = Server.CreateObject (G_FS_RS)
			f_RsMemObj.Open "select isLock from FS_ME_Users where UserNumber='"& NoSqlHack(f_StrName) &"'",User_Conn,1,1
			if not f_RsMemObj.EOF then
				if cint(f_RsMemObj("isLock"))=1 Then UserExist = False
			else
				UserExist = False
			End If
			f_RsMemObj.Close
			set f_RsMemObj = Nothing
		End If
	End Function
    	
	Public Function checkStat(f_StrName,f_StrPWD)
		checkStat = True
		If f_StrName="" Or Len(f_StrPWD)<>16 Or IsNull(f_StrName) Or IsNull(f_StrPWD) Then
			checkStat = False
		Else
			Dim f_RsUserObj
			set f_RsUserObj = Server.CreateObject (G_FS_RS)
			f_RsUserObj.Open "select isLock,UserName,RealName,GroupID,Integral,LoginNum,RegTime, LastLoginTime,LastLoginIP,UserNumber,FS_Money,ConNumber,UserID,HomePage,BothYear,Tel,MSN,QQ,Corner,Province,City,Address,PostCode,PassQuestion,SelfIntro,isOpen,Certificate,CertificateCode,Vocation,HeadPic,NickName,Mobile,CloseTime,IsCorporation,isMessage,Email,sex,safeCode,UserLoginCode,HeadPicsize,OnlyLogin,UserFavor,IsMarray,PassAnswer from FS_ME_Users where UserName='"& NoSqlHack(f_StrName) &"' and UserPassword='"& NoSqlHack(f_StrPWD) &"'",User_Conn,1,1
			if not f_RsUserObj.EOF then
				if cint(f_RsUserObj("isLock"))<>0 Then
					checkStat = False
				Else
					m_StrUserName = f_RsUserObj("UserName")
					m_StrRealName = f_RsUserObj("RealName")
					m_NumGroupID = f_RsUserObj("GroupID")
					m_NumIntegral = f_RsUserObj("Integral")
					m_NumLoginNum = f_RsUserObj("LoginNum")
					m_StrRegTime = f_RsUserObj("RegTime")
					m_Paper =  f_RsUserObj("Certificate")
					m_Papercode =  f_RsUserObj("Certificatecode")
					m_OnlyLogin = f_RsUserObj("OnlyLogin")
					m_PassAnswer = f_RsUserObj("PassAnswer")
					m_IsMarray  = f_RsUserObj("IsMarray")
					m_StrLastLoginTime = f_RsUserObj("LastLoginTime")
					m_StrLastLoginIP = f_RsUserObj("LastLoginIP")
					m_StrUserNumber = f_RsUserObj("UserNumber")
					m_NumFS_Money = f_RsUserObj("FS_Money")
					m_NumConNumber = f_RsUserObj("ConNumber")
					m_UserID = f_RsUserObj("UserID")
					m_StrHomePage = f_RsUserObj("HomePage")
					m_StrBothYear = f_RsUserObj("BothYear")
					m_StrTel = f_RsUserObj("Tel")
					m_StrMSN = f_RsUserObj("MSN")
					m_StrQQ = f_RsUserObj("QQ")
					m_StrCorner = f_RsUserObj("Corner")
					m_StrProvince = f_RsUserObj("Province")
					m_StrCity = f_RsUserObj("City")
					m_StrAddress = f_RsUserObj("Address")
					m_StrPostCode = f_RsUserObj("PostCode")
					m_PassQuestion = f_RsUserObj("PassQuestion")
					m_SelfIntro = f_RsUserObj("SelfIntro")
					m_UserFavor =  f_RsUserObj("UserFavor")
					m_isOpen = f_RsUserObj("isOpen")
					m_Vocation = f_RsUserObj("Vocation")
					m_HeadPic = f_RsUserObj("HeadPic")
					m_HeadPicsize = f_RsUserObj("HeadPicsize")
					m_StrNickName = f_RsUserObj("NickName")
					Mobile = f_RsUserObj("Mobile")
					m_CloseTime = f_RsUserObj("CloseTime")
					m_IsCorporation = f_RsUserObj("IsCorporation")
					isMessage = f_RsUserObj("isMessage")
					m_StrEmail = f_RsUserObj("Email")
					m_NumSex = f_RsUserObj("sex")
					safeCode = f_RsUserObj("safeCode")
					m_UserLoginCode =  f_RsUserObj("UserLoginCode")
				end if
			Else
				checkStat = False
			End If
			f_RsUserObj.Close:set f_RsUserObj = Nothing
		End If
	End Function
	
	Public Function CheckPostinput()
		On Error Resume Next
		Dim server_v1, server_v2
		CheckPost = False
		server_v1 = NoSqlHack(CStr(Request.ServerVariables("HTTP_REFERER")))
		server_v2 = NoSqlHack(CStr(Request.ServerVariables("SERVER_NAME")))
		If Mid(server_v1, 8, Len(server_v2)) = server_v2 Then
			CheckPost = True
		End If
	End Function

	Public Sub out()
		Session("FS_UserName") = ""
		Session("FS_UserNumber") = ""
		Session("FS_UserPassword") = ""
		Session("FS_Group") = ""
		Session("FS_IsCorp") = ""
		Session("FS_NickName") = ""
		response.Cookies("FoosunUserCookies")("UserLogin_Style_Num")  = ""
		Session("UserLoginCode") = ""
	End Sub

	Public Function ChangePWD(f_StrName,StrOldPWD,StrNewPWD)
		If f_StrName="" Or StrOldPWD="" Then
			ChangePWD = "帐号或密码不正确"
		Else
			Dim ObjPWD
			Set ObjPWD = Server.CreateObject(G_FS_RS)
			objPWD.open "select Password from FS_Members where MemName='"&NoSqlHack(f_StrName)&"' and Password='"&NoSqlHack(StrOldPWD)&"'",User_Conn,3,3
			If Not ObjPWD.EOF Then
				ObjPWD("Password")=StrNewPWD
				ObjPWD.update
				Response.Cookies("Foosun")("MemPassword") = StrNewPWD
				ChangePWD = True
			Else
				ChangePWD = "您不是风讯会员"
			End If		
		End If
	End Function

	Public Function FriendList()
		FriendList = ""
		Dim f_RsFriend,f_StrFriend
		Set f_RsFriend = User_Conn.Execute("Select  top 50 F_UserNumber from FS_ME_Friends where FriendType =0 and UserNumber='"& session("FS_UserNumber") &"' order by FriendID desc")
		Do While Not f_RsFriend.EOF 
			if f_RsFriend("F_UserNumber")= "0" then
					f_RsFriend.MoveNext
			Else
				f_StrFriend = f_RsFriend(0)
				Dim f_GetUserClsObj ,f_strGetCls,f_StrTmpFriend,f_StrUserNamechar
				'Call UserExist(f_StrFriend)
				Set f_GetUserClsObj = User_Conn.execute("select UserNumber,RealName,UserName from FS_ME_Users where UserNumber ='"& f_RsFriend("F_UserNumber") &"'")
				if Not f_GetUserClsObj.eof then
					if f_GetUserClsObj("RealName") = "" then
						f_strGetCls = f_GetUserClsObj("UserName")
					Else
						f_strGetCls = f_GetUserClsObj("RealName")
					End if
					f_StrTmpFriend = f_GetUserClsObj("UserName")
					f_StrUserNamechar = "("&f_GetUserClsObj("UserName")&")"
			    Else
					f_RsFriend.MoveNext
			    End if
				FriendList = FriendList & "<option value="""&f_StrTmpFriend&""">·"&f_strGetCls & f_StrUserNamechar&"</option>" & vbcrlf
				f_RsFriend.MoveNext
			End if
		Loop 
		set f_GetUserClsObj = nothing
		Set f_RsFriend = Nothing 
	End Function

	Public Function AddFriend(f_FriendName,f_FriendCName,f_SelfName,f_type)
		Dim f_RsFriend
		Set f_RsFriend = Server.CreateObject(G_FS_RS)
		f_RsFriend.Open "select * from FS_Friend where FriendName='"&NoSqlHack(f_FriendName)&"'",User_Conn,1,3
		If f_RsFriend.EOF = False Then 
			AddFriend = False
		Else
			f_RsFriend.addNew
			f_RsFriend("FriendName")=f_FriendName
			f_RsFriend("RealName")=f_FriendCName
			f_RsFriend("MemName")=f_SelfName
			f_RsFriend("type")=f_type
			f_RsFriend.Update
			AddFriend = True 
		End If 
		Set f_RsFriend = Nothing 
	End Function
	
	Public Function InsertMyPara(f_strUserNumber)
			Dim f_Rsmypara
			Set f_Rsmypara = server.CreateObject(G_FS_RS) 
			f_Rsmypara.open "select  * From FS_ME_MySysPara where 1=0",User_Conn,1,3
			f_Rsmypara.addnew
			f_Rsmypara("DownFileRule") = ",,,,"
			f_Rsmypara("NewsFileRule") = ",,,,"
			f_Rsmypara("ProductFileRule") = ",,,,"
			f_Rsmypara("ilogFileRule") = ",,,,"
			f_Rsmypara("mysiteName") = "我的个人空间"
			f_Rsmypara("UserNumber") = f_strUserNumber
			f_Rsmypara("Keywords") = "风讯,CMS,Foosun"
			f_Rsmypara("Description") = "风讯,CMS,Foosun"
			f_Rsmypara("NaviPic") = ""
			f_Rsmypara("isHtml") = 0
			'f_Rsmypara("RedirectUrl") = ""
			f_Rsmypara.update
			f_Rsmypara.close:Set f_Rsmypara = nothing
	End Function
	
	Public Function DelFriend(f_NumID)
		On Error Resume Next
		User_Conn.Execute("Delete From FS_Friend Where id in("& FormatIntArr(f_NumID) & ")")
		If Err Then 
			Err.clear
			DelFriend = False
		Else
			DelFriend = True
		End If 
	End Function 
	
	Public Function GetFriendNumber(f_strNumber)
		Dim RsGetFriendNumber
		Set RsGetFriendNumber = User_Conn.Execute("Select UserNumber From FS_ME_Users Where UserName = '"& NoSqlHack(f_strNumber) &"'")
		If  Not RsGetFriendNumber.eof  Then 
			GetFriendNumber = RsGetFriendNumber("UserNumber")
		Else
			GetFriendNumber = ""
		End If 
		set RsGetFriendNumber = nothing
	End Function 
	
	Public Function GetFriendName(f_strNumber)
		if f_strNumber="0" then
				GetFriendName = "管理员"
		else
			Dim RsGetFriendName
			Set RsGetFriendName = User_Conn.Execute("Select UserName From FS_ME_Users Where UserNumber = '"& NoSqlHack(f_strNumber) &"'")
			If  Not RsGetFriendName.eof  Then 
				GetFriendName = RsGetFriendName("UserName")
			Else
				GetFriendName = "用户已经被删除"
			End If 
			set RsGetFriendName = nothing
		end if
	End Function 
	
	Public Function ChangeFriend(f_NumID,f_Type)
		On Error Resume Next
		User_Conn.Execute("update FS_Friend set type="&f_Type&" Where id in("& FormatIntArr(f_NumID)&")")
		If Err Then 
			Err.clear
			ChangeFriend = False
		Else
			ChangeFriend = True
		End If 
	End Function 
	
	Public Function getUserConfig(f_Num)
		Dim f_RsUserConfig
		Set f_RsUserConfig = User_Conn.Execute("select MemberType,UserConfer,NumberContPoint,NumberLoginPoint,isEmail,isChange,SendPoint,MaxContent,QPoint,IsReg,IsCheck,IsCorpus,IsFavorite,IsMessage,FirstPoint,IsEmailCert,RegOption,UserGroup,BadName,NumberBadLoginPoint,NumberContPassPoint,NumberContBadPoint,BadLoginTime,BadLoginNum from Fs_Config")
		If f_RsUserConfig.EOF Then 
			getUserConfig = False
		Else
			getUserConfig = f_RsUserConfig(f_Num)
		End If
		Set f_RsUserConfig = Nothing 
	End Function 

	Public Function AddCorpus(f_title,f_subtitle,f_Content,f_User,f_Corpus)
		If f_title="" Or f_Content="" Or f_Corpus="" Or f_User="" Then
			AddCorpus = False
		Else
			Dim f_fields,f_values
			f_fields = "UserName,Corpus,Title,SubTitle,Content,AddTime"
			f_values = "'"&NoSqlHack(f_User)&"','"&NoSqlHack(f_Corpus)&"','"&NoSqlHack(f_title)&"','"&NoSqlHack(f_subtitle)&"','"&NoSqlHack(f_Content)&"','"&Now()&"'"
		'	On Error Resume Next 
			User_Conn.Execute("insert into FS_Corpus("&f_fields&") values("&f_values&")")
			If Err Then 
				Err.clear 
				AddCorpus = False
			Else
				AddCorpus = True 
			End if 
		End If 
	End Function 

	Public Function AddLog(f_type,f_StrUserName,f_Strpoints,fs_Strmoneys,f_StrContent,f_Numstyle)'用户编号,点数,金币,描述
		If f_StrUserName="" Or f_Strpoints="" Or fs_Strmoneys="" Then
			AddLog = False
		Else
			dim f_AddlogObj
			Set f_AddlogObj = server.CreateObject(G_FS_RS)
			f_AddlogObj.open "select  * From FS_ME_Log where 1=0",User_Conn,1,3
			f_AddlogObj.addnew
			f_AddlogObj("LogType")=NoSqlHack(f_type)
			f_AddlogObj("UserNumber")=NoSqlHack(f_StrUserName)
			f_AddlogObj("points")=NoSqlHack(f_Strpoints)
			f_AddlogObj("moneys")=NoSqlHack(fs_Strmoneys)
			f_AddlogObj("LogTime")=Now
			f_AddlogObj("LogContent")=NoSqlHack(f_StrContent)
			if f_Numstyle = 0 then
				f_AddlogObj("Logstyle")=0
			Else
				f_AddlogObj("Logstyle")=1
			End if
			f_AddlogObj.update
			f_AddlogObj.close
			set f_AddlogObj = nothing
			If Err Then 
				Err.clear
				AddLog = False
			Else
				AddLog = True
			End If 
		End If 
	End Function 

	Public Function update(f_Fields,f_values,f_NumID)
		If f_Fields="" Or f_values="" Or f_NumID="" Then
			update = False
		Else
			On Error Resume Next 
			Dim f_ArrField,f_ArrValue,f_StrDeal,i
			If InStr(f_Fields,",")>0 And InStr(f_values,",")>0 Then 
				f_ArrField = Split(f_Fields,",")
				f_ArrValue = Split(f_values,",")
				If UBound(f_ArrField) <> UBound(f_ArrValue) Then update = False : Exit Function 
			Else
				f_ArrField = Array(f_Fields)
				f_ArrValue = Array(f_values)
			End If 
			f_StrDeal = ""
			For i=LBound(f_ArrField) To UBound(f_ArrField)
				If i=LBound(f_ArrField) Then 
					f_StrDeal = f_ArrField(i)&"="&f_ArrValue(i)
				Else
					f_StrDeal = f_StrDeal&","&f_ArrField(i)&"="&f_ArrValue(i)
				End If 
			Next 
			User_Conn.Execute("update FS_members set "&f_StrDeal&" where id="&CintStr(f_NumID))
			If Err Then
				Err.clear
				update = False
			Else
				update = True
			End if
		End If
	End Function
End Class

Class Cls_Message
	Private m_RsMessage,m_Number,m_UserName,m_LenContent
	Public Property Let UserName(ByVal StrValue)
		m_UserName = StrValue
		m_RsMessage.open "Select count(MessageID) from FS_ME_Message Where M_ReadUserNumber='"& NoSqlHack(m_UserName) &"' and M_ReadTF=0 and isDelR=0 and isRecyle=0 and isDraft=0",User_Conn,1,1
		m_Number = m_RsMessage(0)
		m_RsMessage.close
	End Property 
	Public Property Get Number()	'未读信息数量
		Number = m_Number
	End Property

	Public Function LenContent(f_StrUserNumber)	'内容总长度
		m_RsMessage.open "Select sum(LenContent) from FS_ME_Message where M_ReadUserNumber='"& NoSqlHack(f_StrUserNumber) &"' and IsDelR = 0",User_Conn,1,3
		LenContent = m_RsMessage(0)
		m_RsMessage.close
	End Function 
	
	Public Function LenbContent(f_StrUserNumber)	'内容总长度
		dim m_book
		set m_book= Server.CreateObject(G_FS_RS)
		m_book.open "Select sum(LenContent) from FS_ME_book where M_ReadUserNumber='"& NoSqlHack(f_StrUserNumber) &"'",User_Conn,1,3
		LenbContent = m_book(0)
		m_book.close
	End Function 

	Private Sub Class_Initialize()
		Set m_RsMessage = server.CreateObject(G_FS_RS)
	End Sub

	Private Sub Class_Terminate()
		Set m_RsMessage = Nothing 
	End Sub

	Public Function update(f_Fields,f_values,f_NumID)
		If f_Fields="" Or f_values="" Or f_NumID="" Then
			update = False
		ElseIf f_NumID="_new_" Then
			On Error Resume Next 
			User_Conn.Execute("insert into FS_Me_Message("&f_Fields&") values("&f_values&")")
			If Err Then
				Err.clear
				update = False
			Else
				update = True
			End if
		Else 
			On Error Resume Next 
			Dim f_ArrField,f_ArrValue,f_StrDeal,i
			If InStr(f_Fields,",")>0 And InStr(f_values,",")>0 Then 
				f_ArrField = Split(f_Fields,",")
				f_ArrValue = Split(f_values,",")
				If UBound(f_ArrField) <> UBound(f_ArrValue) Then update = False : Exit Function 
			Else
				f_ArrField = Array(f_Fields)
				f_ArrValue = Array(f_values)
			End If 
			f_StrDeal = ""
			For i=LBound(f_ArrField) To UBound(f_ArrField)
				If i=LBound(f_ArrField) Then 
					f_StrDeal = f_ArrField(i)&"="&f_ArrValue(i)
				Else
					f_StrDeal = f_StrDeal&","&f_ArrField(i)&"="&f_ArrValue(i)
				End If 
			Next 
			User_Conn.Execute("update FS_Message set "&f_StrDeal&" where MeId in("& FormatIntArr(f_NumID)&")")
			If Err Then
				Err.clear
				update = False
			Else
				update = True
			End if
		End If
	End Function 
	
	Public Function CreateUserDir(f_UserNumber,f_number)
			
	End Function
End Class
%>





