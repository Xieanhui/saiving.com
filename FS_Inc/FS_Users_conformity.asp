<!--#include file="base64.asp"-->
<%
'begin=================整合参数==================
Const HaveDvbbs = 0		'是否整合动网，整合为1，并按说明修改以下动网设置；不整合请填其他值，
Const HaveOblog = 0		'是否整合Oblog，整合为1，并按说明修改以下Oblog设置；不整合请填其他值
Const HaveDz	= 0		'是否整合Discuz5，整合为1，并按说明修改以下Discuz5设置；不整合请填其他值
'end=================整合参数==================

Dim DvConnStr,DvConn,FsConnStr,FsConn,ObConnStr,ObConn,FSstrShowErr,FSDefaultGroupID

'-----系统的虚拟目录,后面不能带/——-----
Const Const_G_VIRTUAL_ROOT_DIR = ""    '后需要/

'begin===================================风讯设置,请按自己的情况修改
If HaveDvbbs=1 Or HaveOblog=1 Then
	'用户文件目录，User、Test/User；
	Dim FSG_USERFILES_DIR,FSDatabaseStr
	FSG_USERFILES_DIR = Const_G_VIRTUAL_ROOT_DIR&"UserFiles"   ''同FS_Inc/Const.asp
	Const FSG_IS_SQL_User_DB = 0  ''1为SQL数据库0为AC数据库
	if FSG_IS_SQL_User_DB = 0 then 
		FSDatabaseStr = "/"&Const_G_VIRTUAL_ROOT_DIR&"Foosun_Data/tbuser.mdb"	'风讯数据库连接
		FsConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(FSDatabaseStr)
	else
		FsConnStr = "Provider=SQLOLEDB.1;Persist Security Info=False;Server=(local);User ID=sa;Password=;Database=fp_trade_user;"
	end if
	Set FsConn = Server.CreateObject(G_FS_CONN)
	FsConn.open FsConnStr
End If
'end=====================================风讯设置

'begin==================================动网设置，请按自己的情况修改
If HaveDvbbs=1 Then
	Dim DvCookiePathStr,DvIndexStr,DvDatabaseStr
	DvCookiePathStr = "/"&Const_G_VIRTUAL_ROOT_DIR&"dvbbs/"  '动网目录的路径，相对根目录，需要”/“
	DvIndexStr = "/"&Const_G_VIRTUAL_ROOT_DIR&"dvbbs/index.asp" '动网首页,与根目录的相对路径
	Const TheCacheName = "DvCache"	'与动网inc/Dv_ClsMain.asp定义CacheName相同的值，搜索CacheName即可找到
	Const Forum_sn = "DvForum"	'与动网inc/Dv_ClsMain.asp定义Forum_sn相同的值，搜索Forum_sn即可找到
	Const DV_IS_SQL_User_DB = 1  ''1为SQL数据库0为AC数据库
	if DV_IS_SQL_User_DB = 0 then 
		DvDatabaseStr = "/"&Const_G_VIRTUAL_ROOT_DIR&"dvbbs/data/dvbbs7.mdb"	'动网数据库连接
		DvConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(DvDatabaseStr)
	else
		DvConnStr = "Provider=SQLOLEDB.1;Persist Security Info=False;Server=(local);User ID=sa;Password=;Database=dvbbs_sql;"
	end if
	Set DvConn = Server.CreateObject(G_FS_CONN)
	DvConn.open DvConnStr
End If
'end=====================================动网设置

'begin===================================Oblog设置,请按自己的情况修改
If HaveOblog=1 Then
	Dim ObIndexStr
	Const ObCookies_name="oblog" 	'Oblog的cookies名,请参照OBLOG目录conn.asp内容设置
	Const ObCookies_domain="" 	'Oblog的cookeies域名根,请参照OBLOG目录config.asp内容设置
	ObIndexStr = "/"&Const_G_VIRTUAL_ROOT_DIR&"Oblog/index.asp" 'Oblog首页，与根目录的相对路径
	Const Ob_is_password_cookies=1		'Oblog保存cookie是否加密，参照OBLOG目录inc/conn.asp内passcookies的值true=1 ,false = 0设置
	Dim ObDatabaseStr
	ObDatabaseStr = "/"&Const_G_VIRTUAL_ROOT_DIR&"oblog/data/oblog4.mdb"		'Oblog数据库，参照oblog目录conn.asp设置
	ObConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(ObDatabaseStr)
	'如果时SQL SERVER数据库，请修改ObConnStr为(数据库连接参数需要修改)：
	'ObConnStr = "Provider=SQLOLEDB.1;Persist Security Info=False;Server=(local);User ID=sa;Password=;Database=Oblog;"
	Set ObConn = Server.CreateObject(G_FS_CONN)
	ObConn.open ObConnStr
End If
'end=====================================Oblog设置

'begin===================================Discuz5设置,请按自己的情况修改
'Discuz5 通行证设置 参考
'应用程序注册地址:user/Reg_service.asp
'应用程序登录地址:user/login.asp
'应用程序退出地址:user/Loginout.asp
If HaveDz = 1 Then
	Dim dzPassportKey,dzMyWebUrl,dzBbsUrl
	dzPassportKey = "1111111111"				'这里换成你discuz论坛通行证设置的passportkey不能小于10位
	dzMyWebUrl    = "http://192.168.1.180/"		'这里换成你的主页绝对地址或相对地址
	dzBbsUrl      = "http://192.168.1.20:81/"	'这里换成你的discuz论坛绝对地址或相对地址
End If
'end=====================================Discuz5设置


'begin====================动网整合接口函数============================
' 动网注册、登陆、退出、修改密码同步
'***********************************************************
' 同步登陆，在风讯会员验证成功后调用
' 如果返回不成功（FALSE），中止风讯和动网的注册
'***********************************************************
Function DvbbsAddUser(StrName,StrPWD,StrEmail,NumSex,StrQuesion,StrAnswer,DvConn)
	Dim DvRs,userclass
	Set DvRs=Server.CreateObject(G_FS_RS)
	DvRs.open "Select * From [Dv_user] Where Username='"&NoSqlHack(StrName)&"'",DvConn,1,3
	If Not DvRs.EOF Then
		DvbbsAddUser = False
		DvRs.Close:Set DvRs=Nothing
		DvbbsAddUser = "该ID已经在论坛注册,\n请重新注册或者联系管理员！"
	Else
		Dim DvTempRsObj
		DvTempRsObj = DvConn.Execute("Select UserTitle,GroupPic,UserGroupID,IsSetting,ParentGID From Dv_UserGroups Where ParentGID=3 Order By MinArticle")

		Dim StatUserID,UserSessionID,UserTrueIP,Startime,BoardID,username,UserID,TruePassWord
		UserTrueIP = NoSqlHack(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))
		If UserTrueIP="" Then UserTrueIP = NoSqlHack(Request.ServerVariables("REMOTE_ADDR"))
		UserTrueIP = CheckStr(UserTrueIP)
		Startime = Timer()
		StatUserID = checkStr(Trim(Request.Cookies(Forum_sn)("StatUserID")))
		If IsNumeric(StatUserID) = 0 or StatUserID = "" Then
			StatUserID = Replace(UserTrueIP,".","")
			UserSessionID = Replace(Startime,".","")
			If IsNumeric(StatUserID) = 0 or StatUserID = "" Then StatUserID = 0
			StatUserID = Ccur(StatUserID) + Ccur(UserSessionID)
		End If
		StatUserID = Ccur(StatUserID)
		BoardID = NoSqlHack(Request("BoardID"))
		If IsNumeric(BoardID) = 0 or BoardID = "" Then BoardID = 0
		BoardID = Clng(BoardID)
		Session(TheCacheName & "UserID") = StatUserID & "_" & Now & "_" & Now & "_" & BoardID		
		username=CheckStr(NoSqlHack(StrName))
		DvRs.addnew
		DvRs("UserName") = NoSqlHack(username)
		DvRs("UserPassword") = NoSqlHack(StrPWD)
		DvRs("TruePassWord") = NoSqlHack(Createpass)
		DvRs("UserEmail") = NoSqlHack(StrEmail)
		DvRs("Userclass") = NoSqlHack(DvTempRsObj(0))
		DvRs("UserQuesion") = NoSqlHack(StrQuesion)
		DvRs("UserAnswer") = NoSqlHack(StrAnswer)
		DvRs("UserLogins")=1
		DvRs("Lockuser")=0
		DvRs("UserWidth")=32
		DvRs("UserHeight")=32
		DvRs("UserFace")="Images/userface/image1.gif"
		DvRs("TitlePic")=DvTempRsObj(1)
		'DvRs("UserMsg")="1||7||动网先锋论坛" 
		DvRs("UserIM")="||||||||||||||||||"
		DvRs("UserInfo")="||||||||||||||||||||||||||||||||||||||||||"
		DvRs("UserSetting")="1|||0|||0"
		DvRs("userWealth")=100
		DvRs("userEP")=60
		DvRs("userCP")=30
		If NumSex=1 Then
			DvRs("UserSex")=1
		Else
			DvRs("UserSex")=0
		End If 
		DvRs("LastLogin")=NOW()
		DvRs("UserGroupID")=DvTempRsObj(2) 
		DvRs("UserPower")=0
		DvRs("UserDel")=0
		DvRs("UserIsbest")=0
		DvRs("UserMoney")=0
		DvRs("UserTicket")=0
		DvRs("UserFav")="陌生人,我的好友,黑名单"
		DvRs("IsChallenge")=0
		DvRs("UserHidden")=0
		DvRs("UserLastIP")=NoSqlHack(Request.ServerVariables("REMOTE_ADDR"),"'","")
		DvRs.Update
		DvConn.Execute("UpDate Dv_Setup Set Forum_UserNum=Forum_UserNum+1,Forum_lastUser='" & HTMLEncode(NoSqlHack(username)) & "'")
		Set DvTempRsObj = Nothing
		Set DvRs = Nothing

		set DvRs=DvConn.Execute("select usertitle,grouppic,UserGroupID from Dv_UserGroups where ParentGID=3 order by minarticle")
		userclass=DvRs(0)
		Set DvRS = Nothing
		Set DvRs = DvConn.Execute("Select TruePassWord,UserID From [Dv_user] Where Username='"&NoSqlHack(UserName)&"'")
		TruePassWord = DvRs("TruePassWord")
		UserID = DvRs("UserID")
		Set DvRs = Nothing

		Response.Cookies(Forum_sn).path=NoSqlHack(DvCookiePathStr)
	    Response.Cookies(Forum_sn)("StatUserID") = NoSqlHack(StatUserID)
		Response.Cookies(Forum_sn)("usercookies") = "0"
	    Response.Cookies(Forum_sn)("username") = NoSqlHack(username)
		Response.Cookies(Forum_sn)("password") = NoSqlHack(TruePassWord)
	    Response.Cookies(Forum_sn)("userclass") = NoSqlHack(userclass)
		Response.Cookies(Forum_sn)("userid") = NoSqlHack(UserID)
		Response.Cookies(Forum_sn)("userhidden") = 2
		session("regtime")=now()
	
		Call RemoveAllCache()
		DvbbsAddUser = True
	End If
End Function

'***********************************************************
' 同步登陆，在风讯会员验证成功后调用
'***********************************************************
Function DvbbsCheckLogin(StrName,StrPWD,DvConn)
	Dim StatUserID,UserSessionID,UserTrueIP,Startime,BoardID,username,UserID,TruePassWord, DvRs,userclass

	Set DvRs = DvConn.Execute("Select TruePassWord,UserID From [Dv_user] Where Username='"&NoSqlHack(StrName)&"' and Userpassword='"&NoSqlHack(StrPWD)&"'")
	If DvRs.EOF Then
		DvbbsCheckLogin = False
		Set DvRs = Nothing
	Else
		TruePassWord = DvRs("TruePassWord")
		UserID = DvRs("UserID")	
		Set DvRs = Nothing

		UserTrueIP = NoSqlHack(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))
		If UserTrueIP = "" Then UserTrueIP = NoSqlHack(Request.ServerVariables("REMOTE_ADDR"))
		UserTrueIP = CheckStr(UserTrueIP)
		Startime = Timer()
		StatUserID = checkStr(Trim(Request.Cookies(Forum_sn)("StatUserID")))
		If IsNumeric(StatUserID) = 0 or StatUserID = "" Then
			StatUserID = Replace(UserTrueIP,".","")
			UserSessionID = Replace(Startime,".","")
			If IsNumeric(StatUserID) = 0 or StatUserID = "" Then StatUserID = 0
			StatUserID = Ccur(StatUserID) + Ccur(UserSessionID)
		End If
		StatUserID = Ccur(StatUserID)
		BoardID = NoSqlHack(Request("BoardID"))
		If IsNumeric(BoardID) = 0 or BoardID = "" Then BoardID = 0
		BoardID = Clng(BoardID)
		Session(TheCacheName & "UserID") = StatUserID & "_" & Now & "_" & Now & "_" & BoardID
		username=CheckStr(StrName)

		Set DvRs=DvConn.Execute("select usertitle,grouppic,UserGroupID from Dv_UserGroups where ParentGID=3 order by minarticle")
		userclass=DvRs(0)
		Set DvRS = Nothing

		Response.Cookies(Forum_sn).path = NoSqlHack(DvCookiePathStr)
	    Response.Cookies(Forum_sn)("StatUserID") = NoSqlHack(StatUserID)
		Response.Cookies(Forum_sn)("usercookies") = "0"
	    Response.Cookies(Forum_sn)("username") = NoSqlHack(username)
		Response.Cookies(Forum_sn)("password") = NoSqlHack(TruePassWord)
	    Response.Cookies(Forum_sn)("userclass") = NoSqlHack(userclass)
		Response.Cookies(Forum_sn)("userid") = NoSqlHack(UserID)
		Response.Cookies(Forum_sn)("userhidden") = 2
		Call RemoveAllCache()
		DvbbsCheckLogin = True
	End If
End Function

'***********************************************************
' 同步退出，在风讯退出时调用
' 清空动网本地信息，更新缓存
'***********************************************************
Sub DvCleanCookie()
	Response.Cookies(Forum_sn).path=DvCookiePathStr
	Response.Cookies(Forum_sn)("username")=""
	Response.Cookies(Forum_sn)("password")=""
	Response.Cookies(Forum_sn)("userclass")=""
	Response.Cookies(Forum_sn)("userid")=""
	Response.Cookies(Forum_sn)("userhidden")=""
	Response.Cookies(Forum_sn)("usercookies")=""
	Session(TheCacheName & "UserID")=Empty
	Session("flag")=Empty
	Call RemoveAllCache()
End Sub

'***********************************************************
' 更改同步更改密码
' 在风讯会员更改密码验证成功后调用此函数同步更改动网会员密码
' 参数：会员名，会员旧密码，会员新密码，动网连接
' 返回：成功(True)失败(False)
'***********************************************************
Function DvChangePWD(StrName,Stranswer,StrNewPWD,DvConn)
	If StrName="" Or Len(Stranswer)<>16 Then
		DvChangePWD = False
	Else
		DvConn.Execute("Update Dv_user set UserPassword='"&NoSqlHack(StrNewPWD)&"' where UserName='"&NoSqlHack(StrName)&"' and UserAnswer='"&NoSqlHack(Stranswer)&"'")
		DvChangePWD = True
	End If
End Function
'end====================动网整合接口函数=======================================


'begin====================Oblog整合接口函数====================================
' Oblog的同步注册、登陆、退出和修改密码函数，在其他系统整合OBLOG调用
'***********************************************************
' 其他系统注册会员时同步注册Oblog会员
'***********************************************************
Function ObAddUser(StrName,StrPWD,StrEmail,StrSex,StrQuesion,StrAnswer,StrNikeName,ObConn)
	Dim RsOblog,ObCache,ObCache1,ObUserLevel,ObUserIP
	if StrSex=1 then StrSex=0 else StrSex=1 end if
	ObUserLevel = 7
	ObCache = ObConn.Execute("select * from [oblog_setup]").GetRows(1)
	ObCache1 = ObConn.Execute("select groupid from [oblog_groups] where g_level=1").GetRows(1)
	ObUserIP = request.ServerVariables("HTTP_X_FORWARDED_FOR")
	If ObUserIP = "" Then ObUserIP = request.ServerVariables("REMOTE_ADDR")

	set RsOblog=Server.CreateObject(G_FS_RS)
	RsOblog.open "select * from oblog_user where username='"& NoSqlHack(StrName) &"'",ObConn,1,3
	If Not RsOblog.EOF Then
		RsOblog = "用户名已经存在Oblog"
	Else
		RsOblog.AddNew
        RsOblog("username") = NoSqlHack(StrName)
        RsOblog("password") = NoSqlHack(StrPWD)

        RsOblog("question") = NoSqlHack(StrQuesion)
        RsOblog("answer") = NoSqlHack(StrAnswer)
        RsOblog("useremail") = NoSqlHack(StrEmail)
        RsOblog("user_level") = NoSqlHack(ObUserLevel)
        RsOblog("user_isbest") = 0
        RsOblog("blogname") = NoSqlHack(StrName)&"的blog"
        RsOblog("user_classid") = 1
        'RsOblog("nickname")=nickname
        RsOblog("province") = ""
        RsOblog("city") = ""
        RsOblog("sex") = NoSqlHack(StrSex)
        RsOblog("adddate") = Now()
        RsOblog("regip") = NoSqlHack(ObUserIP)
        RsOblog("lastloginip") = NoSqlHack(ObUserIP)
        RsOblog("lastlogintime") = Now()
        RsOblog("user_dir") =ObCache(8,0)
        RsOblog("user_folder") = NoSqlHack(StrName)
        RsOblog("user_group") = ObCache1(0,0)
        RsOblog("scores") = 100
        RsOblog("newbie") = 1
		RsOblog("comment_isasc")=0
		RsOblog("lockuser")=0
		RsOblog.update
		RsOblog.close:Set RsOblog=Nothing
		ObConn.execute("update [oblog_setup] set user_count=user_count+1")
		Session("chk_regtime") = Now()
        If ObCookies_domain <> "" Then
            response.Cookies(ObCookies_name).domain = ObCookies_domain
        End If
        response.Cookies(ObCookies_name)("UserName") = ObCodeCookie(StrName)
        response.Cookies(ObCookies_name)("password") = ObCodeCookie(StrPWD)
        Response.Cookies(ObCookies_name)("CookieDate") = 365
		Response.Cookies(ObCookies_name).Expires=Date+365
		RsOblog = True
	End If
End Function 

'***********************************************************
' 同步登陆
' 如果整合后Oblog中没有该会员，则自动添加，但是需要会员完善自己的Oblog部分资料
'***********************************************************
Function OblogCheckLogin(StrName,StrPWD,ObConn)
	If StrName="" Or StrPWD="" Or IsNull(StrName) Or IsNull(StrPWD) Then
		OblogCheckLogin = "用户名或者密码错误！"
		Exit Function
	End If
	Dim RsOblog,ObCache1,ObUserIP,UserTF,ObCache,ObUserLevel,FS_RS_DB
	ObUserLevel = 7
	ObCache = ObConn.Execute("select * from [oblog_setup]").GetRows(1)
	ObCache1 = ObConn.Execute("select groupid from [oblog_groups] where g_level=1").GetRows(1)
	Set UserTF = ObConn.execute("select username from oblog_user where username='"&NoSqlHack(StrName)&"'")
	ObUserIP =NoSqlHack(request.ServerVariables("HTTP_X_FORWARDED_FOR"))
	If ObUserIP = "" Then ObUserIP = NoSqlHack(request.ServerVariables("REMOTE_ADDR"))
	Set RsOblog = Server.CreateObject(G_FS_RS)
	RsOblog.open "select * from oblog_user where username='"&NoSqlHack(StrName)&"' and password ='"&NoSqlHack(StrPWD)&"'",ObConn,1,3
	If RsOblog.EOF Or RsOblog.BOF Then
		If Not UserTF.EOF Then
			OblogCheckLogin = "Oblog已经存在同名的用户"
			Exit Function
		End If
		dim StrQuesion,StrAnswer,StrEmail,StrSex,Province,City
		set FS_RS_DB = FSConn.execute("select * from FS_ME_Users where UserName='"&NoSqlHack(StrName)&"'")
		if not FS_RS_DB.eof then 
			StrQuesion = FS_RS_DB("PassQuestion")
			StrAnswer = FS_RS_DB("PassAnswer")
			StrEmail = FS_RS_DB("Email")
			StrSex = FS_RS_DB("sex")
			Province = FS_RS_DB("Province")
			City = FS_RS_DB("City")
		end if
		FS_RS_DB.close
		if cint(StrSex)=1 then StrSex=0 else StrSex=1 end if
		RsOblog.AddNew
        RsOblog("username") = NoSqlHack(StrName)
        RsOblog("password") = NoSqlHack(StrPWD)

        RsOblog("question") = NoSqlHack(StrQuesion)
        RsOblog("answer") = NoSqlHack(StrAnswer)
        RsOblog("useremail") = NoSqlHack(StrEmail)
        RsOblog("user_level") = NoSqlHack(ObUserLevel)
        RsOblog("user_isbest") = 0
        RsOblog("blogname") = NoSqlHack(StrName)&"的blog"
        RsOblog("user_classid") = 1
        'RsOblog("nickname")=nickname
        RsOblog("province") = NoSqlHack(Province)
        RsOblog("city") = NoSqlHack(City)
        RsOblog("sex") = NoSqlHack(StrSex)
        RsOblog("adddate") = Now()
        RsOblog("regip") = NoSqlHack(ObUserIP)
        RsOblog("lastloginip") = NoSqlHack(ObUserIP)
        RsOblog("lastlogintime") = Now()
        RsOblog("user_dir") =ObCache(8,0)
        RsOblog("user_folder") = NoSqlHack(StrName)
        RsOblog("user_group") = ObCache1(0,0)
        RsOblog("scores") = 100
        RsOblog("newbie") = 1
        RsOblog("log_count") = 1
		RsOblog("comment_isasc")=0
		RsOblog("lockuser")=0
		RsOblog.update
		RsOblog.close:Set RsOblog=Nothing
		ObConn.execute("update [oblog_setup] set user_count=user_count+1,log_count=log_count+1")
		Session("chk_regtime") = Now()
        If ObCookies_domain <> "" Then
            response.Cookies(ObCookies_name).domain = ObCookies_domain
        End If
        response.Cookies(ObCookies_name)("UserName") = ObCodeCookie(StrName)
        response.Cookies(ObCookies_name)("password") = ObCodeCookie(StrPWD)
        Response.Cookies(ObCookies_name)("CookieDate") = 365
		Response.Cookies(ObCookies_name).Expires=Date+365
		OblogCheckLogin = "向0BLOG添加用户，需要完善部分资料！"
	Else
		If RsOblog("lockuser")=true Then
			OblogCheckLogin = False
			RsOblog.close:Set RsOblog=Nothing
			Exit Function
		End If
		RsOblog("LastLoginIP") = ObUserIP
		RsOblog("LastLoginTime") = Now()
		RsOblog("LoginTimes") = RsOblog("LoginTimes") + 1
		RsOblog("log_count") = RsOblog("log_count") + 1
		RsOblog.update
		RsOblog.close:Set RsOblog=Nothing
		If ObCookies_domain <> "" Then
            response.Cookies(ObCookies_name).domain = ObCookies_domain
        End If
		ObConn.execute("update [oblog_setup] set user_count=user_count+1,log_count=log_count+1")
		Session("chk_regtime") = Now()
        If ObCookies_domain <> "" Then
            response.Cookies(ObCookies_name).domain = ObCookies_domain
        End If
        response.Cookies(ObCookies_name)("UserName") = ObCodeCookie(StrName)
        response.Cookies(ObCookies_name)("password") = ObCodeCookie(StrPWD)
        Response.Cookies(ObCookies_name)("CookieDate") = 365
		Response.Cookies(ObCookies_name).Expires=Date+365
		OblogCheckLogin = "同步登陆成功！"
	End If
End Function

'***********************************************************
' 退出Oblog,清空COOKIE信息
'***********************************************************
Sub OblogLogOut()
	If ObCookies_domain <> "" Then
		response.Cookies(ObCookies_name).domain = ObCookies_domain
	End If
	Response.Cookies(ObCookies_name)("UserName")=ObCodeCookie("")
	Response.Cookies(ObCookies_name)("Password")=ObCodeCookie("")
	Response.Cookies(ObCookies_name)("userlevel")=ObCodeCookie("0")
End Sub

'***********************************************************
' 同步修改Oblog密码
'***********************************************************
Function ObChangePWD(StrName,Stranswer,StrNewPWD,ObConn)
	If StrName="" Or Stranswer="" Then
		ObChangePWD = "Oblog用户名或回答不能为空"
	Else
		Dim RsOblog
		Set RsOblog = Server.CreateObject(G_FS_RS)
		RsOblog.open "select password from oblog_user where UserName='"&NoSqlHack(StrName)&"' and Answer='"&NoSqlHack(Stranswer)&"'",ObConn,1,3
		If RsOblog.EOF Then
			ObChangePWD = "Oblog没有该用户或用户密码错误"
			Exit Function 
		End If
		RsOblog("password")=StrNewPWD
		RsOblog.update
		If ObCookies_domain <> "" Then
            response.Cookies(ObCookies_name).domain = ObCookies_domain
        End If
        response.Cookies(ObCookies_name)("userpassword") = ObCodeCookie(StrNewPWD)
		ObChangePWD = True
	End If
End Function

'end====================Oblog整合接口函数=============================


'=================================================================
' 整合需要的公用函数
'=================================================================
Function setStrLoc1(FS_Str,FS_StrLoc,FS_StrLen,FS_StrRep)
	If Len(FS_Str)>FS_StrLoc Then
		FS_StrLen = CInt(FS_StrLen)
		Fs_strRep = Right(String(FS_StrLen,"0")&FS_StrRep,Fs_StrLen)
		If CInt(FS_StrLoc)=1 Then 
			setStrLoc1 = FS_StrRep & Right(FS_Str,Len(FS_Str)-FS_StrLoc+1-FS_StrLen)
		Else
			setStrLoc1 = Left(FS_Str,FS_StrLoc-1) & FS_StrRep & Right(FS_Str,Len(FS_Str)-FS_StrLoc+1-FS_StrLen)
		End If
	Else
		setStrLoc1 = FS_Str & FS_StrRep
	End If
End Function
Function Checkstr(Str)
	If Isnull(Str) Then
		CheckStr = ""
		Exit Function 
	End If
	Str = Replace(Str,Chr(0),"")
	CheckStr = Replace(Str,"'","''")
End Function

Function Createpass()
	Dim Ran,i,LengthNum
	LengthNum=16
	Createpass=""
	For i=1 To LengthNum
		Randomize
		Ran = CInt(Rnd * 2)
		Randomize
		If Ran = 0 Then
			Ran = CInt(Rnd * 25) + 97
			Createpass =Createpass& UCase(Chr(Ran))
		ElseIf Ran = 1 Then
			Ran = CInt(Rnd * 9)
			Createpass = Createpass & Ran
		ElseIf Ran = 2 Then
			Ran = CInt(Rnd * 25) + 97
			Createpass =Createpass& Chr(Ran)
		End If
	Next
End Function

Function HTMLEncode(fString)
	If Not IsNull(fString) Then
		fString = replace(fString, ">", "&gt;")
		fString = replace(fString, "<", "&lt;")
		fString = Replace(fString, CHR(32), " ")		'&nbsp;
		fString = Replace(fString, CHR(9), " ")			'&nbsp;
		fString = Replace(fString, CHR(34), "&quot;")
		'fString = Replace(fString, CHR(39), "&#39;")	'单引号过滤
		fString = Replace(fString, CHR(13), "")
		fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
		fString = Replace(fString, CHR(10), "<BR> ")
		HTMLEncode = fString
	End If
End Function

Sub RemoveAllCache()
	Dim cachelist,i
	Cachelist=split(GetallCache(),",")
	If UBound(cachelist)>1 Then
		For i=0 to UBound(cachelist)-1
			DelCahe Cachelist(i)
		Next
	End If
End Sub

Function  GetallCache()
	Dim Cacheobj
	For Each Cacheobj in Application.Contents
		GetallCache = GetallCache & Cacheobj & ","
	Next
End Function

Sub DelCahe(MyCaheName)
	Application.Lock
	Application.Contents.Remove(MyCaheName)
	Application.unLock
End Sub

Function ObCodeCookie(Str)
    If Ob_is_password_cookies = 1 Then
        Dim i
        Dim StrRtn
        For i = Len(Str) To 1 Step -1
            StrRtn = StrRtn & AscW(Mid(Str, i, 1))
            If (i <> 1) Then StrRtn = StrRtn & "a"
        Next
        ObCodeCookie = StrRtn
    Else
        ObCodeCookie = Str
    End If
End Function


''--------------------------------新增 awen
Function ObDecodeCookie(Str)
if Ob_is_password_cookies = 1 then
	Dim i
	Dim StrArr,StrRtn
	StrArr = Split(Str,"a")
	For i = 0 to UBound(StrArr)
		If isNumeric(StrArr(i)) = True Then
			StrRtn = Chrw(StrArr(i)) & StrRtn
		Else
			StrRtn = Str
			Exit Function
		End If
	Next
	ObDecodeCookie = StrRtn
else
	ObDecodeCookie=str
end if
End Function

Sub doMsg_awen(msg,url)
	Response.Write"<script language=JavaScript>"
	Response.Write"alert("""&msg&""");"
	if url<>"" then 
		response.Write("window.location='"&url&"';")
	else	
		Response.Write"window.history.go(-1);"
	end if
	Response.Write"</script>"
	response.End()
End Sub

'得到多少位数的随机函数
Function GetRamCode(f_number)
	Randomize
	Dim f_Randchar,f_Randchararr,f_RandLen,f_Randomizecode,f_iR
	f_Randchar="0,1,2,3,4,5,6,7,8,9,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z"
	f_Randchararr=split(f_Randchar,",")
	f_RandLen=f_number '定义密码的长度或者是位数
	for f_iR=1 to f_RandLen
		f_Randomizecode=f_Randomizecode&f_Randchararr(Int((21*Rnd)))
	next
	GetRamCode = f_Randomizecode
End Function

''===========================================
''以下是风讯的，用于Oblog,dvBBS中
''===========================================
Function Login(f_StrName,f_StrPWD,f_Logintye)
	Dim p_LoginLockNum,LoginLockNum
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
				FSstrShowErr = "您已经连续登陆了"& p_LoginLockNum -1 &"次\n此帐户已经临时被锁定，今天不能登陆了!!!"
				Call doMsg_awen(FSstrShowErr,"")
		End if
	End if
	if f_StrName = "" or  f_StrPWD = "" then
		Login = false 
		FSstrShowErr = "请填写您的用户名\n请填写您的密码!"
		Call doMsg_awen(FSstrShowErr,"")
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
			FSstrShowErr = "错误的参数\n请选择登陆方式!"
			Call doMsg_awen(FSstrShowErr,"")
		 End if
		Set f_RsLoginobj = Server.CreateObject(G_FS_RS)
		f_RsLoginobj.open f_RsLoginSQL,FsConn,1,1
		If Not f_RsLoginobj.eof then 
			If f_RsLoginobj(5)<>0 then
				f_RsLoginobj.close
				set f_RsLoginobj = nothing
				Login = false 
				FSstrShowErr = "用户已经被锁定此用户注册没有审核!"
				Call doMsg_awen(FSstrShowErr,"")
			Else
				'更新数据
				Dim p_NumGetPoint,p_LoginGetMoney,p_LoginGetIntegral 
				Dim FSRegisterTFRS
				p_NumGetPoint=0 : p_LoginGetMoney=0 : FSDefaultGroupID = 1
				set FSRegisterTFRS=FsConn.execute("select top 1 LoginPointmoney,DefaultGroupID from FS_ME_SysPara")
				if not FSRegisterTFRS.eof then 
					if not isnull(FSRegisterTFRS("DefaultGroupID")) then FSDefaultGroupID = FSRegisterTFRS("DefaultGroupID")
					if not isnull(FSRegisterTFRS("LoginPointmoney")) then if instr(FSRegisterTFRS("LoginPointmoney"),",") then p_LoginGetMoney=split(FSRegisterTFRS("LoginPointmoney"),",")(0) : p_LoginGetIntegral=split(FSRegisterTFRS("LoginPointmoney"),",")(1)
				end if
				FSRegisterTFRS.close

				Dim f_RsUpdateobj,f_RsUpdateSQL
				Set f_RsUpdateobj = Server.CreateObject(G_FS_RS)
					if f_Logintye = "0" then
						 f_RsUpdateSQL = "Select UserNumber,LoginNum,LastLoginTime,FS_Money,Integral,UserLoginCode,TempLastLoginTime,TempLastLoginTime_1,LastLoginIP,GroupID  From FS_ME_Users Where UserName = '"& NoSqlHack(f_StrName) &"' and UserPassword = '"& NoSqlHack(f_StrPWD) &"'"
					 Elseif f_Logintye = "1" then
						 f_RsUpdateSQL = "Select UserNumber,LoginNum,LastLoginTime,FS_Money,Integral,UserLoginCode,TempLastLoginTime,TempLastLoginTime_1,LastLoginIP,GroupID From FS_ME_Users Where UserNumber = '"& NoSqlHack(f_StrName) &"' and UserPassword = '"& NoSqlHack(f_StrPWD) &"'"
					 Elseif f_Logintye = "2" then
						 f_RsUpdateSQL = "Select UserNumber,LoginNum,LastLoginTime,FS_Money,Integral,UserLoginCode,TempLastLoginTime,TempLastLoginTime_1,LastLoginIP,GroupID From FS_ME_Users Where Email = '"& NoSqlHack(f_StrName) &"' and UserPassword = '"& NoSqlHack(f_StrPWD) &"'"
					 End if
				f_RsUpdateobj.open  f_RsUpdateSQL,FsConn,1,3
				f_RsUpdateobj("LoginNum")=f_RsUpdateobj("LoginNum")+1
				f_RsUpdateobj("LastLoginTime")=now
				f_RsUpdateobj("GroupID")=FSDefaultGroupID
				f_RsUpdateobj("LastLoginIP")=Request.ServerVariables("Remote_Addr")
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

''--------------------------------------注册
Function FSAddUser(p_UserName_1,p_UserPassword_1,p_PassQuestion_1,p_PassAnswer_1,p_safeCode_1,p_Email_1,p_NickName,p_Sex)
	Dim FSRegisterTF,p_NumGetPoint,p_NumGetMoney,p_RegisterCheck
	Dim FSRegisterTFRS
	FSRegisterTF = 1 : p_RegisterCheck = 0 : p_NumGetPoint=0 : p_NumGetMoney=0 : FSDefaultGroupID = 1
	set FSRegisterTFRS=FsConn.execute("select top 1 RegisterTF,RegPointmoney,RegisterCheck,DefaultGroupID from FS_ME_SysPara")
	if not FSRegisterTFRS.eof then 
		if not isnull(FSRegisterTFRS("RegisterTF")) then FSRegisterTF = FSRegisterTFRS("RegisterTF")
		if not isnull(FSRegisterTFRS("RegisterCheck")) then p_RegisterCheck = FSRegisterTFRS("RegisterCheck")
		if not isnull(FSRegisterTFRS("DefaultGroupID")) then FSDefaultGroupID = FSRegisterTFRS("DefaultGroupID")
		if not isnull(FSRegisterTFRS("RegPointmoney")) then if instr(FSRegisterTFRS("RegPointmoney"),",") then p_NumGetPoint=split(FSRegisterTFRS("RegPointmoney"),",")(0) : p_NumGetMoney=split(FSRegisterTFRS("RegPointmoney"),",")(1)
	end if
	FSRegisterTFRS.close
	if FSRegisterTF<>1 then FSAddUser=False:exit Function
	
	Dim AddUserDataTFObj,UserNumberRuleObj,AddUserDataObj
	Set AddUserDataTFObj = Server.CreateObject(G_FS_RS)
	AddUserDataTFObj.open "select  UserName,Email From FS_ME_Users where UserName = '"& NoSqlHack(p_UserName_1) &"'",FsConn,1,3
	If Not AddUserDataTFObj.eof then
				FSstrShowErr = "您提交的用户名或者电子邮件已经被注册!"
				Call doMsg_awen(FSstrShowErr,"")
	End if
	AddUserDataTFObj.close:set AddUserDataTFObj =nothing
	'判断用户编号是否重复
	Set UserNumberRuleObj = Server.CreateObject(G_FS_RS)
	UserNumberRuleObj.open "select  UserNumber From FS_ME_Users where UserName='"& NoSqlHack(p_UserName_1)&"'",FsConn,1,1
	If Not UserNumberRuleObj.eof then
			FSstrShowErr = "您提交的用户编号意外重复，非常抱歉，请重新填写用户资料。!"
			Call doMsg_awen(FSstrShowErr,"")
	End if
	
	'插入用户数据
	Dim  strUserNumberRule
	strUserNumberRule =  GetRamCode(11)
	
	Set AddUserDataObj = Server.CreateObject(G_FS_RS)
	AddUserDataObj.open "select * From FS_ME_Users where 1=0",FsConn,1,3
	AddUserDataObj.addNew
	AddUserDataObj("UserNumber") = NoSqlHack(strUserNumberRule)
	AddUserDataObj("UserName") = NoSqlHack(p_UserName_1)
	AddUserDataObj("UserPassword") = NoSqlHack(p_UserPassword_1)
	AddUserDataObj("PassQuestion") = NoSqlHack(p_PassQuestion_1)
	AddUserDataObj("PassAnswer") = NoSqlHack(p_PassAnswer_1)
	AddUserDataObj("safeCode") = NoSqlHack(p_safeCode_1)
	AddUserDataObj("Email") = NoSqlHack(p_Email_1)
	AddUserDataObj("isMessage") = 0
	AddUserDataObj("HeadPicsize") = "60,60"
	AddUserDataObj("NickName") = NoSqlHack(p_NickName)
	AddUserDataObj("RealName") = "未填"
	AddUserDataObj("Province") = "未填"
	AddUserDataObj("city") = "未填"
	AddUserDataObj("Sex") = NoSqlHack(p_Sex	)
	AddUserDataObj("IsCorporation") = 0
	AddUserDataObj("RegTime") = now
	AddUserDataObj("CloseTime") = "3000-1-1"
	AddUserDataObj("LoginNum") = 0
	AddUserDataObj("Integral") = NoSqlHack(p_NumGetPoint)
	AddUserDataObj("FS_Money") = NoSqlHack(p_NumGetMoney)
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
	AddUserDataObj("GroupID") = NoSqlHack(FSDefaultGroupID)
	AddUserDataObj.Update
	AddUserDataObj.close:set AddUserDataObj = nothing
	'更新数据，获得相应期限或者金币，积分
	'说明：如果为0，则不限制
	'开始建立对象
	Dim rsCreatGroup 
	set rsCreatGroup =FsConn.execute("select GroupID,GroupPoint,GroupMoney,GroupDate From FS_ME_Group where GroupID="&CintStr(FSDefaultGroupID))
	if not rsCreatGroup.eof then
		if rsCreatGroup("GroupPoint")>0 then
			FsConn.execute("Update FS_ME_Users Set Integral=Integral+"& rsCreatGroup("GroupPoint")&" where UserNumber='"& NoSqlHack(strUserNumberRule) &"'")
		end if
		if rsCreatGroup("GroupMoney")>0 then
			FsConn.execute("Update FS_ME_Users Set FS_Money=FS_Money+"& rsCreatGroup("GroupMoney") &" where UserNumber='"& strUserNumberRule &"'")
		end if
		if rsCreatGroup("GroupDate")>0 then
			dim DateCoseed
			DateCoseed = dateAdd("d",rsCreatGroup("GroupDate"),date)
			if FSG_IS_SQL_User_DB=0 then
				FsConn.execute("Update FS_ME_Users Set CloseTime=#"& NoSqlHack(DateCoseed) &"# where UserNumber='"& NoSqlHack(strUserNumberRule) &"'")
			else
				FsConn.execute("Update FS_ME_Users Set CloseTime='"& NoSqlHack(DateCoseed) &"' where UserNumber='"& NoSqlHack(strUserNumberRule) &"'")
			end if
		end if
	end if
	rsCreatGroup.close:set rsCreatGroup = nothing
	
	session("FS_UserName") = NoSqlHack(p_UserName_1)
	session("FS_UserNumber") = NoSqlHack(strUserNumberRule)
	session("FS_NickName") = NoSqlHack(p_NickName)
	session("FS_UserPassword") = NoSqlHack(p_UserPassword_1)
	session("FS_IsCorp") = 0
	session("FS_UserEmail") = NoSqlHack(p_Email_1)
	session("FS_IsLock") = NoSqlHack(p_RegisterCheck)
	Response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") = 1
	Call InsertMyPara(session("FS_UserNumber") )
	Call AddLog("注册",session("FS_UserNumber"),p_NumGetPoint,p_NumGetMoney,"注册获得积分",0) 
	Dim str_isSendMail,FsoObj,Path
	Set FsoObj = Server.CreateObject(G_FS_FSO)  
	Path = strUserNumberRule
	if FsoObj.FolderExists(Server.MapPath("/"&FSG_USERFILES_DIR) ) = false then FsoObj.createFolder Server.MapPath("/"&FSG_USERFILES_DIR) 
	Path = Server.MapPath("/"&FSG_USERFILES_DIR&"/"&Path) 
	if FsoObj.FolderExists(Path) = True then FsoObj.deleteFolder Path
	FsoObj.CreateFolder Path
	str_isSendMail=false
	
	FSAddUser = True
End Function

Function InsertMyPara(f_strUserNumber)
		Dim f_Rsmypara
		Set f_Rsmypara = Server.CreateObject(G_FS_RS) 
		f_Rsmypara.open "select  * From FS_ME_MySysPara where 1=0",FsConn,1,3
		f_Rsmypara.addnew
		f_Rsmypara("DownFileRule") = ",,,,"
		f_Rsmypara("NewsFileRule") = ",,,,"
		f_Rsmypara("ProductFileRule") = ",,,,"
		f_Rsmypara("ilogFileRule") = ",,,,"
		f_Rsmypara("mysiteName") = "我的个人空间"
		f_Rsmypara("UserNumber") = NoSqlHack(f_strUserNumber)
		f_Rsmypara("Keywords") = "风讯,CMS,Foosun"
		f_Rsmypara("Description") = "风讯,CMS,Foosun"
		f_Rsmypara("NaviPic") = ""
		f_Rsmypara("isHtml") = 0
		'f_Rsmypara("RedirectUrl") = ""
		f_Rsmypara.update
		f_Rsmypara.close:Set f_Rsmypara = nothing
End Function

Function AddLog(f_type,f_StrUserName,f_Strpoints,fs_Strmoneys,f_StrContent,f_Numstyle)'用户编号,点数,金币,描述
	If f_StrUserName="" Or f_Strpoints="" Or fs_Strmoneys="" Then
		AddLog = False
	Else
		dim f_AddlogObj
		Set f_AddlogObj = Server.CreateObject(G_FS_RS)
		f_AddlogObj.open "select  * From FS_ME_Log where 1=0",FsConn,1,3
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
''---------修改密码
Function ChangePWD(f_StrName,Stranswer,StrNewPWD)
	If f_StrName="" Or len(Stranswer)<>16 Then
		ChangePWD = "帐号或回答不能为空"
	Else
		Dim ObjPWD
		Set ObjPWD = Server.CreateObject(G_FS_RS)
		objPWD.open "select UserPassword from FS_ME_Users where UserName='"&NoSqlHack(f_StrName)&"' and PassAnswer='"&NoSqlHack(Stranswer)&"'",FsConn,1,3
		If Not ObjPWD.EOF Then
			ObjPWD("UserPassword")=StrNewPWD
			ObjPWD.update
			ChangePWD = True
		Else
			ChangePWD = "您不是风讯会员"
		End If
		objPWD.close
		set objPWD=nothing		
	End If
End Function
''---------------------------------------------退出
Sub FSout()
	Session("FS_UserName") = ""
	Session("FS_UserNumber") = ""
	Session("FS_UserPassword") = ""
	Session("FS_Group") = ""
	Session("FS_IsCorp") = ""
	Session("FS_NickName") = ""
	response.Cookies("FoosunUserCookies")("UserLogin_Style_Num")  = ""
	Session("UserLoginCode") = ""
End Sub

'========discuz5整合接口函数==========================================
'---------------------------------------------------------------------------------
Function DzCheckLogin(p_UserName,p_UserPassword,dzForward,dzEmail,forward)'与Dz同步登录
	Dim DzRs,TempEmail,dzMember,dzVerify,dzAuth
	dzMember=	"time=" 	 	& datediff("s","1970-1-1 00:00:00",now) &_
				"&username=" 	& p_UserName &_
				"&password=" 	& p_UserPassword &_
				"&email="	 	& dzEmail &_
				"&cookietime="	& 0
	dzAuth = passport_encrypt(dzMember,dzPassportKey)
	Response.Cookies("auth")=dzAuth
	If forward="" Then
		dzForward=dzMyWebUrl & "user/main.asp"
	Else
		dzForward=forward
	End If
	dzVerify = md5("login" & dzAuth & dzForward & dzPassportKey,32)
	dzAuth=server.URLEncode(dzAuth)
	dzForward=server.URLEncode(dzForward)
	response.Redirect(dzBbsUrl & "api/passport.php?action=login&auth=" & NoSqlHack(dzAuth) & "&forward=" & NoSqlHack(dzForward) & "&verify=" & NoSqlHack(dzVerify))
End Function
'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------
Sub DzCleanCookie(forward)'注销用户
	Dim dz_Forward,dz_Verify,dz_Auth
	if forward="" then
		dz_Forward=dzMyWebUrl & "user/login.asp"
	else
		dz_Forward=forward
	end If
	dz_Auth=Request.Cookies("auth")
	dz_Verify = md5("logout"& dz_Auth & dz_Forward & dzPassportKey ,32)
	dz_Auth=server.URLEncode(dz_Auth)
	dz_Forward=server.URLEncode(dz_Forward)
	Response.Redirect(dzBbsUrl & "api/passport.php?action=logout&auth=" & NoSqlHack(dz_Auth) & "&forward=" & NoSqlHack(dz_Forward) & "&verify=" & NoSqlHack(dz_Verify))
End Sub
'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------
Function DzReg(p_UserName,p_UserPassword,P_Email,forward)'与Dz同步注册
	Dim dzAuth,dzForward,dzMember,dzVerify
	dzMember =	"time=" 	 	& datediff("s","1970-1-1 00:00:00",now) &_
				"&username=" 	& p_UserName &_
				"&password=" 	& p_UserPassword &_
				"&email="	 	& P_Email
	dzAuth = passport_encrypt(dzMember , dzPassportKey)
	If forward="" Then
		dzForward=dzMyWebUrl & "user/main.asp"
	Else
		dzForward=forward
	End If
	dzVerify = md5("login" & dzAuth & dzForward & dzPassportKey,32)
	dzAuth=server.URLEncode(dzAuth)
	dzForward=server.URLEncode(dzForward)
	DzReg=dzBbsUrl & "api/passport.php?action=login&auth=" & NoSqlHack(dzAuth) & "&forward=" & NoSqlHack(dzForward) & "&verify=" & NoSqlHack(dzVerify)
End Function
'---------------------------------------------------------------------------------

function passport_encrypt(txt, key) 
		dim encrypt_key, encrypt_key_ary,txt_ary
		dim ctr,tmp,i      
        Randomize
		encrypt_key=md5(Int(32000* Rnd),32)
		encrypt_key_ary=strtoary(encrypt_key)		
        txt_ary=strtoary(txt)      
        ctr = 0
        tmp = ""		
        for i = 0 to StrLength(txt)-1
			if ctr=len(encrypt_key)  then 	ctr=0  else ctr=ctr				
            tmp = tmp & encrypt_key_Ary(ctr) &  mxor(txt_ary(i),encrypt_key_ary(ctr))
            ctr=ctr+1				                     
        next
		passport_encrypt=strAnsi2Unicode(base64Encode(passport_key(tmp, key)))		
end function 

'=====================================================
'==========仿discuz Passport 密匙处理函数=============
'=====================================================
function passport_key(annitxt, encrypt_key) 
		dim encrypt_key_tmp,encrypt_key_ary,txt_ary
		dim ctr,tmp,i
		encrypt_key_tmp = md5(encrypt_key,32)		
        encrypt_key_ary=StrToAry(encrypt_key_tmp)		
		txt_ary=AnniToAry(annitxt) 		
		ctr = 0
        tmp = ""		   
        for i=0 to lenb(annitxt)-1			
		    if ctr=len(encrypt_key_tmp) then ctr=0 else ctr=ctr
            tmp= tmp & mxor(txt_ary(i),encrypt_key_ary(ctr))
            ctr=ctr+1
        next
		passport_key=tmp		
end function


function StrToAry(str)
   dim ary(),length,tmpstr,i
   tmpstr=strUnicode2Ansi(str)
   length=lenb(tmpstr)   
   redim ary(length)   
   for i=0 to length-1    		 		
		ary(i)=midb(tmpstr,i+1,1)		
   next  
   StrToAry=ary
end function

function AnniToAry(str)
   dim ary(),length,i
   length=lenb(str)   
   redim ary(length)   
   for i=0 to length-1    		 		
		ary(i)=midb(str,i+1,1)		
   next  
   AnniToAry=ary
end function


'=====================================================
'=================异或操作============================
'=====================================================
function mxor(chrb1,chrb2)
	if chrb1<>"" and chrb2 <>"" then
	mxor=chrb(ascb(chrb1) xor ascb(chrb2))
	end if	
end function


Function StrLength(str)
	ON ERROR RESUME NEXT
	Dim WINNT_CHINESE
	WINNT_CHINESE    = (len("中国")=2)
	If WINNT_CHINESE Then
		Dim l,t,c
		Dim i
		l=len(str)
		t=l
		For i=1 To l
			c=asc(mid(str,i,1))
			If c<0 Then c=c+65536
			If c>255 Then
				t=t+1
			End If
		Next
		strLength=t
	Else 
		strLength=len(str)
	End If
	If err.number<>0 Then err.clear
End Function
%>





