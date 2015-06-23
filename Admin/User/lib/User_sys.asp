<%
Sub User_GetParm()
	set GetUserParaObj  = server.CreateObject(G_FS_RS)
	UserSql = "select Top 1 Login_Style,UserNumberRule,RegisterNotice,RegisterTF,LoginLockNum,LoginPointmoney,Login_Style,AllowChineseName,LenUserName,isCheckCorp,RegisterCheck,MoneyName,RegPointmoney,ReturnUrl,isPrompt,LenPassword,isSendMail,isValidate,OnlyMemberLogin,LenLoginTime,LimitUserName,isPassCard,isYellowCheck,PointChange From FS_ME_SysPara"
	GetUserParaObj.open UserSql,User_Conn,1,3
	if Not GetUserParaObj.eof then
		if GetUserParaObj("RegisterTF")=0 then
			RegisterTF=false
		Else
			RegisterTF=true
		End if
		RegisterNotice = GetUserParaObj("RegisterNotice")
		UserDir = G_USERFILES_DIR
		p_UserNumberRule = GetUserParaObj("UserNumberRule")
		LoginLockNum = cint(GetUserParaObj("LoginLockNum"))
		p_LoginStyle = GetUserParaObj("Login_Style")
		p_AllowChineseName = GetUserParaObj("AllowChineseName")
		p_LenUserName = GetUserParaObj("LenUserName")
		strLenUserNameArr=split(p_LenUserName,",")
		p_NumLenMin = strLenUserNameArr(0)
		p_NumLenMax = strLenUserNameArr(1)
		p_isCheckCorp = cint(GetUserParaObj("isCheckCorp"))
		p_RegisterCheck =  cint(GetUserParaObj("RegisterCheck"))
		p_MoneyName =  GetUserParaObj("MoneyName")
		p_RegPointmoney = GetUserParaObj("RegPointmoney")
		strRegPointmoneyArr=split(p_RegPointmoney,",")
		p_NumGetPoint = trim(strRegPointmoneyArr(0))
		p_NumGetMoney =  trim(strRegPointmoneyArr(1))
		p_NumReturnUrl =  Cint(GetUserParaObj("ReturnUrl"))
		p_isPromptarr = split(GetUserParaObj("isPrompt"),",")
		p_strisPromptTF= cint(trim(p_isPromptarr(0)))
		p_strisPromptnNum=clng(trim( p_isPromptarr(1)))
		p_LenPasswordarr =  split(GetUserParaObj("LenPassword"),",")
		p_LenPassworMin = cint(trim(p_LenPasswordarr(0)))
		p_LenPassworMax = cint(trim(p_LenPasswordarr(1)))
		p_isSendMail = cint(GetUserParaObj("isSendMail"))
		p_isValidate =  cint(GetUserParaObj("isValidate"))
		p_OnlyMemberLogin = cint(GetUserParaObj("OnlyMemberLogin"))
		p_LenLoginTimearr =  split(GetUserParaObj("LenLoginTime"),",")
		p_LoginGetIntegral = Clng(p_LenLoginTimearr(0))
		p_LoginGetMoney = Cint(p_LenLoginTimearr(1))
		p_LoginPointmoneyarr =  split(GetUserParaObj("LoginPointmoney"),",")
		p_LoginPointmoneyarr_1 = Clng(p_LoginPointmoneyarr(0))
		p_LoginPointmoneyarr_2 = Cint(p_LoginPointmoneyarr(1))
		p_LimitUserName = GetUserParaObj("LimitUserName")
		p_isPassCard = GetUserParaObj("isPassCard")
		p_PointChange =  GetUserParaObj("PointChange")
		p_isYellowCheck = GetUserParaObj("isYellowCheck")
		GetUserParaObj.close:set GetUserParaObj = nothing
	Else
		strShowErr = "<li>参数错误</li><li>找不到会员系统配置</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End if
	
	Set GetConfigObj = server.CreateObject(G_FS_RS)
	GetConfigObj.open "select  Top  1 MF_Login_style,MF_Soft_Version From FS_MF_Config",Conn,1,3
	if Not GetConfigObj.eof then
		p_Soft_Version = "版本号:V" & GetConfigObj(1)
	Else
		p_Soft_Version = "<font color=""Red"">Err:Please configure Your Soft</font>"
	End if
	GetConfigObj.close
	Set GetConfigObj = nothing
	
	'获得天数字符串
	if Len(Cstr(month(now)))<2  then  
		strCard_month_1 = "0"&month(now)
	Else
		strCard_month_1 = month(now)
	End  if 
	if Len(Cstr(day(now)))<2  then  
		strCard_day_1 = "0"&day(now)
	Else
		strCard_day_1 = day(now)
	End  if 
	if Len(Cstr(hour(now)))<2  then  
		strCard_hour_1 = "0"&hour(now)
	Else
		strCard_hour_1 = hour(now)
	End  if 
	if Len(Cstr(minute(now)))<2  then  
		strCard_minute_1 = "0"&minute(now)
	Else
		strCard_minute_1 = minute(now)
	End  if 
	strTodaydate = right(year(now),2)&strCard_month_1&strCard_day_1
	strTodaydate_1 = right(year(now),2)&strCard_month_1&strCard_day_1&strCard_hour_1&strCard_minute_1
End Sub

Function GetRandstr()
	 Dim Randchar,Randchararr,RandLen,iR,Randomizecode
	Randomize
		Randchar="0,1,2,3,4,5,6,7,8,9,a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z"
		Randchararr=split(Randchar,",") 
		RandLen=8 '定义密码的长度或者是位数
	for iR=1 to RandLen
		Randomizecode=Randomizecode&Randchararr(Int((21*Rnd)))
	next 
End Function   
 
Function  ReturnError(Str,strURL) 
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(Str)&"&ErrorUrl="& strURL &"")
	Response.end
End Function

%>





