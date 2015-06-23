<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="../../FS_Inc/Md5.asp" -->
<!--#include file="../../API/Cls_PassportApi.asp" -->
<%'Copyright (c) 2006 Foosun Inc. 
'on error resume next
Dim UserNumber
Dim Str_BaseData_List,Str_OtherData_List,strUserNumberRule 
Dim Fs_User
MF_Default_Conn        
MF_User_Conn     
MF_Session_TF
Set Fs_User = New Cls_User
'***************************************
Function CheckPostinput()
	Dim server_v1, server_v2
	CheckPostinput = False
	server_v1 = CStr(NoSqlHack(Request.ServerVariables("HTTP_REFERER")))
	server_v2 = CStr(NoSqlHack(Request.ServerVariables("SERVER_NAME")))
	If Mid(server_v1, 8, Len(server_v2)) = server_v2 Then
		CheckPostinput = True
	End If
End Function
Function CheckCF(FildName,FildValue,Str_LinkStr)
	'判断重复 Str_LinkStr = ' # ""
	CheckCF = User_Conn.execute("select count(*) from FS_ME_Users where "&FildName&"="&Str_LinkStr& FildValue &Str_LinkStr)(0)
	if err.number<>0 then
		Response.Redirect("lib/error.asp?ErrCodes="&server.URLEncode(err.description))
		Response.End()
	end if	
End Function
If CheckPostinput = False Then
	Response.Redirect("lib/Error.asp?ErrCodes=<li>参数错误</li><li> 不要从外部提交数据</li>&ErrorUrl=../User_manage.asp")
	Response.end
End If

Function GetPwdByUserName(UserName)
	if UserName<>"" then 
		GetPwdByUserName = User_Conn.execute("select UserPassword from FS_ME_Users where UserName = '"&UserName&"'")(0)
	else
		GetPwdByUserName = ""
	end if
	if err.number<>0 then
		err.clear : GetPwdByUserName = ""
	end if	
End Function

Str_BaseData_List = "UserName,UserPassword,PassQuestion,PassAnswer,SafeCode,Email"
Str_OtherData_List = "NickName,RealName,Sex,BothYear,Certificate,CerTificateCode,Province,City" _
	&",HeadPic,HeadPicSize,tel,Mobile,isMessage,HomePage,QQ,MSN,Address,PostCode,Vocation,Integral,FS_Money" _
	&",TempLastLoginTime,TempLastLoginTime_1,CloseTime,IsMarray,SelfIntro,isOpen,GroupID,isLock,UserFavor,OnlyLogin"

select case Request.QueryString("Act")
	case "BaseData"
		UserNumber = NoSqlHack(request.Form("frm_UserNumber_Edit1"))
		call save(Str_BaseData_List,1)
	case "OtherData"
		UserNumber = NoSqlHack(request.Form("frm_UserNumber_Edit2"))
		if UserNumber="" then 
			Response.Redirect("lib/error.asp?ErrCodes=<li>修改时必要参数必须填写。</li>")
			response.End()
		end if
		call save(Str_OtherData_List,0)
	case "Add_AllData"
		UserNumber = NoSqlHack(request.Form("frm_UserNumber_Edit2"))	
		call save(Str_BaseData_List &","& Str_OtherData_List,1)
	case "Del"	
		Del
end select 


Sub Del()
	Dim Str_Tmp,Arr_Tmp,s_StrPWD_,strAllUserName
	if request.QueryString("UserName")<>"" then 
		Str_Tmp = NoSqlHack(request.QueryString("UserNumber"))
	else
		Str_Tmp = NoSqlHack(request.Form("frm_UserName"))
	end if
	if Str_Tmp="" then Response.Redirect("lib/error.asp?ErrCodes=<li>你必须至少选择一个进行删除。</li>")
	Str_Tmp = replace(Str_Tmp," ","")
	strAllUserName = Str_Tmp
	Arr_Tmp = split(Str_Tmp,",")
	strShowErr = ""
	'on error resume next
	'REsponse.write strAllUserName
'	Response.end
	'-----------------------------------------------------------------
	'系统整合
	'-----------------------------------------------------------------
	Dim API_Obj,API_SaveCookie,SysKey
	If API_Enable Then
		Set API_Obj = New PassportApi
			API_Obj.NodeValue "action","delete",0,False
			API_Obj.NodeValue "username",strAllUserName,1,False
			SysKey = Md5(API_Obj.XmlNode("username")&API_SysKey,16)
			API_Obj.NodeValue "syskey",SysKey,0,False
			API_Obj.SendHttpData
			If API_Obj.Status = "1" Then
				Response.redirect "showerr.asp?ErrCodes="& API_Obj.Message &"&action=OtherErr"
			End If
		Set API_Obj = Nothing
	End If
	'-----------------------------------------------------------------
	for each Str_Tmp in Arr_Tmp
		s_StrPWD_ = GetPwdByUsername(Str_Tmp)
		if s_StrPWD_<>"" then 
			Call Fs_User.DelUser(Str_Tmp,s_StrPWD_)
		else
			strShowErr = strShowErr & "<li>用户"&Str_Tmp&"未删除，可能该用户已不存在……。</li>"
		end if	
	next
	
	if strShowErr<>"" then strShowErr = "<li>以下是删除失败的描述：</li>"&strShowErr
	Response.Redirect("lib/Success.asp?ErrorUrl=../User_manage.asp&ErrCodes=<li>恭喜，删除成功。</li>"&strShowErr)
End Sub

Sub Save(Str_Tmp,Bit_IsNull)
	Dim Arr_Tmp,UserSql
	Arr_Tmp = split(Str_Tmp,",")
	UserSql = "select UserName,UserNumber,IsCorporation,RegTime,myskin, "&Str_Tmp&" from FS_ME_Users where UserNumber= '"&NoSqlHack(UserNumber)&"'"
	Set UpdateUserRs = CreateObject(G_FS_RS)
	UpdateUserRs.Open UserSql,User_Conn,1,3
	if UserNumber<>"" and not UpdateUserRs.eof then 
	''修改  
		UpdateUserRs("IsCorporation") = 0
		'UserName,UserPassword,PassQuestion,PassAnswer,SafeCode,Email
		'-----------------------------------------------------------------
		'系统整合
		'-----------------------------------------------------------------
		Dim API_Obj,API_SaveCookie,SysKey
		If API_Enable Then
			Set API_Obj = New PassportApi
				API_Obj.NodeValue "action","update",0,False
				API_Obj.NodeValue "username",UpdateUserRs("UserName"),1,False
				API_Obj.NodeValue "email",NoSqlHack(Request.Form("frm_Email")),1,False
				API_Obj.NodeValue "question",NoSqlHack(Request.Form("frm_PassQuestion")),1,False
				API_Obj.NodeValue "answer",NoSqlHack(Request.Form("frm_PassAnswer")),1,False
				SysKey = Md5(API_Obj.XmlNode("username")&API_SysKey,16)
				API_Obj.NodeValue "syskey",SysKey,0,False
				API_Obj.NodeValue "password",NoSqlHack(Request.Form("frm_UserPassword")),0,False
				API_Obj.SendHttpData
				If API_Obj.Status = "1" Then
					Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(API_Obj.Message)&"&ErrorUrl=")
					Response.end
				End If
			Set API_Obj = Nothing
		End If
		'-----------------------------------------------------------------
		for each Str_Tmp in Arr_Tmp
			if Bit_IsNull = 1 then 
				if request.Form("frm_"&Str_Tmp)<>"" then 
					if instr(",UserPassword,PassAnswer,SafeCode,",","&Str_Tmp&",")>0 then 
						UpdateUserRs(Str_Tmp) = Md5(NoSqlHack(request.Form("frm_"&Str_Tmp)),16)
					else
						if NoSqlHack(request.Form("frm_"&Str_Tmp))<>"" then 
							UpdateUserRs(Str_Tmp) = NoSqlHack(request.Form("frm_"&Str_Tmp))
						else
							UpdateUserRs(Str_Tmp) = null
						end if	
					end if		
				end if
			else
				if NoSqlHack(request.Form("frm_"&Str_Tmp))<>"" then 
					UpdateUserRs(Str_Tmp) = NoSqlHack(request.Form("frm_"&Str_Tmp))
				else
					UpdateUserRs(Str_Tmp) = null
				end if	
				'response.Write(Str_Tmp&":"&NoSqlHack(request.Form("frm_"&Str_Tmp))&"<br>")
			end if
		Next
		UpdateUserRs("myskin")=2
		'response.End()
		UpdateUserRs.update
		UpdateUserRs.close
		if err.number<>0 then
			Response.Redirect("lib/error.asp?ErrCodes="&server.URLEncode(err.description))
			Response.End()
		else
			Response.Redirect("lib/success.asp?ErrCodes=<li>恭喜，修改成功。</li>&ErrorUrl="&server.URLEncode("../User_manage.asp?Act=View&Add_Sql=UserNumber='"&UserNumber&"'"))
			Response.End()
		end if
	else
	''新增
		'strUserNumberRule= Fs_User.strUserNumberRule(p_UserNumberRule)
		strUserNumberRule = GetRamCode(10)
		if CheckCF("UserNumber",strUserNumberRule,"'")>0 then 
			Response.Redirect("lib/Error.asp?ErrCodes=<li>用户编号意外重复。请重新提交。</li>")
			Response.end
		end if
		if CheckCF("UserName",NoSqlHack(request.Form("frm_UserName")),"'")>0 then 
			Response.Redirect("lib/Error.asp?ErrCodes=<li>用户名重复。请重新提交。</li>")
			Response.end
		end if
		if CheckCF("Email",NoSqlHack(request.Form("frm_Email")),"'")>0 then 
			Response.Redirect("lib/Error.asp?ErrCodes=<li>Email重复。请重新提交。</li>")
			Response.end
		end if
			
		UpdateUserRs.addnew

		UpdateUserRs("UserNumber") = NoSqlHack(strUserNumberRule)
		UpdateUserRs("IsCorporation") = 0
		UpdateUserRs("RegTime") = now
		for each Str_Tmp in Arr_Tmp
			if Bit_IsNull = 1 then 
				if request.Form("frm_"&Str_Tmp)<>"" then 
					if instr(",UserPassword,PassQuestion,PassAnswer,SafeCode,",","&Str_Tmp&",")>0 then 
						UpdateUserRs(Str_Tmp) = Md5(NoSqlHack(request.Form("frm_"&Str_Tmp)),16)
					else
						if NoSqlHack(request.Form("frm_"&Str_Tmp))<>"" then 
							UpdateUserRs(Str_Tmp) = NoSqlHack(request.Form("frm_"&Str_Tmp))
						else
							UpdateUserRs(Str_Tmp) = null
						end if	
					end if		
					'response.Write(Str_Tmp&" : "&NoSqlHack(request.Form("frm_"&Str_Tmp))&"<br>" )
				end if	
			else
				if NoSqlHack(request.Form("frm_"&Str_Tmp))<>"" then 
					UpdateUserRs(Str_Tmp) = NoSqlHack(request.Form("frm_"&Str_Tmp))
				else
					UpdateUserRs(Str_Tmp) = null
				end if	
			end if	
		next
		'response.End()	
	    UpdateUserRs.update   
		UpdateUserRs.close
		if err.number<>0 then
			Response.Redirect("lib/error.asp?ErrCodes="&server.URLEncode(err.description))
			Response.End()
		else

	'插入会员参数
	call Fs_User.InsertMyPara( strUserNumberRule )
	'插入日志
	call Fs_User.AddLog("注册",strUserNumberRule,p_NumGetPoint,p_NumGetMoney,"注册获得积分",0)
	'给会员发送电子邮件 
	Dim str_isSendMail
	str_isSendMail=false
	'====================================整合动网、OBLOG 2006-11-11=========
	'If HaveOblog=1 Then Call ObAddUser(NoSqlHack(request.Form("frm_UserName")),NoSqlHack(request.Form("frm_UserPassword")),NoSqlHack(request.Form("frm_Email")),NoSqlHack(request.Form("frm_Sex")),NoSqlHack(request.Form("frm_PassQuestion")),NoSqlHack(request.Form("frm_PassAnswer")),NoSqlHack(request.Form("frm_NickName")),ObConn)
	'If HaveDvbbs=1 Then Call DvbbsAddUser(NoSqlHack(request.Form("frm_UserName")),NoSqlHack(request.Form("frm_UserPassword")),NoSqlHack(request.Form("frm_Email")),NoSqlHack(request.Form("frm_Sex")),NoSqlHack(request.Form("frm_PassQuestion")),md5(NoSqlHack(request.Form("frm_PassAnswer")),16),DvConn)
	'========================================================FS.awen ueuo.cn

			Response.Redirect("lib/success.asp?ErrCodes=<li>恭喜，新增成功。</li>&ErrorUrl="&server.URLEncode("../User_manage.asp?Act=View&Add_Sql=UserNumber='"&strUserNumberRule&"'"))
			Response.End()
		end if
	end if
End Sub
''=========================================================
User_Conn.Close
Set User_Conn=nothing
%>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. --> 





