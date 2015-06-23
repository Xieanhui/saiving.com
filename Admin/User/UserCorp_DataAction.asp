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
	server_v1 = CStr(Request.ServerVariables("HTTP_REFERER"))
	server_v2 = CStr(Request.ServerVariables("SERVER_NAME"))
	If Mid(server_v1, 8, Len(server_v2)) = server_v2 Then
		CheckPostinput = True
	End If
End Function
If CheckPostinput = False Then
	Response.Redirect("lib/Error.asp?ErrCodes=<li>参数错误</li><li> 不要从外部提交数据</li>&ErrorUrl=../UserCorp.asp")
	Response.end
End If

Function CheckCF(FildName,FildValue,Str_LinkStr)
	'判断重复 Str_LinkStr = ' # ""
	CheckCF = User_Conn.execute("select count(*) from FS_ME_Users where "&FildName&"="&Str_LinkStr& FildValue &Str_LinkStr)(0)
	if err.number>0 then
		Response.Redirect("lib/error.asp?ErrCodes="&server.URLEncode(err.description))
		Response.End()
	end if	
End Function

Function GetPwdByUserNumber(UserNumber)
	if UserNumber<>"" then 
		GetPwdByUserNumber = User_Conn.execute("select UserPassword from FS_ME_Users where UserNumber = '"&UserNumber&"'")(0)
	else
		GetPwdByUserNumber = ""
	end if
	if err.number>0 then
		err.clear : GetPwdByUserNumber = ""
	end if	
End Function

Str_BaseData_List = "UserName,UserPassword,PassQuestion,PassAnswer,SafeCode,Email"
Str_OtherData_List = "NickName,RealName,Sex,BothYear,Certificate,CerTificateCode,Province,City" _
	&",HeadPic,HeadPicSize,tel,Mobile,isMessage,HomePage,QQ,MSN,Address,PostCode,Vocation,Integral,FS_Money" _
	&",TempLastLoginTime,TempLastLoginTime_1,CloseTime,IsMarray,SelfIntro,isOpen,GroupID,isLock,UserFavor,OnlyLogin"

select case Request.QueryString("Act")
	case "BaseData"
		UserNumber = NoSqlHack(trim(request.Form("frm_UserNumber_Edit1")))
		call save(Str_BaseData_List,1,0)
	case "OtherData"
		UserNumber = NoSqlHack(trim(request.Form("frm_UserNumber_Edit2")))
		if UserNumber="" then 
			Response.Redirect("lib/error.asp?ErrCodes=<li>修改时必要参数必须填写。</li>")
			response.End()
		end if
		call save(Str_OtherData_List,0,1)
	case "ThreeData"
		UserNumber = NoSqlHack(trim(request.Form("frm_UserNumber_Edit3")))
		if UserNumber="" then 
			Response.Redirect("lib/error.asp?ErrCodes=<li>修改时必要参数必须填写。</li>")
			response.End()
		end if
		Call SaveOtherData(UserNumber)
	case "Add_AllData"
		UserNumber = NoSqlHack(trim(request.Form("frm_UserNumber_Edit3")))	
		call save(Str_BaseData_List &","& Str_OtherData_List,1,2)
	case "Del"	
		Del
end select 


Sub Del()
	Dim Str_Tmp,Arr_Tmp,s_StrPWD_,strAllUserName
	if request.QueryString("UserNumber")<>"" then 
		Str_Tmp = NoSqlHack(Trim(request.QueryString("UserNumber")))
	else
		Str_Tmp = request.Form("frm_UserNumber")
	end if
	if Str_Tmp="" then Response.Redirect("lib/error.asp?ErrCodes=<li>你必须至少选择一个进行删除。</li>")
	Str_Tmp = replace(Str_Tmp," ","")
	Arr_Tmp = split(Str_Tmp,",")
	strShowErr = ""
	'on error resume next
	'-----------------------------------------------------------------
	'系统整合
	'-----------------------------------------------------------------
	Dim API_Obj,API_SaveCookie,SysKey
	If API_Enable Then
		Dim AllDelName
		AllDelName = request.Form("frm_UserName")
		IF AllDelName = "" Then Response.Redirect("lib/error.asp?ErrCodes=<li>你必须至少选择一个进行删除。</li>")
		AllDelName = Replace(AllDelName," ","")
		strAllUserName = AllDelName
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
		s_StrPWD_ = GetPwdByUserNumber(Str_Tmp)
		if s_StrPWD_<>"" then 
			Call Fs_User.DelUser(Str_Tmp,s_StrPWD_)
		else
			strShowErr = strShowErr & "<li>用户"&Str_Tmp&"未删除，可能该用户已不存在……。</li>"
		end if	
	next
	if strShowErr<>"" then strShowErr = "<li>以下是删除失败的描述：</li>"&strShowErr
	Response.Redirect("lib/Success.asp?ErrorUrl=../UserCorp.asp&ErrCodes=<li>恭喜，删除成功。</li>"&strShowErr)
End Sub

Sub Save(Str_Tmp,Bit_IsNull,Action)
	Dim Arr_Tmp,UserSql
	Arr_Tmp = split(Str_Tmp,",")
	UserSql = "select UserNumber,IsCorporation, "&Str_Tmp&" from FS_ME_Users where UserNumber= '"&NoSqlHack(UserNumber)&"'"
	Set UpdateUserRs = CreateObject(G_FS_RS)
	UpdateUserRs.Open UserSql,User_Conn,3,3
	if UserNumber<>"" and not UpdateUserRs.eof then 
	''修改
		UpdateUserRs("IsCorporation") = 1
		for each Str_Tmp in Arr_Tmp
			if Bit_IsNull = 1 then 
				if request.Form("frm_"&Str_Tmp)<>"" then 
					if instr(",UserPassword,PassQuestion,PassAnswer,SafeCode,",","&Str_Tmp&",")>0 then 
						UpdateUserRs(Str_Tmp) = Md5(request.Form("frm_"&Str_Tmp),16)
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
			end if	
		next	
		UpdateUserRs.update
		UpdateUserRs.close
		if err.number>0 then
			strShowErr = "<li>基础设置未修改成功。</li>"
			Response.Redirect("lib/error.asp?ErrCodes="&strShowErr)
			Response.End()
		else
			if Action > 1 then 
				Call SaveOtherData(UserNumber) ''保存公司特有信息。
			end if
			Response.Redirect("lib/success.asp?ErrCodes=<li>恭喜，修改成功。</li>&ErrorUrl="&server.URLEncode("../UserCorp.asp?Act=View&Add_Sql=A.UserNumber='"&UserNumber&"'"))
			Response.End()
		end if
	else
	''新增
		strUserNumberRule= Fs_User.strUserNumberRule(p_UserNumberRule)
		if CheckCF("UserNumber",strUserNumberRule,"'")>0 then 
			Response.Redirect("lib/Error.asp?ErrCodes=<li>用户编号以外重复。请重新提交。</li>")
			Response.end
		end if
		if CheckCF("UserName",NoSqlHack(request.Form("frm_UserName")),"'")>0 then 
			Response.Redirect("lib/Error.asp?ErrCodes=<li>用户名重复。请重新提交。</li>")
			Response.end
		end if
		if CheckCF("Email",NoSqlHack(request.Form("frm_Email")),"'")>0 then 
			Response.Redirect("lib/Error.asp?ErrCodes=<li>用户名重复。请重新提交。</li>")
			Response.end
		end if
			
		UpdateUserRs.addnew

		UpdateUserRs("UserNumber") = strUserNumberRule
		UpdateUserRs("IsCorporation") = 1
		UpdateUserRs("RegTime") = now

		for each Str_Tmp in Arr_Tmp
			if Bit_IsNull = 1 then 
				if request.Form("frm_"&Str_Tmp)<>"" then 
					if instr(",UserPassword,PassQuestion,PassAnswer,SafeCode,",","&Str_Tmp&",")>0 then 
						UpdateUserRs(Str_Tmp) = Md5(request.Form("frm_"&Str_Tmp),16)
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
			end if	
		next
		UpdateUserRs.update
		'response.End()	
		UpdateUserRs.close
		if err.number>0 then
			strShowErr = "<li>基础设置未添加成功。</li>"
			Response.Redirect("lib/error.asp?ErrCodes="&strShowErr)
			Response.End()
		else
			if Action > 1 then 
				Call SaveOtherData(UserNumber) ''保存公司特有信息。
		
				'插入会员参数
				call Fs_User.InsertMyPara( strUserNumberRule )
				'插入日志
				call Fs_User.AddLog("注册",strUserNumberRule,p_NumGetPoint,p_NumGetMoney,"注册获得积分",0)
				'给会员发送电子邮件 
				Dim str_isSendMail
				str_isSendMail=false
			
			end if
			Response.Redirect("lib/success.asp?ErrCodes=<li>恭喜，新增成功。</li>&ErrorUrl="&server.URLEncode("../UserCorp.asp?Act=View&Add_Sql=A.UserNumber='"&strUserNumberRule&"'"))
			Response.End()
		end if
	end if
End Sub
Sub SaveOtherData(UserNumber)
''保存公司独有信息。
		Dim AddCorpDataObj,Str_Tmp_,Arr_Tmp_
		Set AddCorpDataObj = server.CreateObject(G_FS_RS)
		AddCorpDataObj.open "select  * From FS_ME_CorpUser where UserNumber='"&NoSqlHack(UserNumber)&"'",User_Conn,1,3
		if AddCorpDataObj.eof then 
			strShowErr = "<li>用户"&UserNumber&"在公司用户表中不存在……。</li>"
			Response.Redirect("lib/error.asp?ErrCodes="&strShowErr)
			Response.End()
		end if
		Str_Tmp_="C_Name,C_ShortName,C_Province,C_City,C_Address,C_PostCode,C_ConactName,C_Tel,isLockCorp,C_Fax,C_VocationClassID,C_Website,C_size,C_Capital,C_BankName,C_BankUserName"
		Arr_Tmp_ = split(Str_Tmp_,",")
		if UserNumber="" then AddCorpDataObj.addNew
		AddCorpDataObj("UserNumber") = UserNumber
		for each Str_Tmp_ in Arr_Tmp_
			if NoSqlHack(request.Form("frm_"&Str_Tmp_))<>"" then 
				AddCorpDataObj(Str_Tmp_) = NoSqlHack(request.Form("frm_"&Str_Tmp_))
			else
				AddCorpDataObj(Str_Tmp_) = null
			end if	
			'response.Write(Str_Tmp_&":"&NoSqlHack(request.Form("frm_"&Str_Tmp_))&"<br>")
		next	
		AddCorpDataObj("isYellowPage") = 0 
		AddCorpDataObj("isYellowPageCheck") = 0 
		AddCorpDataObj.update
		AddCorpDataObj.close
		set AddCorpDataObj = nothing
		if err.number>0 then
			'回滚
			s_StrPWD_ = GetPwdByUserNumber(UserNumber)
			if s_StrPWD_<>"" then 
				Call Fs_User.DelUser(UserNumber,s_StrPWD_)
				strShowErr = strShowErr & "<li>保存公司扩展信息时出错，用户"&UserNumber&"基本信息已删除。</li>"
			else
				strShowErr = strShowErr & "<li>用户"&UserNumber&"未删除，可能该用户已不存在……。</li>"
			end if	
			Response.Redirect("lib/error.asp?ErrCodes="&strShowErr)
			Response.End()
		else
			Response.Redirect("lib/success.asp?ErrCodes=<li>恭喜，修改成功。</li>&ErrorUrl="&server.URLEncode("../UserCorp.asp?Act=View&Add_Sql=A.UserNumber='"&UserNumber&"'"))
			Response.End()
		end if
End Sub
''=========================================================
User_Conn.Close
Set User_Conn=nothing
%>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. --> 





