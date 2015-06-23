<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->

<%
'on error resume next
Dim Conn,User_Conn,InsertUserSysParaRs
Dim AllowReg,AllowChinese,NeedAudit,isCheckCorp,Login_Filed,OnlyMemberLogin,ID_Rule,IDRule_Array,ID_Elem,ID_Postfix,UserName_Length,Forbid_UserName,Pwd_Length,Pwd_Contain_Word,ResigterNeedFull,isSendMail,Email_Aduit,Reg_Help
Dim CheckCodeStyle,LoginStyle,ReturnUrl,ErrorPwdTimes,ShowNumberPerPage,UpfileType,UpfileSize,MessageSize,RssFeed,LimitClass,CertDir,LimitReviewChar,isPassCard,isYellowCheck
Dim DefaultGroupID,MoneyName,RegPointmoney,LoginPointmoney,PointChange,isPrompt,LenLoginTime,contrPoint,contrMoney,contrAuditPoint,contrAuditMoney,UserSystemName
MF_Default_Conn
MF_User_Conn
MF_Session_TF
Set InsertUserSysParaRs=Server.CreateObject(G_FS_RS)
'***************************************
if NoSqlHack(Request.QueryString("Act"))="BaseParam" then
	DefaultGroupID = NoSqlHack(Request.Form("DefaultGroupID"))
	AllowReg=NoSqlHack(Request.Form("AllowReg"))
	AllowChinese=NoSqlHack(Request.Form("AllowChinese"))
	NeedAudit=NoSqlHack(Request.Form("NeedAudit"))
	isCheckCorp=NoSqlHack(Request.Form("Corp_NeedAudit"))
	Login_Filed=Request.Form("Login_Filed")
	OnlyMemberLogin=NoSqlHack(Request.Form("OnlyMemberLogin"))
	ID_Rule=trim(Request.Form("ID_Prefix"))&","&NoSqlHack(replace(Request.Form("ID_Elem"),",","")&","&NoSqlHack(Request.Form("ID_Postfix")))&","&NoSqlHack(request.Form("ID_Devide"))
	if len(trim(Request.Form("NeedWord")))>0 then
		ID_Rule=ID_Rule&","&NoSqlHack(Request.Form("NeedWord"))
	else
		ID_Rule=ID_Rule&","
	end if
	UserName_Length=NoSqlHack(Request.Form("UserName_Length_Min"))&","&NoSqlHack(Request.Form("UserName_Length_Max"))
	Forbid_UserName=NoSqlHack(Request.Form("Forbid_UserName"))
	If Forbid_UserName = "" Or Isnull(Forbid_UserName) Then
	    Response.Redirect("../error.asp?ErrCodes=<li>“禁止注册的用户名”不能为空！</li>")
	    Response.End
	End if
	
	Pwd_Length=NoSqlHack(Request.Form("Pwd_Length_Min"))&","&NoSqlHack(Request.Form("Pwd_Length_Max"))
	isSendMail=NoSqlHack(Request.Form("isSendMail"))
	Email_Aduit=NoSqlHack(Request.Form("Email_Aduit"))
	Reg_Help=NoSqlHack(Request.Form("Reg_Help"))
	UserSystemName=NoSqlHack(Request.Form("UserSystemName"))
	'****************update DateBase***************************
	InsertUserSysParaRs.open "select RegisterTF,AllowChineseName,RegisterCheck,isCheckCorp,LoginStyle,OnlyMemberLogin,UserNumberRule,LenUserName,LimitUserName,LenPassword,isSendMail,isValidate,RegisterNotice,DefaultGroupID,UserSystemName from FS_ME_SysPara",User_Conn,1,3
	InsertUserSysParaRs("DefaultGroupID")=DefaultGroupID
	InsertUserSysParaRs("RegisterTF")=AllowReg
	InsertUserSysParaRs("AllowChineseName")=AllowChinese
	InsertUserSysParaRs("RegisterCheck")=NeedAudit
	InsertUserSysParaRs("isCheckCorp")=isCheckCorp
	InsertUserSysParaRs("LoginStyle")=Login_Filed
	InsertUserSysParaRs("OnlyMemberLogin")=OnlyMemberLogin
	InsertUserSysParaRs("UserNumberRule")=ID_Rule
	InsertUserSysParaRs("LenUserName")=UserName_Length
	InsertUserSysParaRs("LimitUserName")=Forbid_UserName
	InsertUserSysParaRs("LenPassword")=Pwd_Length
	InsertUserSysParaRs("isSendMail")=isSendMail
	InsertUserSysParaRs("isValidate")=Email_Aduit
	InsertUserSysParaRs("RegisterNotice")=Reg_Help
	InsertUserSysParaRs("UserSystemName")=UserSystemName
	InsertUserSysParaRs.update
	InsertUserSysParaRs.close
elseif NoSqlHack(Request.QueryString("Act"))="OtherParam" then
	LoginStyle=Request.Form("LoginStyle")
	ReturnUrl=Request.Form("ReturnUrl")
	ErrorPwdTimes=Request.Form("ErrorPwdTimes")
	ShowNumberPerPage=Request.Form("ShowNumberPerPage")
	UpfileType=Request.Form("UpfileType")
	UpfileSize=Request.Form("UpfileSize")
	MessageSize=Request.Form("MessageSize")
	RssFeed=Request.Form("RssFeed")
	LimitClass=trim(Request.Form("LimitClass"))&","&Trim(Request.Form("LimitClass2"))
	CertDir=Request.Form("CertDir")
	isPassCard=Request.Form("isPassCard")
	isYellowCheck=Request.Form("isYellowCheck")
	LimitReviewChar=Request.Form("LimitReviewChar")
	'*********************************************
	InsertUserSysParaRs.open "Select Login_Style,ReturnUrl,LoginLockNum,MemberList,UpfileType,UpfileSize,MessageSize,RssFeed,LimitClass,CertDir,LimitReviewChar,isPassCard,isYellowCheck,ReviewTF,FilesSpace From FS_ME_SysPara",User_Conn,1,3
	InsertUserSysParaRs("Login_Style")=LoginStyle
	InsertUserSysParaRs("ReturnUrl")=ReturnUrl
	InsertUserSysParaRs("LoginLockNum")=ErrorPwdTimes
	InsertUserSysParaRs("MemberList")=ShowNumberPerPage
	InsertUserSysParaRs("UpfileType")=UpfileType
	InsertUserSysParaRs("UpfileSize")=UpfileSize
	InsertUserSysParaRs("MessageSize")=MessageSize
	InsertUserSysParaRs("RssFeed")=RssFeed
	InsertUserSysParaRs("LimitClass")=LimitClass
	InsertUserSysParaRs("CertDir")=CertDir
	InsertUserSysParaRs("LimitReviewChar")=LimitReviewChar
	if Request.Form("ReviewTF")="1" then
		InsertUserSysParaRs("ReviewTF")=1
	else
		InsertUserSysParaRs("ReviewTF")=0
	end if
	if Request.Form("FilesSpace")<>"" then
		InsertUserSysParaRs("FilesSpace")=cint(Request.Form("FilesSpace"))
	else
		InsertUserSysParaRs("FilesSpace")=2
	end if
	InsertUserSysParaRs("isYellowCheck")=isYellowCheck
	InsertUserSysParaRs("isPassCard")=isPassCard
	InsertUserSysParaRs.update
	InsertUserSysParaRs.close
elseif Request.QueryString("Act")="MoneyParam" then 
	MoneyName=NoSqlHack(Request.Form("MoneyName"))
	RegPointmoney=NoSqlHack(Request.Form("Reg_Point"))&","&NoSqlHack(Request.Form("Reg_Money"))
	LoginPointmoney=NoSqlHack(Request.Form("Login_Point"))&","&NoSqlHack(Request.Form("Login_Money"))
	PointChange=NoSqlHack(Request.Form("PointChange"))&","&NoSqlHack(Request.Form("Money_to_Point"))&","&NoSqlHack(Request.Form("Point_to_Money"))
	isPrompt=NoSqlHack(Request.Form("isPrompt"))&","&NoSqlHack(Request.Form("PromptCondition"))
	LenLoginTime=NoSqlHack(Request.Form("LenLoginTime"))&","&NoSqlHack(Request.Form("LenLoginTime2"))
	contrPoint=NoSqlHack(Request.Form("txt_contrPoint"))
	contrMoney=NoSqlHack(Request.Form("txt_contrMoney"))
	contrAuditPoint=NoSqlHack(Request.Form("txt_contrAuditPoint"))
	contrAuditMoney=NoSqlHack(Request.Form("txt_contrAuditMoney"))
	'*************************************************
	InsertUserSysParaRs.open "select MoneyName,RegPointmoney,LoginPointmoney,PointChange,isPrompt,LenLoginTime,contrPoint,contrMoney,contrAuditPoint,contrAuditMoney from FS_ME_SysPara",User_Conn,1,3
	InsertUserSysParaRs("MoneyName")=MoneyName
	InsertUserSysParaRs("RegPointmoney")=RegPointmoney
	InsertUserSysParaRs("LoginPointmoney")=LoginPointmoney
	InsertUserSysParaRs("PointChange")=PointChange
	InsertUserSysParaRs("isPrompt")=isPrompt
	InsertUserSysParaRs("LenLoginTime")=LenLoginTime
	InsertUserSysParaRs("MoneyName")=MoneyName
	InsertUserSysParaRs("RegPointmoney")=RegPointmoney
	InsertUserSysParaRs("contrPoint")=contrPoint
	InsertUserSysParaRs("contrMoney")=contrMoney
	InsertUserSysParaRs("contrAuditPoint")=contrAuditPoint
	InsertUserSysParaRs("contrAuditMoney")=contrAuditMoney
	InsertUserSysParaRs.update
	InsertUserSysParaRs.close
End if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>FoosunCMS</title>
</head>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<body>
<%
if err.number>0 then
	Response.Redirect("../error.asp?ErrCodes="&err.description&"&ErrorUrl="&Request.ServerVariables("HTTP_REFERER"))
	Response.End()
else
	Response.Redirect("../success.asp?ErrCodes=<li>操作成功</li>&ErrorUrl="&Request.ServerVariables("HTTP_REFERER"))
	Response.End()
end if
%>
</body>
<%
	Conn.close
	Set Conn=nothing
	User_Conn.Close
	Set User_Conn=nothing
%>
</html>






