<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Md5.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="../API/Cls_PassportApi.asp" -->
<%
Dim p_RsLoginObj,p_RsLogObj
Dim p_UserName,p_UserPassword,p_Url,p_LoginLockNum,p_Logintype,p_vercode,p_sUrl,ReturnValue
Dim Fs_User,dzEmail,dz_UserPassword
Dim forward
forward =""
forward = Request.QueryString("forward")
User_GetParm
p_UserName = NoSqlHack(Replace(Replace(Trim(Request.Form("Name")),"'","''"),Chr(39),""))
p_UserPassword = MD5(Request.Form("password"),16)
dz_UserPassword = MD5(Request.Form("password"),32)
p_Logintype = NoSqlHack(Request.Form("Logintype"))
If p_Logintype="" Then
	p_Logintype = "0"
End If
p_vercode = NoSqlHack(lcase(Replace(Trim(Request("vercode")),"'","")))
Set Fs_User = New Cls_User
Dim CheckPost
Fs_User.CheckPostinput()
If CheckPost = False Then
	strShowErr = "<li>参数错误</li><li> 不要从外部提交数据</li>"
	Call ReturnError(strShowErr,"../login.asp")
End If
ReturnValue = Fs_User.Login(p_UserName,p_UserPassword,p_Logintype,p_vercode)
dzEmail = session("FS_UserEmail")
'-----------------------------------------------------------------
'系统整合
'-----------------------------------------------------------------
Dim API_Obj,API_SaveCookie,SysKey
If API_Enable Then
	Set API_Obj = New PassportApi
		'API_Obj.NodeValue "syskey",SysKey,0,False
		API_Obj.NodeValue "action","login",0,False
		API_Obj.NodeValue "username",p_UserName,1,False
		SysKey = Md5(API_Obj.XmlNode("username")&API_SysKey,16)
		API_Obj.NodeValue "syskey",SysKey,0,False
		API_Obj.NodeValue "savecookie","0",0,False
		API_Obj.NodeValue "password",NoSqlHack(Request.Form("password")),0,False
		API_Obj.NodeValue "userip","",0,False
		
		API_Obj.SendHttpData
		If API_Obj.Status = "1" Then
			Call ReturnError(API_Obj.Message,"../login.asp")
			Response.End()
		Else
			API_SaveCookie = API_Obj.SetCookie(SysKey,p_UserName,p_UserPassword,0)
			Response.Write API_SaveCookie
			Response.Flush
		End If
	Set API_Obj = Nothing
End If
'-----------------------------------------------------------------
Set Fs_User = Nothing
Dim ThisPage
ThisPage=NoSqlHack(Request.ServerVariables("SCRIPT_NAME"))
If forward="" Then
	forward = left(ThisPage,InStrRev(ThisPage,"/"))&"Main.asp"
End If
If ReturnValue = true Then 
	Response.Write "<script language=""JavaScript"">window.location.href="""&forward&""";</script>"
	Response.End
Else
	strShowErr = "<li>错误的的登陆用户名,或用户编号,或电子邮件及密码</li>"
	p_Url = "lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl="&forward
	Response.Write "<script language=""JavaScript"">window.location.href="""&p_Url&""";</script>"
End If
Set Fs_User = Nothing
Set Conn = Nothing
Set User_Conn = Nothing
%>





