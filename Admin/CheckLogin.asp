<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/md5.asp" -->
<%
Response.Cookies("FoosunMFCookies")=Empty
Response.Cookies("FoosunSUBCookie")=Empty
Dim Conn
Dim p_RsLoginObj,p_RsLogObj
Dim p_UserName,p_UserPass,p_VerifyCode,p_System,p_SqlLog,p_SqlLogin,p_Url,p_TempUserPass
Dim p_PassArr,p_TrueResult,p_CheckedResult,GetRamCode_adminCode
MF_Default_Conn
GetRamCode_adminCode = GetRamCode(8)
if Request.form("URLs")<>"" then:p_Url = "index.asp?Urls="&NoSqlHack(Request.form("URLs")):else:p_Url = "index.asp":end if
p_UserName = NoSqlHack(Request.Form("Name"))
p_TempUserPass = NoSqlHack(Request.Form("Password"))
p_VerifyCode = lcase(NoSqlHack(Request("VerifyCode")))
if p_UserName = "" or  p_TempUserPass = "" then
	Response.Write("<script>alert(""错误：\n请填写完整"");location.href=""Login.asp"";</script>")
	Response.End
end if

if cstr(Session("GetCode"))<>cstr(p_VerifyCode) then
	Response.Write("<script>alert(""错误：\n您输入的确认码和系统产生的不一致，请重新输入。\n返回后请刷新登录页面后重新输入正确的信息"");location.href=""Login.asp"";</script>")
	Response.End
end if

p_System = Request.ServerVariables("HTTP_USER_AGENT")
if Instr(p_System,"Windows NT 5.2") then
	p_System = "Win2003"
elseif Instr(p_System,"Windows NT 5.0") then
	p_System="Win2000"
elseif Instr(p_System,"Windows NT 5.1") then
	p_System = "WinXP"
elseif Instr(p_System,"Windows NT") then
	p_System = "WinNT"
elseif Instr(p_System,"Windows 9") then
	p_System = "Win9x"
elseif Instr(p_System,"unix") or instr(p_System,"linux") or instr(p_System,"SunOS") or instr(p_System,"BSD") then
	p_System = "类Unix"
elseif Instr(p_System,"Mac") then
	p_System = "Mac"
else
	p_System = "Other"
end if
p_PassArr=split(G_SAFE_PASS_SET_STR,",")

If p_PassArr(0)=1 then
	If p_PassArr(3)="1" then
		p_TrueResult=Trim(Cstr(Cint(mid(Session("GetCode"),Cint(p_PassArr(1)),1))+Cint(mid(Session("GetCode"),Cint(p_PassArr(2)),1))))
	Else
		p_TrueResult=Trim(Cstr(Cint(mid(Session("GetCode"),Cint(p_PassArr(1)),1))*Cint(mid(Session("GetCode"),Cint(p_PassArr(2)),1))))
	End If
	If p_PassArr(4)="0" then
		p_CheckedResult=left(p_TempUserPass,Len(p_TrueResult))
		p_UserPass=mid(p_TempUserPass,Len(p_TrueResult)+1)
	ElseIf Cint(p_PassArr(4))>len(p_TempUserPass)-len(p_TrueResult) then
		p_CheckedResult=right(p_TempUserPass,Len(p_TrueResult))
		p_UserPass=left(p_TempUserPass,len(p_TempUserPass)-Len(p_TrueResult))
	Else
		p_CheckedResult=mid(p_TempUserPass,p_PassArr(4)+1,Len(p_TrueResult))
		p_UserPass=left(p_TempUserPass,p_PassArr(4))&mid(p_TempUserPass,Cint(p_PassArr(4))+len(p_TrueResult)+1)
	End If
Else
	p_UserPass=p_TempUserPass
End If
Session("GetCode")=""

Set p_RsLoginObj = server.CreateObject (G_FS_RS)
p_SqlLogin = "select Admin_Name,Admin_Pass_Word,Admin_Is_Locked,Admin_Parent_Admin,Admin_Is_Super,Admin_Pop_List,Admin_Add_Admin,Admin_Style_Num,Admin_FilesTF from FS_MF_Admin where Admin_Name='"&p_UserName&"' and  Admin_Pass_Word='"&md5(p_UserPass,16)&"'"
p_RsLoginObj.Open p_SqlLogin,Conn,1,1
if Not p_RsLoginObj.EOF then
	if cint(p_RsLoginObj("Admin_Is_Locked")) = 1 then
		Response.Write("<script>alert(""错误:\n您已经被锁定,请与管理联系\n"");window.close();</script>")
		Response.End
	end if
	Session("Admin_Name") = p_RsLoginObj("Admin_Name")
	Session("Admin_Pass_Word") = p_RsLoginObj("Admin_Pass_Word")
	Session("Admin_Parent_Admin") = p_RsLoginObj("Admin_Parent_Admin")
	Session("Admin_Is_Super") = p_RsLoginObj("Admin_Is_Super")
	Session("Admin_Pop_List") = p_RsLoginObj("Admin_Pop_List")
	Session("Admin_Add_Admin") = p_RsLoginObj("Admin_Add_Admin")
	Session("Admin_Style_Num") = p_RsLoginObj("Admin_Style_Num")
	Session("Admin_FilesTF") = p_RsLoginObj("Admin_FilesTF")
	If Cint(p_RsLoginObj("Admin_Is_Super")) <> 1 Then
		Dim p_FSO,tmps_path,Temps_AdminPath
		Set p_FSO = Server.CreateObject(G_FS_FSO)
			Temps_AdminPath = "..\"& G_UP_FILES_DIR &"\adminFiles"
			if p_FSO.FolderExists(Server.MapPath(Temps_AdminPath)) = false then p_FSO.CreateFolder(Server.MapPath(Temps_AdminPath))
			tmps_path = Temps_AdminPath & "\" & UCase(md5(p_RsLoginObj("Admin_Name"),16))
			if p_FSO.FolderExists(Server.MapPath(tmps_path)) = false then p_FSO.CreateFolder(Server.MapPath(tmps_path))
		set p_FSO = nothing
	End If	
	'更新随机码
	Conn.execute("Update FS_MF_Admin Set admin_Code = '"& GetRamCode_adminCode &"' where Admin_Name='"& p_UserName &"' and Admin_Pass_Word='"& md5(p_UserPass,16) &"'")	
	session("fs_admin_Code")= GetRamCode_adminCode
	Response.Cookies("FoosunCookie")("LoginStyle")=p_RsLoginObj("Admin_Style_Num")
	Set p_RsLogObj = Server.Createobject(G_FS_RS)
	p_SqlLog = "Select Admin_Name,Log_IP,Log_OS_Sys,Log_TF,Log_Time from FS_MF_Login_Log where 1=0"
	p_RsLogObj.Open p_SqlLog,Conn,3,3
	p_RsLogObj.AddNew
	p_RsLogObj("Admin_Name") = p_UserName
	p_RsLogObj("Log_IP") = 	NoSqlHack(Request.ServerVariables("Remote_Addr"))
	p_RsLogObj("Log_OS_Sys") = p_System
	p_RsLogObj("Log_TF") = 1
	p_RsLogObj("Log_Time") = Now()
	p_RsLogObj.Update
	p_RsLogObj.Close:set p_RsLogObj = Nothing
	If CBool(Request.Form("AutoGet")) or NoSqlHack(Request.Form("AutoGet"))<>"" Then
        Response.Cookies("FoosunCookie")("AdminName")=Session("Admin_Name")
        Response.Cookies("FoosunCookie").Expires=Date()+365
    Else
        Response.Cookies("FoosunCookie")("AdminName")=""
        Response.Cookies("FoosunCookie").Expires=Date()-1
    End If
	SubSys_Cookies:MFConfig_Cookies:NSConfig_Cookies:DSConfig_Cookies:MSConfig_Cookies
	If p_TrueResult=p_CheckedResult then
		Session("GetCode")=empty
		Conn.Close:set Conn = Nothing
		Response.Redirect p_Url
		Response.End
	Else
		Conn.Close:set Conn = Nothing
		Response.Write("<script>alert(""错误:\n请检查用户名和密码的正确性\n"");location.href=""Login.asp"";</script>")
		Response.End
	End If
else
	Set p_RsLogObj = Server.Createobject(G_FS_RS)
	p_SqlLog = "Select ID,Admin_Name,Log_IP,Log_OS_Sys,Log_Error_Pass,Log_TF,Log_Time from FS_MF_Login_Log where 1=0"
	p_RsLogObj.open p_SqlLog,Conn,3,3
	p_RsLogObj.AddNew
	p_RsLogObj("Admin_Name") = NoSqlHack(Request.Form("Name"))
	p_RsLogObj("Log_IP") = NoSqlHack(Request.ServerVariables("Remote_Addr"))
	p_RsLogObj("Log_OS_Sys") = p_System
	p_RsLogObj("Log_Error_Pass") = NoSqlHack(Request.Form("Password"))
	p_RsLogObj("Log_TF") = 0
	p_RsLogObj("Log_Time") = Now()
	p_RsLogObj.Update
	p_RsLogObj.Close:Set p_RsLogObj = Nothing
	Response.Write("<script>alert(""错误:\n请检查用户名和密码的正确性\n"");location.href=""Login.asp"";</script>")
	Response.End
end if
Conn.Close:set Conn = Nothing
Set p_RsLoginObj = Nothing
%>