<%
Dim Fs_User,ReturnValue,UserUrl,ThisRoot,ThisPage,QUERY_STRING
Set Fs_User = New Cls_User
Dim RsSysPara,UserDoMain,tmp_Url,tmp_Url_1
set RsSysPara = nothing
ReturnValue = Fs_User.checkStat(session("FS_UserName"),session("FS_UserPassword"))
ThisPage=Request.ServerVariables("SCRIPT_NAME")
QUERY_STRING = Request.ServerVariables("QUERY_STRING")
If QUERY_STRING<>"" Then
	ThisPage = ThisPage&"?"&QUERY_STRING
End If
If G_VIRTUAL_ROOT_DIR <>"" Then
	ThisRoot = "/" & G_VIRTUAL_ROOT_DIR & "/" & G_USER_DIR & "/"
Else
	ThisRoot = "/" & G_USER_DIR & "/"
End If
If ReturnValue=False Then
		Response.Redirect ThisRoot&"Login.asp?forward="&Server.URLEncode(ThisPage)
		Response.end
Else
	if Fs_User.OnlyLogin = 1 then
		 if session("UserLoginCode") <>Fs_User.UserLoginCode then
			Session("FS_UserName") = ""
			Session("FS_UserNumber") = ""
			Session("FS_UserPassword") = ""
			Session("FS_Group") = ""
			Session("FS_IsCorp") = ""
			Session("FS_NickName") = ""
			response.Cookies("FoosunUserCookies")("UserLogin_Style_Num")  = ""
			Session("UserLoginCode") = ""
			strShowErr = "<li>此帐户锁定,系统已经退出</li><li>有人在其他地方登陆您的此帐号!</li>"
			Response.Redirect ThisRoot& "lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl="&ThisRoot&"Login.asp"
			Response.end
		 End if
	End if
End If 
Sub getGroupIDinfo()
	set getRsGroup = User_Conn.execute("select GroupId,GroupNumber,GroupPoint,GroupDate,GroupMoney,UpfileNum,UpfileSize,LimitInfoNum,GroupDebateNum From FS_ME_Group where GroupId="&Fs_User.NumGroupID)
	if getRsGroup.eof then
		GroupNumber = ""
		GroupPoint = ""
		GroupDate = ""
		GroupMoney = ""
		UpfileNum = ""
		UpfileSize = ""
		LimitInfoNum = ""
		GroupDebateNum = ""
		getRsGroup.close:set getRsGroup = nothing
	else
		GroupNumber = getRsGroup("GroupNumber")
		GroupPoint = getRsGroup("GroupPoint")
		GroupDate = getRsGroup("GroupDate")
		GroupMoney = getRsGroup("GroupMoney")
		UpfileNum = getRsGroup("UpfileNum")
		UpfileSize = getRsGroup("UpfileSize")
		LimitInfoNum = getRsGroup("LimitInfoNum")
		GroupDebateNum = getRsGroup("GroupDebateNum")
		getRsGroup.close:set getRsGroup = nothing
	end if 
	'response.Write UpfileSize
	'response.end
end Sub
%>





