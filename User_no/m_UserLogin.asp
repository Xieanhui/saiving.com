<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Md5.asp" -->
<!--#include file="lib/strlib.asp" -->
<%
session.CodePage="936"
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.Charset="GB2312"
Server.ScriptTimeOut=9999999

User_GetParm

Dim Str_Type,Str_Spanid,Str_UserInfo,Str_UserLogin,Login_DisType,TypeStr
Dim Cookie_Domain,Str_Act,SysRoot
Dim ShowInfoStr
Dim Str_name,Str_Pass
Dim User_Path_Str

Cookie_Domain = Get_MF_Domain()
if Cookie_Domain="" then 
	Cookie_Domain = "localhost"
else
	if left(lcase(Cookie_Domain),len("http://"))="http://" then Cookie_Domain = mid(Cookie_Domain,len("http://")+1)
	if right(Cookie_Domain,1)="/" then Cookie_Domain = mid(Cookie_Domain,1,len(Cookie_Domain) - 1)
end if	

If G_VIRTUAL_ROOT_DIR<>"" Then
	SysRoot = "/"&G_VIRTUAL_ROOT_DIR
Else
	SysRoot = ""
End If
User_Path_Str = SysRoot & "/" & G_USER_DIR
User_Path_Str = replace(User_Path_Str,"//","/")
TypeStr = CintStr(Request.QueryString("DisTF")) '0,1
Login_DisType = NoSqlHack(Request.QueryString("DisType"))'transverse:vertical  
Str_Spanid = NoSqlHack(Request.QueryString("spanid"))'FS400_User_Login
If TypeStr = "" Or Not IsNumeric(TypeStr) Or Login_DisType = "" Or Str_Spanid = "" Then
	Response.Write "Err$$$登录标签意外错误，请重建。"
End If
If Cint(TypeStr) = 0 Then
	Str_Type = Login_DisType
Else
	Str_Type = ""
End If
Str_Act = Trim(Request.QueryString("Act"))
If Str_Act <> "" Then
	IF Str_Act = "Login" Then
		Str_name = Request.QueryString("UserName")
		Str_Pass = Request.QueryString("Password")
		If Str_name = "" Or Str_Pass = "" Then
			Response.Write "Err$$$用户名或密码不能为空或包含特殊字符"
		Else
			Str_Pass = Md5(Str_Pass,16)
		End If
		Str_UserLogin = UserLogIn(NoSqlHack(Str_name),NoSqlHack(Str_Pass),"0")
		If Str_UserLogin = True Then
			Response.Write ShowUserInfo(Str_Type,TypeStr,Login_DisType)
		Else
			Response.Write Str_UserLogin
		End If
	ElseIf Str_Act = "Logout" Then
		Call User_LogOut()
		Call Check_User_State()
	Else
		Call Check_User_State()	
	End If	
Else
	Call User_Login_Script()
End If	


Sub Check_User_State()
'=========================================================================
'判断用户状态，如果已经登陆，则显示已登陆界面，如果未登陆，则显示登陆界面
'=========================================================================	
	If Session("FS_UserName") = "" Or Session("FS_UserPassword") = "" Then
		ShowInfoStr = GetUserLoginForm(TypeStr,Login_DisType)
		ShowInfoStr = Replace(Replace(ShowInfoStr,Chr(10),""),Chr(13),"")
		Response.Write ShowInfoStr
	Else
		Str_UserInfo = ShowUserInfo(Str_Type,TypeStr,Login_DisType)
		If Str_UserInfo = "" Then
			ShowInfoStr = GetUserLoginForm(TypeStr,Login_DisType)
			ShowInfoStr = Replace(Replace(ShowInfoStr,Chr(10),""),Chr(13),"")
			Response.Write ShowInfoStr
		Else
			ShowInfoStr = Str_UserInfo
			ShowInfoStr = Replace(Replace(ShowInfoStr,Chr(10),""),Chr(13),"")
			Response.Write ShowInfoStr
		End If
	End If
End Sub	
	

'============================================
'自定义用户登录界面
'============================================
Function GetUserLoginForm(DisTF,DisType)
	If DisTF = "" Or Not IsNumeric(DisTF) Or DisType = "" Then
		GetUserLoginForm = "登陆标签发生意外错误，请重建"
		Exit Function
	End If
	IF Cint(DisTF) = 0 Then
		If DisType <> "vertical" And DisType <> "transverse" Then
			DisType = "vertical"
		End If
		If DisType = "vertical" Then
			GetUserLoginForm = "<form action="""&User_Path_Str&"/Checklogin.asp?forward="&Server.URLEncode(Request.ServerVariables("HTTP_REFERER"))&""" method=""post"" name=""LoginForm"" id=""LoginForm"" style=""margin:0;"">" & vbnewline
			GetUserLoginForm = GetUserLoginForm & "<label for=""Logintype"">类型：</label>" & vbnewline
			GetUserLoginForm = GetUserLoginForm & "<select name=""Logintype"" class=""Input_FSLogin"" id=""Logintype"">" & vbnewline
			GetUserLoginForm = GetUserLoginForm & "<option value=""0"" selected>用户名</option>" & vbnewline
			GetUserLoginForm = GetUserLoginForm & "<option value=""1"">用户编号</option>" & vbnewline
			GetUserLoginForm = GetUserLoginForm & "<option value=""2"">电子邮件</option>" & vbnewline
			GetUserLoginForm = GetUserLoginForm & "</select>" & vbnewline
			GetUserLoginForm = GetUserLoginForm & "<br>"
			GetUserLoginForm = GetUserLoginForm & "<label for=""Name"">用户：</label>" & vbnewline
			GetUserLoginForm = GetUserLoginForm & "<input name=""Name"" type=""text"" class=""Input_FSLogin"" id=""Name"" value="""" size=""12"" maxlength=""18"">" & vbnewline
			GetUserLoginForm = GetUserLoginForm & "<label for=""FS400_AutoGet""></label>" & vbnewline
			GetUserLoginForm = GetUserLoginForm & "<br>" & vbnewline
			GetUserLoginForm = GetUserLoginForm & "<label for=""password"">密码：</label>" & vbnewline
			GetUserLoginForm = GetUserLoginForm & "<input name=""password"" type=""password"" class=""Input_FSLogin"" id=""password"" size=""12"" maxlength=""20"">" & vbnewline
			GetUserLoginForm = GetUserLoginForm & "<div><input type=""submit"" name=""Login"" onclick=""return User_Login_Check(this.form);"" class=""Button_FSLogin"" value=""登录"" />&nbsp;&nbsp;&nbsp;&nbsp;<input type=""reset"" name=""FS400_Reset"" class=""Button_FSLogin"" value=""清空"" /></div>" & vbnewline
  			GetUserLoginForm = GetUserLoginForm & "<div><a href=""" & SysRoot & "/" & G_USER_DIR & "/GetPassword.asp"" target=""_blank"">忘记密码</a>&nbsp;&nbsp;<a href=""" & SysRoot & "/" & G_USER_DIR & "/Register.asp"" target=""_blank"">注册用户</a></div></form>"
		Else
			GetUserLoginForm = "<form action="""&User_Path_Str&"/Checklogin.asp?forward="&Server.URLEncode(Request.ServerVariables("HTTP_REFERER"))&""" method=""post"" name=""LoginForm"" id=""LoginForm"" style=""margin:0;"">" & vbnewline
			GetUserLoginForm = GetUserLoginForm &"<select name=""Logintype"" class=""Input_FSLogin"" id=""Logintype"">" & vbnewline
			GetUserLoginForm = GetUserLoginForm & "<option value=""0"" selected>用户名</option>" & vbnewline
			GetUserLoginForm = GetUserLoginForm & "<option value=""1"">用户编号</option>" & vbnewline
			GetUserLoginForm = GetUserLoginForm & "<option value=""2"">电子邮件</option>" & vbnewline
			GetUserLoginForm = GetUserLoginForm & "</select>" & vbnewline
			GetUserLoginForm = GetUserLoginForm & "<label for=""Name"">用户：</label>" & vbnewline
			GetUserLoginForm = GetUserLoginForm & "<input name=""Name"" type=""text"" class=""Input_FSLogin"" id=""Name"" value="""" size=""12"" maxlength=""18"" />" & vbnewline
			GetUserLoginForm = GetUserLoginForm & "<label for=""password"">密码：</label>" & vbnewline
			GetUserLoginForm = GetUserLoginForm & "<input name=""password"" type=""password"" class=""Input_FSLogin"" id=""password"" size=""12"" maxlength=""20"" />" & vbnewline
			GetUserLoginForm = GetUserLoginForm & "&nbsp;&nbsp;<input type=""submit"" name=""Login"" onclick=""return User_Login_Check(this.form);"" class=""Button_FSLogin"" value=""登录"" />&nbsp;&nbsp;" & vbnewline
			GetUserLoginForm = GetUserLoginForm & "<input type=""reset"" name=""FS400_Reset"" class=""Button_FSLogin"" value=""清空"" />&nbsp;&nbsp;" & vbnewline
			GetUserLoginForm = GetUserLoginForm & "<a href=""" & SysRoot & "/" & G_USER_DIR & "/GetPassword.asp"" target=""_blank"">忘记密码</a>&nbsp;&nbsp;<a href=""" & SysRoot & "/" & G_USER_DIR & "/Register.asp"" target=""_blank"">注册用户</a></form>"
		End If
	Else
		GetUserLoginForm = ReplaceStyleForLogin(DisType)
	End If	
End Function

'============================================
'自定义用户登录界面样式替换
'============================================
Function ReplaceStyleForLogin(DisType)
	Dim StyleID,SeCss,MenuCss,TxtCss,SumBitCss,ResestCss,Regcss,GetpasCss
	Dim SimButStr,SelectStr,NameStr,PassStr,ResetStr,RegLinkStr,GetPassStr
	Dim GetStyleRs,Str_Content
	If Instr(DisType,"┆") = 0 Then
		ReplaceStyleForLogin = "登陆标签发生意外错误，请重建"
		Exit Function
	Else
		If UBound(Split(DisType,"┆")) <> 7 Then
			ReplaceStyleForLogin = "登陆标签发生意外错误，请重建"
			Exit Function
		End If	 
	End If	
	StyleID = Split(DisType,"┆")(0)
	SeCss = Split(DisType,"┆")(1)
	MenuCss = Split(DisType,"┆")(2)
	TxtCss = Split(DisType,"┆")(3)
	SumBitCss = Split(DisType,"┆")(4)
	ResestCss = Split(DisType,"┆")(5)
	Regcss = Split(DisType,"┆")(6)
	GetpasCss = Split(DisType,"┆")(7)
	If StyleID = "" Or Not IsNumeric(StyleID) Then
		ReplaceStyleForLogin = "登陆标签没有选择样式，请重建"
		Exit Function
	End If
	If SeCss <> "" Then
		SeCss = " class=""" & SeCss & """"
	End If
	If MenuCss <> "" Then
		MenuCss = " class=""" & MenuCss & """"
	End If
	SelectStr = "<select name=""Logintype""" & SeCss & " id=""Logintype"">" & vbnewline
	SelectStr = SelectStr & "<option value=""0"" selected" & MenuCss & ">用户名</option>" & vbnewline
	SelectStr = SelectStr & "<option value=""1""" & MenuCss & ">用户编号</option>" & vbnewline
	SelectStr = SelectStr & "<option value=""2""" & MenuCss & ">电子邮件</option>" & vbnewline
	SelectStr = SelectStr & "</select>"
	IF TxtCss <> "" Then
		TxtCss = " class=""" & TxtCss & """"
		NameStr = "<input name=""Name""" & TxtCss & " type=""text"" id=""Name"" value="""" />"
		PassStr = "<input name=""password""" & TxtCss & " type=""password"" id=""password""  />"
	Else
		NameStr = "<input name=""Name"" class=""Input_FSLogin""  type=""text"" id=""Name"" value="""" />"
		PassStr = "<input name=""password"" class=""Input_FSLogin""  type=""password"" id=""password""  />"
	End if
	IF SumBitCss = "" Then
		SimButStr = "<input type=""submit"" name=""Login"" onclick=""return User_Login_Check(this.form);"" class=""Button_FSLogin"" value=""登录"" />"	
	Else
		If IsPicTF(SumBitCss) Then
			SimButStr = "<img src=""" & SumBitCss & """ border=""0"" onclick=""if(User_Login_Check(this.form)) document.LoginForm.submit();"" style=""cursor:hand;"">"
		Else
			SimButStr = "<input type=""submit"" name=""Login"" onclick=""return User_Login_Check(this.form);"" class=""" & SumBitCss & """ value=""登录"" />"	
		End if
	End If
	IF ResestCss = "" Then
		ResetStr = "<input type=""reset"" name=""FS400_Reset"" class=""Button_FSLogin"" value=""清空"" />"	
	Else
		If IsPicTF(ResestCss) Then
			ResetStr = "<img src=""" & ResestCss & """ border=""0"" onclick=""document.LoginForm.reset();"" style=""cursor:hand;"">"
		Else
			ResetStr = "<input type=""reset"" name=""FS400_Reset"" class=""" & ResestCss & """ value=""清空"" />"	
		End if
	End If
	If Regcss = "" Then
		RegLinkStr = "<a href=""" & SysRoot & "/" & G_USER_DIR & "/Register.asp"" target=""_blank"">注册用户</a>"
	Else
		If IsPicTF(Regcss) Then
			RegLinkStr = "<a href=""" & SysRoot & "/" & G_USER_DIR & "/Register.asp"" target=""_blank""><img src=""" & Regcss & """ border=""0""></a>"
		Else
			RegLinkStr = "<a class=""" & Regcss & """ href=""" & SysRoot & "/" & G_USER_DIR & "/Register.asp"" target=""_blank"">注册用户</a>"
		End If	
	End If
	If GetpasCss = "" Then
		GetPassStr = "<a href=""" & SysRoot & "/" & G_USER_DIR & "/GetPassword.asp"" target=""_blank"">忘记密码</a>"
	Else
		If IsPicTF(GetpasCss) Then
			GetPassStr = "<a href=""" & SysRoot & "/" & G_USER_DIR & "/GetPassword.asp"" target=""_blank""><img src=""" & GetpasCss & """ border=""0""></a>"
		Else
			GetPassStr = "<a class=""" & GetpasCss & """ href=""" & SysRoot & "/" & G_USER_DIR & "/GetPassword.asp"" target=""_blank"">忘记密码</a>"
		End If	
	End If			
	Set GetStyleRs = Conn.ExeCute("Select Content From FS_MF_Labestyle Where StyleType = 'Login' And ID = " & CintStr(StyleID))
	If GetStyleRs.Eof Then
		ReplaceStyleForLogin = "登陆标签所选样式不存在，请重建"
		Exit Function
	Else
		Str_Content = GetStyleRs(0)
		If Instr(Str_Content,"$*$") > 0 Then
			Str_Content = Split(Str_Content,"$*$")(0)
		Else
			Str_Content = Str_Content
		End If
	End If		
	GetStyleRs.Close : Set GetStyleRs = Nothing
	If Instr(Str_Content,"{Login_Name}") = 0 Or Instr(Str_Content,"{Login_Password}") = 0 Or Instr(Str_Content,"{Login_Simbut}") = 0 Then
		ReplaceStyleForLogin = "登陆标签所选样式不存在必选项，请重建"
		Exit Function
	End If
	Str_Content = Replace(Str_Content,"{Login_Name}",NameStr)
	Str_Content = Replace(Str_Content,"{Login_Password}",PassStr)
	Str_Content = Replace(Str_Content,"{Login_Simbut}",SimButStr)
	If Instr(Str_Content,"{Login_Type}") > 0 Then
		Str_Content = Replace(Str_Content,"{Login_Type}",SelectStr)
	End If
	If Instr(Str_Content,"{Login_Reset}") > 0 Then
		Str_Content = Replace(Str_Content,"{Login_Reset}",ResetStr)
	End If
	If Instr(Str_Content,"{Reg_LinkUrl}") > 0 Then
		Str_Content = Replace(Str_Content,"{Reg_LinkUrl}",RegLinkStr)
	End If
	If Instr(Str_Content,"{Get_PassLink}") > 0 Then
		Str_Content = Replace(Str_Content,"{Get_PassLink}",GetPassStr)
	End If
	ReplaceStyleForLogin = "<form action="""&User_Path_Str&"/Checklogin.asp?forward="&Server.URLEncode(Request.ServerVariables("HTTP_REFERER"))&""" method=""post"" name=""LoginForm"" id=""LoginForm"" style=""margin:0;"">"&Str_Content&"</form>"
End Function

'========================================================
'判断字符串是否为图片地址，传入字符串，返回True Or False
'========================================================	
Function IsPicTF(Str)
	If Str = "" Then
		IsPicTF = False
	End If
	If (Instr(Str,"/") > 0 Or Left(Str,1) = "/" Or Left(Lcase(Str),7) = "http://") And (Right(Lcase(Str),4) = ".jpg" Or Right(Lcase(Str),4) = ".gif" Or Right(Lcase(Str),4) = ".png") Then
		IsPicTF = True
	Else
		IsPicTF = False
	End if	
End Function


'============================================
'已登陆界面
'============================================
Function ShowUserInfo(Str_Type,TypeStr,Login_DisType)
	Dim Rs_UserInfo,Str_UserInfo,Str_BR
	Set Rs_UserInfo = User_Conn.execute("select Integral,FS_Money,LoginNum,ConNumber From FS_ME_Users where UserName='"&session("FS_UserName")&"' and UserPassword='"& Session("FS_UserPassword")&"'")
	Str_UserInfo=""
	If Not Rs_UserInfo.Eof Then
		If Cint(TypeStr) = 0 Then
			If Str_Type = "vertical" Then
				Str_BR = "<br />"
			Else
				Str_BR = ""	
			End If
			Str_UserInfo = "欢迎您,"&Session("FS_UserName")&" "&Str_BR&vbNewLine
			Str_UserInfo = Str_UserInfo & "积分:"&Rs_UserInfo("Integral")&"&nbsp;&nbsp;金币:"&Rs_UserInfo("FS_Money")&" "&Str_BR&"登录次数:"&Rs_UserInfo("LoginNum")&"&nbsp;&nbsp;投稿数:"&Rs_UserInfo("ConNumber")&" "&Str_BR&vbNewLine
			Str_UserInfo = Str_UserInfo & "<a href=""http://"&Cookie_Domain&"/"&G_USER_DIR&"/main.asp"" target=""_blank"">管理面板</a>  <a href=""http://"&Cookie_Domain&"/"&G_USER_DIR&"/Loginout.asp"" >注销</a>"
			ShowUserInfo = Str_UserInfo
		Else
			If ReplaceLoginDisStr(Login_DisType,Rs_UserInfo) = "" Then
				Str_UserInfo = "欢迎您,"&Session("FS_UserName")&" "&Str_BR&vbNewLine
				Str_UserInfo = Str_UserInfo & "积分:"&Rs_UserInfo("Integral")&"&nbsp;&nbsp;金币:"&Rs_UserInfo("FS_Money")&" "&Str_BR&"登录次数:"&Rs_UserInfo("LoginNum")&"&nbsp;&nbsp;投稿数:"&Rs_UserInfo("ConNumber")&" "&Str_BR&vbNewLine
				Str_UserInfo = Str_UserInfo & "<a href=""http://"&Cookie_Domain&"/"&G_USER_DIR&"/main.asp"" target=""_blank"">管理面板</a>  <a href=""http://"&Cookie_Domain&"/"&G_USER_DIR&"/Loginout.asp"">注销</a>"
				ShowUserInfo = Str_UserInfo
			Else
				ShowUserInfo = 	ReplaceLoginDisStr(Login_DisType,Rs_UserInfo)
			End If	
		End If	
	Else
		Call User_LogOut()
		ShowUserInfo = ""
	End If
End Function

'============================================
'替换显示后样式
'============================================
Function ReplaceLoginDisStr(Login_DisType,Obj)
	Dim StyLeRs,StyleID,ContStr
	If Instr(Login_DisType,"┆") = 0 Then
		ReplaceLoginDisStr = "登陆标签发生意外错误，请重建"
		Exit Function
	Else
		If UBound(Split(Login_DisType,"┆")) <> 7 Then
			ReplaceLoginDisStr = "登陆标签发生意外错误，请重建"
			Exit Function
		End If	 
	End If
	StyleID = Split(Login_DisType,"┆")(0)
	Set StyLeRs = Conn.ExeCute("Select Content From FS_MF_Labestyle Where StyleType = 'Login' And ID = " & CintStr(StyleID))
	If StyLeRs.Eof Then
		ReplaceLoginDisStr = "登陆标签所选样式不存在，请重建"
		Exit Function
	Else
		ContStr = StyLeRs(0)
		If Instr(ContStr,"$*$") > 0 Then
			ContStr = Split(ContStr,"$*$")(1)
		Else
			ContStr = ""
		End If	
	End If
	StyLeRs.Close : Set StyLeRs = Nothing
	If ContStr = "" Then
		ReplaceLoginDisStr = ""
	Else  ',,,,,,
		If Instr(ContStr,"{User_Name}") > 0 Then
			ContStr = Replace(ContStr,"{User_Name}",Session("FS_UserName"))
		End If
		If Instr(ContStr,"{User_JiFen}") > 0 Then
			ContStr = Replace(ContStr,"{User_JiFen}",Obj("Integral"))
		End If
		If Instr(ContStr,"{User_JinBi}") > 0 Then
			ContStr = Replace(ContStr,"{User_JinBi}",Obj("FS_Money"))
		End If
		If Instr(ContStr,"{User_LoginTimes}") > 0 Then
			ContStr = Replace(ContStr,"{User_LoginTimes}",Obj("LoginNum"))
		End If
		If Instr(ContStr,"{User_TouGao}") > 0 Then
			ContStr = Replace(ContStr,"{User_TouGao}",Obj("ConNumber"))
		End If
		If Instr(ContStr,"{User_ConCenter}") > 0 Then
			ContStr = Replace(ContStr,"{User_ConCenter}","<a target=""_blank"" href=""http://"&Cookie_Domain&"/"&G_USER_DIR&"/main.asp"" target=""_blank"">管理面板</a>")
		End If
		If Instr(ContStr,"{User_LogOut}") > 0 Then
			ContStr = Replace(ContStr,"{User_LogOut}","<a href=""http://"&Cookie_Domain&"/"&G_USER_DIR&"/Loginout.asp"">注销</a>")
		End If
		ReplaceLoginDisStr = ContStr
	End If	
End Function

Conn.Close():Set Conn = Nothing
User_Conn.Close : set User_Conn = Nothing
%>
<% Sub User_Login_Script() %>
var DisIDStr = '<% = Str_Spanid %>';
var FS400_UserPath = '<%= User_Path_Str %>';
var Login_DisType = '<% = Server.URLEncode(Login_DisType) %>';
var Login_TypeStr = '<% = TypeStr %>'; 
function User_Login_Check(loginobj)
{
	var CheckStr = '';
	var Fs_FocusStr = '';
	var Namestr = $('Name').value;
	var PassStr = $('password').value;
	if (Namestr == '')
	{
		CheckStr = '用户名  不能为空';
		Fs_FocusStr = $('Name');
	}
	else
	{
		if (PassStr == '')
		{
			CheckStr = CheckStr + '\n密  码  不能为空';
			Fs_FocusStr = $('password');
		}
		else
		{
			CheckStr = '';
			Fs_FocusStr = '';
		}
		}
	if (CheckStr != '')
	{
		CheckStr = '提示：\n' + CheckStr;
		alert(CheckStr);
		Fs_FocusStr.focus();
		return false;
	}
	else
	{
		return true;
	}
}

function Request(action)
{
	var LoginUrl = FS400_UserPath+'/m_UserLogin.asp';
	var myAjax = new Ajax.Request(
		LoginUrl,
		{method:'get',
		parameters:action,
		onComplete:Login_Receive
		}
		);
}
function Login_Receive(XmlObj)
{
	if (XmlObj.responseText.indexOf("$$$")>-1)
	{
		check=XmlObj.responseText.split("$$$");
		switch (check[0]) 
		{
			case "Err" :
				alert(check[1]);
				break;
			default :
				alert(check[1]);
		}
	}
	else
	{
		$(DisIDStr).innerHTML=XmlObj.responseText;
	}
	
}
function User_Check_State()
{
	var action = 'Act=CheckState&DisTF='+Login_TypeStr+'&DisType='+Login_DisType+'&spanid='+DisIDStr;
	Request(action);
}
User_Check_State();
<% End Sub %>






