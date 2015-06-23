<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Md5.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<!--#include file="../API/Cls_PassportApi.asp" -->
<%
If Request.Form("Action") = "Save" then
		if Trim(Request.Form("PassQuestion"))="" then
			strShowErr = "<li>请输入密码提示问题!</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Myvalidcode.asp")
			Response.end
		End if
		Dim RsSaveTFObj
		Set RsSaveTFObj = server.CreateObject(G_FS_RS)
		RsSaveTFObj.open "select  UserID From FS_ME_Users where UserNumber = '"& Fs_User.UserNumber &"' and UserPassword='"& md5(Request.Form("UserPassword"),16)&"'",User_Conn,1,3
		if Not RsSaveTFObj.eof then
			Dim RsSaveIObj
			Set RsSaveIObj = server.CreateObject(G_FS_RS)
			RsSaveIObj.open "select  UserID,UserPassword,PassQuestion,PassAnswer,safeCode,OnlyLogin From FS_ME_Users where UserNumber = '"& Fs_User.UserNumber &"'",User_Conn,1,3
			if Request.Form("UserPassword_1")<>"" then
				RsSaveIObj("UserPassword") = Md5(Request.Form("UserPassword_1"),16)
			End if
			RsSaveIObj("PassQuestion") = NoSqlHack(Replace(Request.Form("PassQuestion"),"''",""))
			if Request.Form("PassAnswer")<>"" then
				RsSaveIObj("PassAnswer") = Md5(Request.Form("PassAnswer"),16)
			End if
			if Request.Form("safeCode")<>"" then
				RsSaveIObj("safeCode") = Md5(Request.Form("safeCode"),16)
			End if
			if Request.Form("OnlyLogin")<>"" then
				RsSaveIObj("OnlyLogin")  =0
			Else
				RsSaveIObj("OnlyLogin")  =1
			End if
			RsSaveIObj.update
			RsSaveIObj.close:set RsSaveIObj = nothing
			'-----------------------------------------------------------------
			'系统整合
			'-----------------------------------------------------------------
			Dim API_Obj,API_SaveCookie,SysKey
			If API_Enable Then
				Set API_Obj = New PassportApi
					API_Obj.NodeValue "action","update",0,False
					API_Obj.NodeValue "username",Fs_User.UserName,1,False
					API_Obj.NodeValue "email","",1,False
					API_Obj.NodeValue "question",NoSqlHack(Request.Form("PassQuestion")),1,False
					API_Obj.NodeValue "answer",NoSqlHack(Request.Form("PassAnswer")),1,False
					SysKey = Md5(API_Obj.XmlNode("username")&API_SysKey,16)
					API_Obj.NodeValue "syskey",SysKey,0,False
					API_Obj.NodeValue "password",NoSqlHack(Request.Form("UserPassword_1")),0,False
					API_Obj.SendHttpData
					If API_Obj.Status = "1" Then
						Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(API_Obj.Message)&"&ErrorUrl=")
						Response.end
					End If
				Set API_Obj = Nothing
			End If
			'-----------------------------------------------------------------
			strShowErr = "<li>安全资料修改成功!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Myvalidcode.asp")
			Response.end
		Else
			strShowErr = "<li>您的原密码不正确!</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Myvalidcode.asp")
			Response.end
		End if
		RsSaveTFObj.close:set RsSaveTFObj = nothing
Else
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-我的安全资料</title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,风讯,风讯网站内容管理系统,风讯系统,风讯新闻系统,风讯商城,风讯b2c,新闻系统,CMS,域名空间,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head>
<body>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td>
      <!--#include file="top.asp" -->
    </td>
  </tr>
</table>
<table width="98%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  
    <tr class="back"> 
      <td   colspan="2" class="xingmu" height="26"> <!--#include file="Top_navi.asp" --> </td>
    </tr>
    <tr class="back"> 
      <td width="18%" valign="top" class="hback"> <div align="left"> 
          <!--#include file="menu.asp" -->
        </div></td>
      <td width="82%" valign="top" class="hback"><table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td class="hback"><strong>位置：</strong><a href="../">网站首页</a> &gt;&gt; 
            <a href="main.asp">会员首页</a> &gt;&gt; 安全资料</td>
        </tr>
      </table> 
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="form1" method="post" action="">
          <tr class="hback"> 
            <td class="hback_1"><div align="right"><strong>原密码</strong></div></td>
            <td class="hback"><input name="UserPassword" type="password" id="UserPassword" size="26" maxlength="20"> 
            </td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right"><strong>新密码</strong></div></td>
            <td class="hback"><input name="UserPassword_1" type="password" id="UserPassword_1" size="26" maxlength="20">
              不修改保持为空</td>
          </tr>
          <tr class="hback"> 
            <td width="14%" class="hback_1"><div align="right"><strong>密码提示问题</strong></div></td>
            <td width="86%" class="hback"><input name="PassQuestion" type="text" id="PassQuestion" value="<% = Fs_User.PassQuestion%>" size="26" maxlength="20"></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right"><strong>密码答案</strong></div></td>
            <td class="hback"><input name="PassAnswer" type="text" id="PassAnswer" size="26" maxlength="50">
              不修改保持为空</td>
          </tr>
          <tr class="hback">
            <td class="hback_1"><div align="right"><strong>安全码</strong></div></td>
            <td class="hback"><input name="safeCode" type="text" id="safeCode" size="26" maxlength="50" readonly>
              不修改保持为空</td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right"><strong>多人登陆</strong></div></td>
            <td class="hback"><input type="checkbox" name="OnlyLogin" value="0"  <% if  Fs_User.OnlyLogin = 0 then response.Write("checked")%>>
              开启</td>
          </tr>
          <tr class="hback"> 
            <td height="45" colspan="2" class="hback"> <div align="left"> 
                <input name="Action" type="hidden" id="Action" value="Save">
                <input type="submit" name="Submit" value="保存资料"   onClick="{if(confirm('确认保存您所修改的参数吗?')){this.document.form1.submit();return true;}return false;}">
                　 
                <input type="reset" name="Submit3" value="重新填写">
              </div></td>
          </tr>
          <tr class="hback"> 
            <td height="45" colspan="2" class="hback">　 
              <div align="left"></div></td>
          </tr>
        </form>
      </table>
       </td>
    </tr>
    <tr class="back"> 
      <td height="20"  colspan="2" class="xingmu"> <div align="left"> 
          <!--#include file="Copyright.asp" -->
        </div></td>
    </tr>
</table>
</body>
</html>
<%
End if
Set Fs_User = Nothing
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->





