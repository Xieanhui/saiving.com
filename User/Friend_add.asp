<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
Dim Returvaluestr_1,StrRealName,StrEmail,StrTel,Strmobile,StrQQ,StrMSN,StrContent,strFriendType,strFriendType_1
if Request.Form("Action") = "Save" then
	Dim UserName ,RealName,ResultMTF,id
	UserName=NoSqlHack(Request.Form("UserName"))
	RealName=NoSqlHack(Request.Form("RealName"))
	id=NoSqlHack(Request.Form("ID"))
	if Trim(UserName) = "" or Trim(RealName) = ""  then
		strShowErr = "<li>请输入用户名及备注的姓名</li>"
		Call ReturnError(strShowErr,"")
	Elseif Len(Request.Form("Content"))>200  then
		strShowErr = "<li>备注不能大于200个字符</li>"
		Call ReturnError(strShowErr,"")
	Elseif UserName=Fs_User.UserName  then
		strShowErr = "<li>不能自己添加自己</li>"
		Call ReturnError(strShowErr,"")
	Else
		Dim Returvaluestr,RsCheckTFObj,RsGObj
		Returvaluestr = Fs_User.GetFriendNumber(UserName)
		Set RsGObj = server.CreateObject(G_FS_RS)
		RsGObj.open "select  isLock,UserID From FS_ME_Users where UserNumber = '"& Returvaluestr &"'",User_Conn,1,3
		if RsGObj.eof then
				strShowErr = "<li>找不到此用户！</li>"
				Call ReturnError(strShowErr,"")
		Else
				if RsGObj(0) =1 then
					strShowErr = "<li>用户已经被锁定，不能添加！</li>"
					Call ReturnError(strShowErr,"")
				End if
		End if
		Set RsCheckTFObj = server.CreateObject(G_FS_RS)
		RsCheckTFObj.open "select  FriendID From FS_ME_Friends where F_UserNumber = '"& Returvaluestr &"' and UserNumber='"& Fs_User.UserNumber&"'",User_Conn,1,3
		if Not RsCheckTFObj.eof then
			if id ="" then
					strShowErr = "<li>好友已经存在！</li>"
					Call ReturnError(strShowErr,"")
			End if
		End if
			Dim RsaddFLObj,addFLSQL,strUpdatechar
			if id <>"" then
				Set RsaddFLObj = server.CreateObject(G_FS_RS)
				addFLSQL = "select  * From FS_ME_Friends  where FriendID ="& CintStr(id) &""
				RsaddFLObj.open addFLSQL,User_Conn,1,3
			Else
				Set RsaddFLObj = server.CreateObject(G_FS_RS)
				addFLSQL = "select  * From FS_ME_Friends Where 1=0"
				RsaddFLObj.open addFLSQL,User_Conn,1,3
				RsaddFLObj.addnew
			End if
			RsaddFLObj("UserNumber") = Fs_User.UserNumber
			RsaddFLObj("FriendType") = CintStr(Request.Form("FriendType"))
			RsaddFLObj("F_UserNumber") = NoSqlHack(Returvaluestr)
			RsaddFLObj("AddTime") = now
			RsaddFLObj("Updatetime") = now
			RsaddFLObj("RealName") = NoSqlHack(Request.Form("RealName"))
			RsaddFLObj("Content") = NoSqlHack(NoHtmlHackInput(Request.Form("Content")))
			RsaddFLObj("Email") = NoSqlHack(Request.Form("Email"))
			RsaddFLObj("Tel") = NoSqlHack(Request.Form("Tel"))
			RsaddFLObj("Mobile") = NoSqlHack(Request.Form("Mobile"))
			RsaddFLObj("QQ") = NoSqlHack(Request.Form("QQ"))
			RsaddFLObj("MSN") = NoSqlHack(Request.Form("MSN"))
			RsaddFLObj.update
			RsaddFLObj.close:set RsaddFLObj = nothing
			strShowErr = "<li>好友操作成功！</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Friend.asp")
			Response.end
	End if
End if
if Request.QueryString("FriendID") <> "" then
	Dim FriendID
	FriendID = CintStr(Request.QueryString("FriendID"))
	Dim RsUserFriendObj1,RsUserFriendSQL
	Set RsUserFriendObj1 = Server.CreateObject(G_FS_RS)
	RsUserFriendSQL = "Select FriendID,UserNumber,FriendType,F_UserNumber,AddTime,Updatetime,RealName,Content,Email,Tel,Mobile,QQ,MSN From FS_ME_Friends  where UserNumber='"&Fs_User.UserNumber&"' and FriendID = "& FriendID &" Order by FriendID desc"
	RsUserFriendObj1.Open RsUserFriendSQL,User_Conn,1,1
	If Trim(Request.QueryString("action"))="addFriend" then
		Returvaluestr_1 = Fs_User.GetFriendName(request.QueryString("ToUserNumber"))
	Else
		Returvaluestr_1 = Fs_User.GetFriendName(RsUserFriendObj1("F_UserNumber"))
	End if
	StrRealName = RsUserFriendObj1("RealName")
	StrEmail = RsUserFriendObj1("Email")
	StrTel = RsUserFriendObj1("Tel")
	Strmobile = RsUserFriendObj1("mobile")
	StrQQ = RsUserFriendObj1("qq")
	StrMSN = RsUserFriendObj1("MSN")
	StrContent =  RsUserFriendObj1("Content")
	strFriendType = RsUserFriendObj1("FriendType")
	RsUserFriendObj1.close
	set RsUserFriendObj1 = nothing
Else
	if Request.QueryString("type") =1 then
		strFriendType = 1
	Elseif Request.QueryString("type") =2 then
		strFriendType = 2
	Else
		strFriendType = 0
		strFriendType_1 = NoSqlHack(request.QueryString("UserName"))
	End if
	If Trim(Request.QueryString("action"))="addFriend" then
		Returvaluestr_1 = Fs_User.GetFriendName(NoSqlHack(request.QueryString("ToUserNumber")))
	End if
End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-添加/修改朋友</title>
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
            <a href="main.asp">会员首页</a> &gt;&gt; <a href="Friend.asp">朋友管理</a> &gt;&gt;  添加/修改朋友</td>
          </tr>
        </table>
        
      
        
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="UserForm" method="post" action="" onSubmit="return CheckForm();">
          <tr class="hback"> 
            <td width="11%" class="hback_1"><div align="center" class="tx">* 用户名</div></td>
            <td width="30%" class="hback"><div align="left"> 
                <input name="UserName" type="text" id="UserName" value="<% = Returvaluestr_1 %><% = strFriendType_1%>" size="30" maxlength="50" <%if Request.QueryString("FriendID")<>"" then Response.Write("ReadOnly")%>>
              </div></td>
            <td width="59%" class="hback"><div align="left"> </div></td>
          </tr>
          <tr class="hback"> 
            <td colspan="3" class="xingmu"><div align="left">备注部分</div></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center"  class="tx">*姓名</div></td>
            <td class="hback"><input name="RealName" type="text" id="RealName" value="<% = StrRealName %>" size="30" maxlength="20"></td>
            <td class="hback">好友的备注名字</td> 
          </tr> 
          <tr class="hback"> 
            <td class="hback_1"><div align="center">电子邮件</div></td> 
            <td class="hback">
				<input name="Email" type="text" id="Email" value="<% = StrEmail%>" size="30" maxlength="150"></td>
            <td class="hback">好友的备注电子邮件</td> 
          </tr> 
          <tr class="hback"> 
            <td class="hback_1"><div align="center">电话</div></td>
            <td class="hback">
<input name="Tel" type="text" id="Tel" value="<% = StrTel%>" size="30" maxlength="24"></td>
            <td class="hback">好友的备注电话</td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center">手机</div></td>
            <td class="hback">
<input name="mobile" type="text" id="mobile" value="<% = Strmobile%>" size="30" maxlength="23"></td>
            <td class="hback">好友的备注手机</td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center">QQ</div></td>
            <td class="hback">
<input name="qq" type="text" id="qq" value="<% = StrQQ%>" size="30" maxlength="15"></td>
            <td class="hback">好友的备注QQ</td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center">MSN</div></td>
            <td class="hback">
<input name="MSN" type="text" id="MSN" value="<% = StrMSN%>" size="30" maxlength="150"></td>
            <td class="hback">好友的备注MSN</td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center">备注</div></td>
            <td class="hback">
<textarea name="Content" cols="30" rows="5" id="Content"><% = StrContent%></textarea></td>
            <td class="hback">好友的备注，最大200字符</td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center">类型</div></td>
            <td colspan="2" class="hback">
<select name="FriendType" id="FriendType">
                <option value="0" <%if strFriendType = 0 then response.Write("selected")%>>好朋友</option>
                <option value="1" <%if strFriendType= 1 then response.Write("selected")%>>陌生人</option>
                <option value="2" <%if strFriendType = 2 then response.Write("selected")%>>黑名单</option>
              </select></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1">&nbsp;</td>
            <td colspan="2" class="hback">
<input type="submit" name="Submit" value="提交好友资料"> 
              <input name="Action" type="hidden" id="Action" value="Save">
              <input name="Id" type="hidden" id="Id" value="<%=Trim(Request.QueryString("FriendID"))%>"></td>
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
Set Fs_User = Nothing
%>
<script language="JavaScript" type="text/javascript">
function CheckForm()
{
	if(document.UserForm.UserName.value=="")
	{
		alert("请填写用户名!");
		document.UserForm.UserName.focus();
		return false;
	}
	if(document.UserForm.RealName.value=="")
	{
		alert("请填写备注姓名!");
		document.UserForm.RealName.focus();
		return false;
	}
	}
</script>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->





