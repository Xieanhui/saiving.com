<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
If Request.Form("Action") = "Save" then
	if  len(trim(Request.Form("C_BankUserName")))>200 then
		strShowErr = "<li>卡号及开户名不能超过200个字符</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		Dim RsSaveIObj
		Set RsSaveIObj = server.CreateObject(G_FS_RS)
		RsSaveIObj.open "select C_BankUserName,C_BankName From FS_ME_CorpUser where UserNumber = '"& Fs_User.UserNumber &"' and CorpID="&CintStr(Request.Form("ID")),User_Conn,1,3
		RsSaveIObj("C_BankName") = NoSqlHack(Replace(Request.Form("C_BankName"),"''",""))
		RsSaveIObj("C_BankUserName")  = NoHtmlHackInput(Replace(Request.Form("C_BankUserName"),"''",""))
		RsSaveIObj.update
		RsSaveIObj.close
		set RsSaveIObj = nothing
		strShowErr = "<li>银行资料修改成功!</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Corp_Bank.asp")
		Response.end
	End if
Else
		Dim RsCorpObj
		Set RsCorpObj = server.CreateObject(G_FS_RS)
		RsCorpObj.open "select  CorpID,isLockCorp,C_BankName,C_BankUserName,UserNumber From FS_ME_CorpUser where UserNumber = '"& Fs_User.UserNumber &"'",User_Conn,1,3
		if RsCorpObj.eof then
			strShowErr = "<li>找不到企业数据</li>"
			Call ReturnError(strShowErr,"")
		End if
		if RsCorpObj("isLockCorp") = 1 then
			strShowErr = "<li>您的企业数据还没审核通过</li>"
			Call ReturnError(strShowErr,"")
		End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-企业银行管理</title>
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
            <a href="main.asp">会员首页</a> &gt;&gt; 企业银行管理</td>
        </tr>
      </table> 
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="UserForm" method="post" action="">
          <tr class="hback"> 
            <td width="16%" class="hback_1"><div align="center"><strong>开户银行</strong></div></td>
            <td class="hback"> <input name="C_BankName" type="text" id="C_BankName" value="<% = RsCorpObj("C_BankName")%>" size="26" maxlength="50"> 
              <div align="center"></div></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1" ><div align="center"><strong>帐号及开户名</strong></div></td>
            <td class="hback"><textarea name="C_BankUserName" cols="50" rows="6" id="C_BankUserName"><% = RsCorpObj("C_BankUserName")%></textarea>
              限制200个字符</td>
          </tr>
          <tr class="hback"> 
            <td colspan="2" class="hback"><div align="center"> 
                <input name="ID" type="hidden" id="ID" value="<% = RsCorpObj("CorpID")%>">
                <input name="Action" type="hidden" id="Action" value="Save">
                <input type="submit" name="Submit" value="保存资料"  onClick="{if(confirm('确认保存吗?')){this.document.UserForm.submit();return true;}return false;}">
                　 
                <input type="reset" name="Submit3" value="重新填写">
                　 </div></td>
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
	RsCorpObj.close
	set RsCorpObj = nothing
End if
Set Fs_User = Nothing
%>
<script language="JavaScript" type="text/JavaScript">
	function OpenWindowAndSetValue(Url,Width,Height,WindowObj,SetObj)
	{
		var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:no;');
		if (ReturnStr!='') SetObj.value=ReturnStr;
		return ReturnStr;
	}
</script>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->





