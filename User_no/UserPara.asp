<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
Dim str_CurrPath,FileName,str_FileName,rs_sys,str_FileExtName
str_CurrPath = Replace("/"&G_VIRTUAL_ROOT_DIR &"/"&G_USERFILES_DIR&"/"&Session("FS_UserNumber"),"//","/")
if Request.Form("Action") = "Save" then
	If len(Request.Form("strDescription"))>1000 Then
		strShowErr = "<li>站点描述不能超过1000个字符</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End If
	Dim RsAddParaObj
	Set RsAddParaObj = server.CreateObject(G_FS_RS)
	if Trim(Request.Form("state")) = "nodata" then
		RsAddParaObj.open "select  * From FS_ME_MySysPara where 1=0",User_Conn,1,3
		RsAddParaObj.addNew
		RsAddParaObj("UserNumber") = Fs_User.UserNumber
	Elseif Trim(Request.Form("state")) = "data" then
		RsAddParaObj.open "select  * From FS_ME_MySysPara where UserNumber='"& Fs_User.UserNumber &"'",User_Conn,1,3
	End if
	RsAddParaObj("DownFileRule") = ",,,,"
	RsAddParaObj("NewsFileRule") = ",,,,"
	RsAddParaObj("ProductFileRule") = ",,,,"
	RsAddParaObj("ilogFileRule") = ",,,,"
	RsAddParaObj("mysiteName") = NoSqlHack(Request.Form("mysiteName"))
	RsAddParaObj("visitePassword") = NoSqlHack(Request.Form("visitePassword"))
	RsAddParaObj("Keywords") = NoSqlHack(Request.Form("Keywords"))
	RsAddParaObj("Description") = NoSqlHack(Request.Form("strDescription"))
	RsAddParaObj("NaviPic") = NoSqlHack(Request.Form("NaviPic"))
	RsAddParaObj("RedirectUrl") = NoSqlHack(Request.Form("RedirectUrl"))
	if Request.Form("hideIP") <> "" then
		RsAddParaObj("hideIP") = 1
	Else
		RsAddParaObj("hideIP") = 0
	End if
	RsAddParaObj.Update
	RsAddParaObj.close
	set RsAddParaObj=nothing
		strShowErr = "<li>站点参数保存成功</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-站点设置</title>
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
            <a href="main.asp">会员首页</a> &gt;&gt; 站点设置</td>
          </tr>
        </table>
<%
Dim RsParaObj,strstate,mysiteName,visitePassword,Keywords,strDescription,NaviPic,isHtml,RedirectUrl,hideIP
Set RsParaObj = server.CreateObject(G_FS_RS)
RsParaObj.open "select  * From FS_ME_MySysPara where UserNumber='"& Fs_User.UserNumber &"'",User_Conn,1,3
if RsParaObj.eof then
	strstate = "nodata"
	mysiteName = "我的个人空间"
	visitePassword = ""
	Keywords = "风讯,CMS,Foosun"
	strDescription = "风讯,CMS,Foosun"
	NaviPic = ""
	'isHtml = RsParaObj("isHtml")
	RedirectUrl = ""
	hideIP = 0
Else
	strstate = "data"
	mysiteName = RsParaObj("mysiteName")
	visitePassword = RsParaObj("visitePassword") 
	Keywords = RsParaObj("Keywords")
	strDescription = NoHtmlHackInput(RsParaObj("Description"))
	NaviPic = RsParaObj("NaviPic")
	'isHtml = RsParaObj("isHtml")
	RedirectUrl = RsParaObj("RedirectUrl")
	hideIP = RsParaObj("hideIP")
End if
RsParaObj.Close
set RsParaObj = nothing
%>
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="form1" method="post" action="">
          <tr class="hback"> 
            <td width="15%" class="hback_1"><div align="right">我的站点名称 </div></td>
            <td width="39%" class="hback"><input name="mysiteName" type="text" id="mysiteName" value="<% = mysiteName %>" size="30"></td>
            <td width="46%" class="hback">&nbsp;</td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">站点密码</div></td>
            <td class="hback"><input name="visitePassword" type="text" id="visitePassword" value="<% = visitePassword%>" size="30"></td>
            <td class="hback">不设置密码请保持为空,不能出现非法字符</td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">站点关键字</div></td>
            <td class="hback"><input name="Keywords" type="text" value="<% = Keywords %>" size="30"></td>
            <td class="hback">最多100字符</td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">站点描述</div></td>
            <td class="hback"><textarea name="strDescription" cols="40" rows="6" id="strDescription"><% = strDescription %></textarea></td>
            <td class="hback">最多1000字符</td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">站点导航图片</div></td>
            <td class="hback"><input name="NaviPic" type="text" id="NaviPic" value="<% = NaviPic %>" size="30">
            <img  src="Images/upfile.gif" width="44" height="22" onClick="OpenWindowAndSetValue('CommPages/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,320,window,document.form1.NaviPic);" style="cursor:hand;"></td>
            <td class="hback">&nbsp;</td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">站点转向地址</div></td>
            <td class="hback"><input name="RedirectUrl" type="text" id="RedirectUrl" value="<% = RedirectUrl %>" size="30"></td>
            <td class="hback">&nbsp;</td>
          </tr>
          <tr class="hback">
            <td class="hback_1"><div align="right">隐藏IP</div></td>
            <td colspan="2" class="hback"><input name="hideIP" type="checkbox" id="hideIP" value="1" <% if hideIP =1 then response.Write("checked")%>>
              开启</td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right"></div></td>
            <td colspan="2" class="hback"><input type="button" name="Submit" value="保存设置"  onClick="{if(confirm('确认保存您所修改的参数吗?')){this.document.form1.submit();return true;}return false;}"> 
              <input name="state" type="hidden" id="state" value="<% = strstate %>">
              <input name="Action" type="hidden" id="Action" value="Save"></td>
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
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->





