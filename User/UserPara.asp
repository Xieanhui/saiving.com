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
		strShowErr = "<li>վ���������ܳ���1000���ַ�</li>"
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
		strShowErr = "<li>վ���������ɹ�</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-վ������</title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,��Ѷ,��Ѷ��վ���ݹ���ϵͳ,��Ѷϵͳ,��Ѷ����ϵͳ,��Ѷ�̳�,��Ѷb2c,����ϵͳ,CMS,�����ռ�,asp,jsp,asp.net,SQL,SQL SERVER" />
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
            
          <td class="hback"><strong>λ�ã�</strong><a href="../">��վ��ҳ</a> &gt;&gt; 
            <a href="main.asp">��Ա��ҳ</a> &gt;&gt; վ������</td>
          </tr>
        </table>
<%
Dim RsParaObj,strstate,mysiteName,visitePassword,Keywords,strDescription,NaviPic,isHtml,RedirectUrl,hideIP
Set RsParaObj = server.CreateObject(G_FS_RS)
RsParaObj.open "select  * From FS_ME_MySysPara where UserNumber='"& Fs_User.UserNumber &"'",User_Conn,1,3
if RsParaObj.eof then
	strstate = "nodata"
	mysiteName = "�ҵĸ��˿ռ�"
	visitePassword = ""
	Keywords = "��Ѷ,CMS,Foosun"
	strDescription = "��Ѷ,CMS,Foosun"
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
            <td width="15%" class="hback_1"><div align="right">�ҵ�վ������ </div></td>
            <td width="39%" class="hback"><input name="mysiteName" type="text" id="mysiteName" value="<% = mysiteName %>" size="30"></td>
            <td width="46%" class="hback">&nbsp;</td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">վ������</div></td>
            <td class="hback"><input name="visitePassword" type="text" id="visitePassword" value="<% = visitePassword%>" size="30"></td>
            <td class="hback">�����������뱣��Ϊ��,���ܳ��ַǷ��ַ�</td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">վ��ؼ���</div></td>
            <td class="hback"><input name="Keywords" type="text" value="<% = Keywords %>" size="30"></td>
            <td class="hback">���100�ַ�</td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">վ������</div></td>
            <td class="hback"><textarea name="strDescription" cols="40" rows="6" id="strDescription"><% = strDescription %></textarea></td>
            <td class="hback">���1000�ַ�</td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">վ�㵼��ͼƬ</div></td>
            <td class="hback"><input name="NaviPic" type="text" id="NaviPic" value="<% = NaviPic %>" size="30">
            <img  src="Images/upfile.gif" width="44" height="22" onClick="OpenWindowAndSetValue('CommPages/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,320,window,document.form1.NaviPic);" style="cursor:hand;"></td>
            <td class="hback">&nbsp;</td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">վ��ת���ַ</div></td>
            <td class="hback"><input name="RedirectUrl" type="text" id="RedirectUrl" value="<% = RedirectUrl %>" size="30"></td>
            <td class="hback">&nbsp;</td>
          </tr>
          <tr class="hback">
            <td class="hback_1"><div align="right">����IP</div></td>
            <td colspan="2" class="hback"><input name="hideIP" type="checkbox" id="hideIP" value="1" <% if hideIP =1 then response.Write("checked")%>>
              ����</td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right"></div></td>
            <td colspan="2" class="hback"><input type="button" name="Submit" value="��������"  onClick="{if(confirm('ȷ�ϱ��������޸ĵĲ�����?')){this.document.form1.submit();return true;}return false;}"> 
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
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->





