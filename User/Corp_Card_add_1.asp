<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
Dim p_id
p_id = NoSqlHack(Trim(Request.QueryString("CorpCardid")))
if Request.Form("Action") = "Save" then
	Dim str_f_UserNumber,str_Content,str_CorpCardid
	str_f_UserNumber = NoSqlHack(Trim(Request.Form("UserNumber")))
	str_Content = NoHtmlHackInput(Request.Form("Content"))
	str_CorpCardid = CintStr(Request.Form("CorpCardid"))
	if str_f_UserNumber ="" then
			strShowErr = "<li>����Ļ�Ա���</li>"
			Call ReturnError(strShowErr,"")
	End if
	if len(str_Content) >200 then
			strShowErr = "<li>��ע���ܴ���200���ַ�</li>"
			Call ReturnError(strShowErr,"")
	End if
			dim RsaddObj
			Set RsaddObj = server.CreateObject(G_FS_RS)
			RsaddObj.open "select * From FS_ME_CorpCard where CorpCardid="&CintStr(str_CorpCardid),User_Conn,1,3
			RsaddObj("Content") = str_Content
			RsaddObj.update
			RsaddObj.close
			set RsaddObj = nothing
			strShowErr = "<li>�޸ĳɹ�</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Corp_card.asp")
			Response.end
Else
			Dim RsCorpCardObj
			Set RsCorpCardObj = server.CreateObject(G_FS_RS)
			RsCorpCardObj.open "select CorpCardid,F_UserNumber,Content From FS_ME_CorpCard where CorpCardid="& CintStr(p_id),User_Conn,1,1
			iF RsCorpCardObj.eof then
				strShowErr = "<li>����Ĳ���</li>"
				Call ReturnError(strShowErr,"")
			End if
			dim strTmp
			strTmp = Fs_User.GetFriendName(RsCorpCardObj("F_UserNumber"))
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-�޸���Ƭ</title>
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
            <a href="main.asp">��Ա��ҳ</a> &gt;&gt; �����Ƭ</td>
          </tr>
        </table>
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="UserForm" method="post" action="">
          <tr class="hback"> 
            <td width="12%" class="hback_1"><div align="center" class="tx">�û���</div></td>
            <td width="40%" class="hback"><div align="left"> 
                <input name="UserNumber" type="text" id="UserNumber" value="<% = strTmp%>" size="30" ReadOnly>
              </div></td>
            <td width="48%" class="hback"><div align="left"> </div></td>
          </tr>
          <tr class="hback"> 
            <td colspan="3" class="xingmu"><div align="left">��ע����</div></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center">��ע</div></td>
            <td class="hback"> <textarea name="Content" cols="40" rows="5" id="Content"><% = RsCorpCardObj("Content")%></textarea></td>
            <td class="hback">��Ƭ�ı�ע�����200�ַ�</td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1">&nbsp;</td>
            <td colspan="2" class="hback"> <input type="submit" name="Submit" value="ȷ���ղ���Ƭ"  onClick="{if(confirm('ȷ���޸Ĵ���Ƭ��?')){this.document.UserForm.submit();return true;}return false;}"> 
              <input name="Action" type="hidden" id="Action" value="Save"> <input name="CorpCardid" type="hidden" id="CorpCardid" value="<% = Request.QueryString("CorpCardid")%>"> 
            </td>
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
	RsCorpCardObj.close
	set RsCorpCardObj = nothing
End if
Set Fs_User = Nothing
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->





