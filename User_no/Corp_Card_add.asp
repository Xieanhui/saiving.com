<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
Dim p_F_UserNumber
User_GetParm
p_F_UserNumber = NoSqlHack(Request.QueryString("UserNumber"))
if Request.Form("Action") = "Save" then
	Dim str_f_UserNumber,str_Content
	str_f_UserNumber = NoSqlHack(Request.Form("UserNumber"))
	str_Content = NoHtmlHackInput(NoSqlHack(Request.Form("Content")))
	if str_f_UserNumber ="" then
			strShowErr = "<li>����Ļ�Ա���</li>"
			Call ReturnError(strShowErr,"")
	End if
	if len(str_Content) >200 then
			strShowErr = "<li>��ע���ܴ���200���ַ�</li>"
			Call ReturnError(strShowErr,"")
	End if
			Dim RsaddObj,RsaddTFObj
			Set RsaddTFObj = server.CreateObject(G_FS_RS)
			RsaddTFObj.open "select F_UserNumber From FS_ME_CorpCard where F_UserNumber='"& str_f_UserNumber &"'",User_Conn,1,1
			if Not RsaddTFObj.eof then
				strShowErr = "<li>����Ƭ�Ѿ��ղ�</li><li>���Ѿ��������Ƭ�ղ����󣬶Է���û��������</li>"
				Call ReturnError(strShowErr,"")
			End if
			Set RsaddObj = server.CreateObject(G_FS_RS)
			RsaddObj.open "select * From FS_ME_CorpCard where 1=0",User_Conn,1,3
			RsaddObj.addnew
			RsaddObj("UserNumber") = Fs_User.UserNumber
			RsaddObj("F_UserNumber") = str_f_UserNumber
			RsaddObj("Content") = str_Content
			RsaddObj("AddTime") = now
			if p_isPassCard = 0 then
				RsaddObj("isLock") = 0
			Else
				RsaddObj("isLock") = 1
			End if
			RsaddObj.update
			RsaddObj.close
			set RsaddObj = nothing
			if p_isPassCard = 0 then
				strShowErr = "<li>��Ƭ�ղسɹ�</li>"
			Else
				strShowErr = "<li>���Ѿ���Է����������Ƭ�����󣬵ȴ�ͨ��</li>"
			End if
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
Else
	if p_F_UserNumber = "" then
			strShowErr = "<li>����Ĳ���</li>"
			Call ReturnError(strShowErr,"")
	Else
			Dim RsCorpCardObj,RsCorpUserObj
			Set RsCorpUserObj = server.CreateObject(G_FS_RS)
			RsCorpUserObj.open "select islock,UserNumber From FS_ME_Users where UserNumber='"& p_F_UserNumber&"' and islock=0",User_Conn,1,1
			if RsCorpUserObj.eof then
				strShowErr = "<li>�Ҳ����û���Ϣ</li>"
				Call ReturnError(strShowErr,"")
			End if
			Set RsCorpCardObj = server.CreateObject(G_FS_RS)
			RsCorpCardObj.open "select  C_Name,UserNumber From FS_ME_CorpUser where UserNumber='"& p_F_UserNumber&"'",User_Conn,1,1
			iF RsCorpCardObj.eof then
				strShowErr = "<li>�Ҳ�����ҵ��Ϣ</li>"
				Call ReturnError(strShowErr,"")
			End if
			Dim str_C_Name
			str_C_Name = RsCorpCardObj("C_Name")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-�����Ƭ</title>
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
        <form name="UserForm" method="post" action="" onSubmit="return CheckForm();">
          <tr class="hback"> 
            <td width="12%" class="hback_1"><div align="center" class="tx">�û���</div></td>
            <td width="40%" class="hback"><div align="left"> 
                <input name="UserNumber" type="text" id="UserNumber" value="<% = p_F_UserNumber%>" size="30" ReadOnly>
              </div></td>
            <td width="48%" class="hback"><div align="left"> </div></td>
          </tr><tr class="hback"> 
            <td width="12%" class="hback_1"><div align="center" class="tx">��˾����</div></td>
            <td width="40%" class="hback"><div align="left"> 
                <input name="UserNumber_c" type="text" id="UserNumber" value="<% = str_C_Name%>" size="40" ReadOnly>
              </div></td>
            <td width="48%" class="hback"><div align="left"> </div></td>
          </tr>
          <tr class="hback"> 
            <td colspan="3" class="xingmu"><div align="left">��ע����</div></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center">��ע</div></td>
            <td class="hback"> <textarea name="Content" cols="40" rows="5" id="Content"><% = str_C_Name%></textarea></td>
            <td class="hback">��Ƭ�ı�ע�����200�ַ�</td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1">&nbsp;</td>
            <td colspan="2" class="hback"> <input type="submit" name="Submit" value="ȷ���ղ���Ƭ"> 
              <input name="Action" type="hidden" id="Action" value="Save">
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
	RsCorpUserObj.close
	set RsCorpUserObj = nothing
	RsCorpCardObj.close
	set RsCorpCardObj = nothing
	End if
End if
Set Fs_User = Nothing
%>
<script language="JavaScript" type="text/javascript">
function CheckForm()
{
	if(document.UserForm.UserName.value=="")
	{
		alert("����д�û���!");
		document.UserForm.UserName.focus();
		return false;
	}
	if(document.UserForm.RealName.value=="")
	{
		alert("����д��ע����!");
		document.UserForm.RealName.focus();
		return false;
	}
	}
</script>

<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->





