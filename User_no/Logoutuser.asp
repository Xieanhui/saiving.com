<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="../FS_Inc/Md5.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
If Request.Form("action")="submit" Then 
	Dim t_return
	t_return = Fs_User.DelUser(NoSqlHack(Fs_User.UserNumber),md5(NoSqlHack(Request.Form("pwd")),16))
	If t_return Then 
		Call Fs_User.out()
		strShowErr = "<li>ע���ɹ�</li><li> ���������Ѿ�������</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../../")
		Response.end
	Else
		strShowErr = "<li>ע��ʧ��</li><li> ��������������������</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End If 
End If 

%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-ע���ҵĻ�Ա�ʸ�</title>
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
            <a href="main.asp">��Ա��ҳ</a> &gt;&gt; ע���ҵĻ�Ա�ʸ�</td>
        </tr>
      </table> 
      
        
      
      <table width="98%" height="98" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="form1" method="post" action="">
          <tr> 
            <td class="hback"> ע�⣺ע����Ա�ʸ�����ڱ�ϵͳ��һ�����Ͻ��ᱻ���٣�<br>
              <br>
              <br>
              ���������룺 
              <input type="password" name="pwd" value=""> 
              <input name="action" type="hidden" id="action" value="submit"> <input type="submit" name="Submit" value="ע���ʻ�" onClick="if (confirm('ע����������Ͻ��᳹����ʧ�������ָܻ�\nȷ��ע����')) return true;else return false;"></td>
          </tr>
        </form>
      </table></td>
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





