<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
User_GetParm
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-我的帐户</title>
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
            <a href="main.asp">会员首页</a> &gt;&gt; 我的帐户</td>
          </tr>
        </table>
        
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td width="15%" class="hback_1"><div align="center"><strong>我的金币</strong></div></td>
          <td width="21%"><div align="left"><%=formatNumber(Fs_User.NumFS_Money,2,-1)%>&nbsp;<%=p_MoneyName%></div></td>
          <td width="64%"><a href="change_pm.asp?action=changepoint">兑换成积分?</a>　　　<a href="../help?lable=changePoint" target="_blank" style="cursor:help">帮助</a></td>
        </tr>
        <tr class="hback"> 
          <td height="32"  class="hback_1"><div align="center"><strong>我的积分</strong></div></td>
          <td class="hback"><div align="left"><%=FormatNumber(Fs_User.NumIntegral,2,-1)%>点 　　　 </div></td>
          <td class="hback"><a href="change_pm.asp?action=changemoney">兑换成金币?</a>　　　<a href="../help?lable=changeMoney" target="_blank" style="cursor:help">帮助</a></td>
        </tr>
      </table>
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td width="100%" height="32" class="hback"> <a href="onlinepay.asp">冲值帐户(在线银行)</a>　　　<a href="card.asp">冲值帐户(点卡冲值)</a>　　　<a href="history.asp">交易记录</a>　　　<span class="top_user"><a href="award/award.asp">抽奖/兑奖</a></span></td>
        </tr>
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
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->





