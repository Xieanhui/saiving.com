<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
Dim strNews,RsShownSQL,RsShownObj
Dim p_Newsid,RsObj
p_Newsid = NoSqlHack(Request.QueryString("Newsid"))
User_Conn.execute("update FS_ME_News set hits=hits+1 where NewsID="& CintStr(p_Newsid) &"")
strNews = NoSqlHack(Request.QueryString("NewsId"))
RsShownSQL = "select  NewsID,Title,Content,Addtime,GroupID,NewsPoint,isLock,hits From FS_ME_News where NewsID="&CintStr(strNews)
Set RsShownObj = server.CreateObject(G_FS_RS)
RsShownObj.Open RsShownSQL,User_Conn,1,3
if RsShownObj.eof then
	strShowErr = "<li>错误的参数</li><li>找不到记录</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
Else
   '判断公告浏览权限
   '---------
   if RsShownObj("isLock") = 1 then
		strShowErr = "<li>错误的参数</li><li>此公告已经被管理员锁定</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
   End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-浏览公告</title>
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
            <a href="main.asp">会员首页</a> &gt;&gt; <a href="Callboard.asp">会员公告</a>&gt;&gt;&gt;浏览公告</td>
        </tr>
      </table> 
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td height="24" class="xingmu"><span class="menu_left">
            <% = RsShownObj("title")%>
            </span></td>
        </tr>
        <tr class="hback"> 
          <td height="170" valign="top" class="hback"> 
            <% = RsShownObj("Content")%>
          </td>
        </tr>
        <tr class="hback">
          <td class="hback"><div align="right">添加日期: 
              <% = RsShownObj("Addtime")%>&nbsp;|&nbsp; 阅读: 
              <% = RsShownObj("hits")%>
              次 </div></td>
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
End if
RsShownObj.close
Set RsShownObj = nothing
Set Fs_User = Nothing
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->





