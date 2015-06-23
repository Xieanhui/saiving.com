<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn,User_Conn,tmp_type,strShowErr
MF_Default_Conn
MF_User_Conn
MF_Session_TF
if not MF_Check_Pop_TF("SS_site") then Err_Show
if not MF_Check_Pop_TF("SS001") then Err_Show

dim tmp_str
tmp_str=Replace("/"&G_VIRTUAL_ROOT_DIR &"/","//","/")
%>
<html>
<head>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>流量统计代码</title>
</head>
<body topmargin="2" leftmargin="2">
<table width="98%" height="90"  border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr  class="xingmu"> 
    <td colspan="2" valign="middle" class="xingmu"><strong>流量统计代码调用</strong></td>
  </tr>
  <tr  class="hback"> 
    <td width="16%" valign="middle" class="hback"> <div align="center">有图标</div></td>
    <td width="84%" valign="middle"><SPAN class=small2><FONT face="宋体">&lt;script 
      language=&quot;JavaScript&quot; src=&quot;<% = tmp_str %>stat/index.asp?code=1&quot; type=&quot;text/JavaScript&quot;&gt;&lt;/script&gt;</FONT></SPAN></td>
  </tr>
  <tr  class="hback"> 
    <td valign="middle"  class="hback"> <div align="center">无图标</div></td>
    <td  class="hback"><FONT face="宋体">&lt;script 
      language=&quot;JavaScript&quot; src=&quot;<% = tmp_str %>stat/index.asp?code=0&quot; type=&quot;text/JavaScript&quot;&gt;&lt;/script&gt;</FONT></td>
  </tr>
  <tr  class="hback"> 
    <td valign="middle"  class="hback"> <div align="center">文字统计</div></td>
    <td valign="middle"  class="hback"><FONT face="宋体">&lt;script language=&quot;JavaScript&quot; src=&quot;<% = tmp_str %>stat/index.asp?code=2&quot; type=&quot;text/JavaScript&quot;&gt;&lt;/script&gt;</FONT></td>
  </tr>
</table>
<div align="center"></div>
</body>
</html>






