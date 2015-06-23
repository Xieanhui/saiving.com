<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<%
dim obj_mf_sys_obj,MF_Domain,MF_Site_Name,tmp_c_path
set obj_mf_sys_obj = Conn.execute("select top 1 MF_Domain,MF_Site_Name from FS_MF_Config")
if obj_mf_sys_obj.eof then
	strShowErr = "<li>找不到主系统配置信息！</li>"
	Response.Redirect("../lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
else
	MF_Domain = obj_mf_sys_obj("MF_Domain")
	MF_Site_Name = obj_mf_sys_obj("MF_Site_Name")
end if
obj_mf_sys_obj.close:set obj_mf_sys_obj = nothing
tmp_c_path =MF_Domain &"/"&G_VIRTUAL_ROOT_DIR
%>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=GetUserSystemTitle%>-抽奖</title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=GB2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,风讯,风讯网站内容管理系统,风讯系统,风讯新闻系统,风讯商城,风讯b2c,新闻系统,CMS,域名空间,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="javascript" src="../../FS_Inc/prototype.js"></script>
<script language="javascript" src="../../FS_Inc/CheckJs.js"></script>
<script language="javascript" src="../../FS_Inc/PublicJS.js"></script>
<script language="javascript" src="../../FS_Inc/coolWindowsCalendar.js"></script>
</head>
<body id="mainContainer">
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td>
      <!--#include file="../top.asp" -->
    </td>
  </tr>
</table>
<table width="98%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
    <tr class="back"> 
      <td   colspan="2" class="xingmu" height="26"> <!--#include file="../Top_navi.asp" --> </td>
    </tr>
    <tr class="back"> 
      <td width="18%" valign="top" class="hback"> <div align="left"> 
          <!--#include file="../menu.asp" -->
        </div></td>
      <td width="82%" valign="top" class="hback"><table width="99%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td class="hback"><strong>位置：</strong><a href="../../">网站首页</a> &gt;&gt; 
            <a href="../main.asp">会员首页</a> &gt;&gt; <a href="award.asp">抽奖</a></td>
        </tr>
        <tr class="hback">
          <td class="hback"><a href="#" onClick="gotoPanel(1)"><img src="../images/award.gif" alt="积分抽奖" border="0">积分抽奖</a>&nbsp;&nbsp;<a href="#" onClick="gotoPanel(2)"><img src="../images/award.gif" alt="积分兑换" border="0">积分兑换</a>&nbsp;&nbsp;<a href="#" onClick="gotoPanel(3)"><img src="../images/award.gif" alt="积分问答" border="0">积分问答</a></td>
        </tr>
        <tr class="hback">
          <td class="hback"><img src="../images/Currentaward.gif" alt="我的中奖记录" width="16" height="16"><a href="#" onClick="gotoPanel(4)">我的中奖记录</a></td>
        </tr>
        <tr class="hback">
          <td class="hback"></td>
        </tr>
		<tr>
		<td align="center" class="hback">
			<iframe src="awardPane.asp" frameborder="0" id="AwardPane" width="100%" height="600"></iframe>
		</td>
		</tr>
      </table>
	  </td>
    </tr>
      </table>
	  </td>
    </tr>
    <tr class="back"> 
      <td height="20"  colspan="2" class="xingmu"> <div align="left"> 
          <!--#include file="../Copyright.asp" -->
        </div></td>
    </tr>
</table>
</body>
</html>
<script language="javascript">
<!--
function gotoPanel(panel)
{
	switch(panel)
	{
		case 2:url="awardChange.asp";break;
		case 3:url="awardAnswer.asp";break;
		case 4:url="myAwardRecord.asp";break
		default:url="awardPane.asp";break;
	}
	$("AwardPane").src=url;
}
-->
</script>
<%
Set Conn=nothing
Set User_Conn=nothing
Set Fs_User = Nothing
%>





