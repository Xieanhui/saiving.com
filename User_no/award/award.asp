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
	strShowErr = "<li>�Ҳ�����ϵͳ������Ϣ��</li>"
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
<title><%=GetUserSystemTitle%>-�齱</title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=GB2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,��Ѷ,��Ѷ��վ���ݹ���ϵͳ,��Ѷϵͳ,��Ѷ����ϵͳ,��Ѷ�̳�,��Ѷb2c,����ϵͳ,CMS,�����ռ�,asp,jsp,asp.net,SQL,SQL SERVER" />
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
          <td class="hback"><strong>λ�ã�</strong><a href="../../">��վ��ҳ</a> &gt;&gt; 
            <a href="../main.asp">��Ա��ҳ</a> &gt;&gt; <a href="award.asp">�齱</a></td>
        </tr>
        <tr class="hback">
          <td class="hback"><a href="#" onClick="gotoPanel(1)"><img src="../images/award.gif" alt="���ֳ齱" border="0">���ֳ齱</a>&nbsp;&nbsp;<a href="#" onClick="gotoPanel(2)"><img src="../images/award.gif" alt="���ֶһ�" border="0">���ֶһ�</a>&nbsp;&nbsp;<a href="#" onClick="gotoPanel(3)"><img src="../images/award.gif" alt="�����ʴ�" border="0">�����ʴ�</a></td>
        </tr>
        <tr class="hback">
          <td class="hback"><img src="../images/Currentaward.gif" alt="�ҵ��н���¼" width="16" height="16"><a href="#" onClick="gotoPanel(4)">�ҵ��н���¼</a></td>
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





