<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<%
'Response.Write(Session.Timeout)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>网站内容管理系统--管理后台</title><%
If G_SESSION_TIME_OUT=1 Then
	Response.Write "<meta http-equiv=""Refresh"" content=""300"" />"
End If
%>
<meta name="Keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统" />
<link href="images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
</head>
<body scroll="no">
<script type="text/javascript" src="../FS_Inc/wz_tooltip.js"></script>
<table width="100%" height="51" border="0" cellpadding="0" cellspacing="0" background="images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>_bg_01.gif">
	<tr>
		<td valign="top">
			<table width="100%" height="39" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td width="167" valign="top"><img src="images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>_logo.gif" width="167" height="39"></td>
					<td valign="bottom">
						<table width="100%" height="25" border="0" cellpadding="2" cellspacing="2">
							<tr>
								<td width="87" align="right" valign="middle">
									<div align="left"><img src="Images/changeskin.gif" width="86" height="12" border="0" usemap="#Map"></div>
								</td>
								<td align="right" valign="middle">
									<div align="center">
										<%if MF_Check_Pop_TF("MF_Templet") then %>
										<A href="Templets_List.asp" onMouseOver="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';this.T_FIX=[330,10];return escape('<div align=\'left\'>模板管理<br> </div>')" target="ContentFrame" class="nav_l2">模板</A><span class="nav_l2">┆</span>
										<%end if%>
										<%if MF_Check_Pop_TF("MF_Style") then %>
										<A href="Label/All_Label_style.asp" onMouseOver="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';this.T_FIX=[380,10];return escape('<div align=\'left\'>样式管理<br> </div>')" target="ContentFrame" class="nav_l2">样式管理</A><span class="nav_l2">┆</span>
										<%end if%>
										<%if MF_Check_Pop_TF("MF_sPublic") then %>
										<A href="Label/All_Label_Stock.asp" target="ContentFrame" onMouseOver="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';this.T_FIX=[450,10];return escape('<div align=\'left\'>标签管理<br> </div>')" class="nav_l2">标签库</A><span class="nav_l2">┆</span>
										<%end if%>
										<%if MF_Check_Pop_TF("MF_Public") then %>
										<A href="Sys_Public.asp" target="ContentFrame" onMouseOver="this.T_BGCOLOR='#FFFFC4';this.T_FONTCOLOR='#000000';this.T_BORDERCOLOR='#000000';this.T_TEMP=2000;this.T_WIDTH='120px';this.T_FIX=[520,10];return escape('<div align=\'left\'>发布管理<br> </div>')" class="nav_l2">发布管理</A>
										<%end if%>
									</div>
								</td>
								<td align="right" valign="middle">
									<div align="center"><a href="updatecls.asp" target="ContentFrame" class="nav_l2">更新缓存</a><span class="nav_l2">┆</span><a href="changpassword.asp" target="ContentFrame" class="nav_l2">修改密码</a><span class="nav_l2">┆</span><a href="../Help?Label=Directory" target="_blank" class="nav_l2">帮助</a><span class="nav_l2">┆</SPAN><a href="http://help.foosun.net" target="_blank" class="nav_l2">在线帮助</a><span class="nav_l2">┆</span><a href="http://bbs.foosun.net" target="_blank" class="nav_l2">论坛</a><span class="nav_l2">┆</span><a href="Loginout.asp" target="_top" class="nav_l2">退出</a></div>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<map name="Map">
	<area shape="rect" coords="1,0,13,13" href="Changeskin.asp?Style_num=1" target="_top" alt="默认风格">
	<area shape="rect" coords="18,0,32,15" href="Changeskin.asp?Style_num=2" target="_top" alt="银色风格">
	<area shape="rect" coords="37,0,49,11" href="Changeskin.asp?Style_num=3" target="_top" alt="蓝色海洋">
	<area shape="rect" coords="54,0,68,12" href="Changeskin.asp?Style_num=4" target="_top" alt="浪漫咖啡">
	<area shape="rect" coords="72,0,87,12" href="Changeskin.asp?Style_num=5" target="_top" alt="青青河草">
</map>
</body>
</html>





