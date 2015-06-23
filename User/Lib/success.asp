<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="strlib.asp" -->
<%
If Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")  = "" then
	response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") = "2"
End if
dim p_ErrorUrl,p_ErrCodes
p_ErrorUrl = Replace(Request.QueryString("ErrorUrl"),"''","")
p_ErrorUrl=replace(p_ErrorUrl,"***","&")
p_ErrCodes = Request.QueryString("ErrCodes")
if trim(p_ErrorUrl) = "" then
	p_ErrorUrl = "javascript:history.back(-1);"
Else
	p_ErrorUrl = 	p_ErrorUrl
End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title>操作成功</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="keywords" content="风讯,风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head>
	<body>
		<table width="60%" height="60%" border="0" align="center" cellpadding="1" cellspacing="1">
			<tr>
				<td>
					<table width="100%" height="175" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
						<form action="../CheckLogin.asp" method="post" name="myform" id="myform" onsubmit="return CheckForm();">
						<tr class="back">
							<td height="24" class="xingmu">
								<font style="font-size: 30; font-weight: bolder; color: red">√</font>操作成功
							</td>
						</tr>
						<tr class="back">
							<td width="84%" class="hback">
								<table width="100%" border="0" cellspacing="0" cellpadding="0">
									<tr>
										<td height="6">
										</td>
									</tr>
								</table>
								<span class="tx"><strong>操作描述：</strong></span>
								<table width="100%" border="0" cellspacing="0" cellpadding="0">
									<tr>
										<td height="10">
										</td>
									</tr>
								</table>
								<span class="tx"><strong></strong></span>
								<div align="left">
									<% = p_ErrCodes %>
									<table width="100%" border="0" cellspacing="0" cellpadding="0">
										<tr>
											<td height="3">
											</td>
										</tr>
									</table>
								</div>
								<li><a href="<% =  p_ErrorUrl %>">返回上一级</a> <a href="../main.asp">会员首页</a> <a href="../../">网站首页</a>
									<table width="100%" border="0" cellspacing="0" cellpadding="0">
										<tr>
											<td height="3">
											</td>
										</tr>
									</table>
								</li>
								<li><a href="http://www.foosun.cn" target="_blank">Powered by Foosun Inc.</a> <a href="http://help.foosun.net" target="_blank" style="cursor: help">帮助中心</a> <a href="http://bbs.foosun.net" target="_blank">技术论坛</a></li>
							</td>
						</tr>
						<tr class="back">
							<td height="24" class="xingmu">
								<div align="right">
									<% = p_Soft_Version %>
								</div>
							</td>
						</tr>
						</form>
					</table>
				</td>
			</tr>
		</table>
	</body>
</html>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->
<%
set Conn =nothing
set User_Conn =nothing
%>
