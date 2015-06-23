<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<%
dim p_ErrorUrl,p_ErrCodes
p_ErrorUrl = Replace(Request.QueryString("ErrorUrl"),"''","")
p_ErrCodes = Request.QueryString("ErrCodes")
if trim(p_ErrorUrl) = "" then
	p_ErrorUrl = "javascript:history.back();"
Else
	p_ErrorUrl = 	p_ErrorUrl
End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title>错误信息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="keywords" content="风讯,风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<link href="images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head>
	<body>
		<table width="60%" height="60%" border="0" align="center" cellpadding="1" cellspacing="0">
			<tr>
				<td>
					<table width="100%" height="182" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
						<form action="../../../CheckLogin.asp" method="post" name="myform" id="myform" onsubmit="return CheckForm();">
						<tr class="back">
							<td width="100%" height="24" class="xingmu">
								错误的参数
							</td>
						</tr>
						<tr class="back">
							<td height="130" class="hback">
								<table width="100%" border="0" cellspacing="0" cellpadding="0">
									<tr>
										<td height="6">
										</td>
									</tr>
								</table>
								<span class="tx"><strong>出错原因：</strong></span>
								<table width="100%" border="0" cellspacing="0" cellpadding="0">
									<tr>
										<td width="12%" height="10">
											<div align="center">
												<span class="tx"><font style="font-size: 30px; face: Arial, Helvetica, sans-serif"><strong>× </strong></font></span>
											</div>
										</td>
										<td width="88%">
											<div align="left">
												<% = p_ErrCodes %>
												<li>请仔细阅读帮助文件？ <a href="<% =  p_ErrorUrl %>">返回上一级</a> </li>
												<li><a href="http://www.foosun.cn" target="_blank">Powered by Foosun Inc.</a> <a href="http://help.foosun.net" target="_blank" style="cursor: help">帮助中心</a> <a href="http://bbs.foosun.net" target="_blank">技术论坛</a></li>
										</td>
									</tr>
								</table>
								<span class="tx"></span>
								<div align="left">
								</div>
							</td>
						</tr>
						<tr class="back">
							<td height="24" class="xingmu">
								<div align="right">
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
