<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn,cookies_obj,cookies_str,FileObj,MF_Site_Name,GetConfigObj ,Sql,p_Soft_Version
MF_Default_Conn
'��鰲װ�ļ��Ƿ���ڣ�������ڣ�ת�򵽰�װĿ¼
Set FileObj=Server.CreateObject(G_FS_FSO)
If FileObj.FileExists(Server.MapPath("../Install/install.asp")) = True then
	response.Redirect "../Install/install.asp"
Elseif FileObj.FolderExists(Server.MapPath("../Install")) Then
	FileObj.DeleteFolder Server.MapPath("../Install"),True
End if
Set GetConfigObj = server.CreateObject (G_FS_RS)
Sql = "select  Top  1 MF_Login_style,MF_Soft_Version,MF_Site_Name From FS_MF_Config"
GetConfigObj.open sql,Conn,1,3
if Not GetConfigObj.eof then
	Session("Admin_Style_Num") = GetConfigObj(0)
	p_Soft_Version = "�汾��:V" & GetConfigObj(1)
	MF_Site_Name=GetConfigObj(2)
	GetConfigObj.close:Set GetConfigObj = nothing
Else
	p_Soft_Version = "<font color=""Red"">Err:Please configure Your Soft</font>"
	Session("Admin_Style_Num") = "1"
	MF_Site_Name = "��Ѷ��վ���ݹ���ϵͳ"
	GetConfigObj.close:Set GetConfigObj = nothing
End if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><% = MF_Site_Name %>--��̨��½--[�����IE5.5���ϰ汾�������1024*768�ֱ��ʽ����̨]</title>
<style type="text/css">

a{text-decoration: none;} /* �������»���,��Ϊunderline */
a:link {color: #232323;} /* δ���ʵ����� */
a:visited {color: #232323;} /* �ѷ��ʵ����� */
a:hover{color: #FFCC00;} /* ����������� */
a:active {color: #FFCC00;} /* ����������� */
form
{
	padding:0px;
	margin:0px;
}
body
{
	margin-top:80px;
	background-image: url(Images/log_bg.gif);
}
td ,body{
	color:#232323;
	font-size:12px;
	line-height: 18px;
}
.input{
    FONT-FAMILY: "������";
    FONT-SIZE: 12px;
	COLOR:#F3F3F3;
    text-decoration: none;
    line-height: 150%;
    background:#0099CC;
    border: solid 1px #FFFFFF;
	padding:0px;
	margin:0px;
}
.input_1{
    FONT-FAMILY: "������";
    FONT-SIZE: 12px;
	COLOR:#006699;
    text-decoration: none;
    line-height: normal;
    background:#FFFFFF;
    border: solid 1px #80CCFF;
	padding:0px;
    margin:0px;
}
</style>
<script type="text/javascript">
	window.status = "�������IE5.5���ϰ汾�������1024*768�ֱ��ʽ����̨\��ϵͳ��Maxthon,Mozilla Firefox������������õ�֧��";
</script>
</head>
<body>
<table align="center" width="486" border="0" cellspacing="3" cellpadding="0" bgcolor="#00CCFF">
	<tr>
		<td>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td colspan="2"><img src="images/i_1.gif" width="486" height="87" border="0" usemap="#Map" /></td>
				</tr>
				<tr>
					<td width="339" background="images/i_2.gif">
						<form action="CheckLogin.asp"  method="post" name="myform" id="Form1"  onsubmit="return CheckForm();">
							<table width="317" height="112" border="0" align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td width="23%" height="16" align="center">
										�û���
									</td>
									<td width="77%" height="16">
										<input name="name" type="text" class="input_1" id="name" tabindex="1" style="width:160px" value="<%=Request.Cookies("FoosunCookie")("AdminName")%>" />
										<input name="AutoGet" type="checkbox" id="AutoGet" tabindex="4" value="1"<% If Request.Cookies("FoosunCookie")("AdminName")<>"" Then Response.Write " checked" End If%> />
										��ס</td>
								</tr>
								<tr>
									<td height="15" align="center">
										�ܡ���
									</td>
									<td height="15">
										<input name="password" type="password" class="input_1" id="password" tabindex="2" style="width:160px;FONT-SIZE:12px;" />
									</td>
								</tr>
								<tr>
									<td height="19" align="center">
										������
									</td>
									<td height="19">
										<input name="VerifyCode" type="text" class="input_1" id="VerifyCode" size="10" tabindex="3" maxlength="10"/>
										������������� <img src="../Fs_Inc/vCode.asp?" onclick="this.src+=Math.random()" alt="ͼƬ�����壿������µõ���֤��" style="cursor:hand;" />
										<input name="URLs" type="hidden" id="URLs" value="<% = Request.QueryString("URLs")%>" />
									</td>
								</tr>
								<tr class="back">
									<td height="21" class="hback">&nbsp;</td>
									<td height="21" class="hback">
										<input class="input" type="submit" value="ȷ����½" name="Submit" />
										<input class="input" onclick="javascript:location.href='../'" type="button" value="������ҳ" name="Submit1" />
										<input class="input" onclick="window.location.reload()" type="button" value="ˢ�±�ҳ" name="Submit4" />
									</td>
								</tr>
							</table>
						</form>
					</td>
					<td width="147"><img src="images/i_3.gif" /></td>
				</tr>
				<tr>
					<td height="77" colspan="2" background="images/i_4.jpg"></td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<map id="Map" name="Map">
  <area shape="rect" coords="333,56,394,77" href="http://www.foosun.cn" target="_blank" alt="Foosun Inc." />
</map>
</body>
</html>
<%
set Conn =nothing
%>
<script type="text/javascript">
	CheckBrowerVersion();
	var ErrInfo = '<% = Request("ErrInfo")%>';
	function CheckBrowerVersion() {
		var MajorVer = navigator.appVersion.match(/MSIE (.)/)[1];
		var MinorVer = navigator.appVersion.match(/MSIE .\.(.)/)[1];
		var IE6OrMore = MajorVer >= 5.5 || (MajorVer >= 5.5 && MinorVer >= 5.5);
		if (!IE6OrMore) {
			alert('IE������汾̫�ͣ�ϵͳ�������������С����ȷ�������㵽���ص�ַ��');
			location.href = "http://nj.onlinedown.net/soft/17441.htm"
		}
	}
	if (ErrInfo != '') alert(ErrInfo);
	function SetFocus() {
		if (document.myform.name.value == "")
			document.myform.name.focus();
		else
			document.myform.name.select();
	}
	function CheckForm() {
		if (document.myform.name.value == "") {
			alert("�����������û�����");
			document.myform.name.focus();
			return false;
		}
		if (document.myform.password.value == "") {
			alert("�������������룡");
			document.myform.password.focus();
			return false;
		}
		if (document.myform.VerifyCode.value == "") {
			alert("������������֤�룡");
			document.myform.VerifyCode.focus();
			return (false);
		}
	}
</script>