<% Option Explicit %>
<!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_Inc/Function.asp"-->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../../FS_InterFace/NS_Function.asp" -->
<%
Dim Conn
MF_Default_Conn
MF_Session_TF 
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<title>选择图片</title>
	<link href="../../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css"
		rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0">
	<table width="99%" border="0" align="center" cellpadding="1" cellspacing="0">
		<tr>
			<td height="25">
				<select onchange="ChangeFolder(this.value);" id="FolderSelectList" style="width: 100%;"
					name="select">
				</select>
			</td>
			<td rowspan="2" align="center" valign="middle">
				<iframe id="PreviewArea" width="100%" height="315" frameborder="1" src="PreviewImage.asp">
				</iframe>
			</td>
		</tr>
		<tr>
			<td width="70%" align="center">
				<iframe id="FolderList" width="100%" height="290" frameborder="1" src="FolderImageList.asp">
				</iframe>
			</td>
		</tr>
		<tr>
			<td height="10" colspan="2">
				<table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
					<tr>
						<td width="80" height="10">
							<div align="center">
								Url地址</div>
						</td>
						<td>
							<input style="width: 40%" type="text" name="UserUrl" id="UserUrl">
							<input type="button" onclick="SetUserUrl();" name="Submit" value=" 确 定 ">
							<input type="button" onclick="UpFileo();" name="Submit" value=" 上 传 " <%if not MF_Check_Pop_TF("MF025") then Response.Write"disabled"%>>
							<input onclick="WhenCancel();" type="button" name="Submit" value=" 取 消 ">
						</td>
					</tr>
					<tr>
						<td height="10" colspan="2" align="center">
							<span class="tx">在空白处点鼠标右键可以进行文件类操作,双击文件选择</span>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</body>
</html>

<script type="text/javascript">
	function ChangeFolder(FolderName) {
		frames["FolderList"].location = 'FolderImageList.asp?CurrPath=' + FolderName;
	}
	function UpFileo() {
		OpenWindow('Frame.asp?FileName=UpFileForm.asp&Path=' + frames["FolderList"].CurrPath, 350, 170, window);
		frames["FolderList"].location = 'FolderImageList.asp?CurrPath=' + frames["FolderList"].CurrPath;
	}
	function SetUserUrl() {
		if (document.all.UserUrl.value == '') alert('请填写Url地址');
		else {
			window.returnValue = document.all.UserUrl.value;
			try {
				window.opener.SetUrl(document.all.UserUrl.value);
			} catch (ex) {
			}
			window.close();
		}
	}
	function WhenCancel() {
		window.returnValue = "";
		window.close();
	}
	window.onunload = CheckReturnValue;
	function CheckReturnValue() {
		if (typeof (window.returnValue) != 'string') window.returnValue = '';
	}
	function OpenWindow(Url, Width, Height, WindowObj) {
		var ReturnStr = showModalDialog(Url, WindowObj, 'dialogWidth:' + Width + 'pt;dialogHeight:' + Height + 'pt;status:no;help:no;scroll:no;');
		return ReturnStr;
	}
</script>

