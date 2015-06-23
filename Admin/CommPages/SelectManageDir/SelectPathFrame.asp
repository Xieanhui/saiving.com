<% Option Explicit %>
<!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_Inc/Function.asp"-->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../../FS_InterFace/NS_Function.asp" -->
<%
Dim Conn
Dim TypeSql,RsTypeObj,LableSql,RsLableObj
MF_Default_Conn

Dim CurrPath
CurrPath = Request("CurrPath")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>选择路径</title>
<link href="../../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<script language="JavaScript" src="../JS/PublicJS.js"></script>
<body>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td><select onChange="ChangeFolder(this.value);" id="FolderSelectList" style="width:100%;" name="select">
        <option selected value="<% = CurrPath %>"><% = CurrPath %></option>
      </select></td>
  </tr>
  <tr> 
    <td width="70%"> <iframe id="FolderList" width="100%" height="210" frameborder="1" src="SelectPath.asp?CurrPath=<% = CurrPath %>" scrolling="auto"></iframe></td>
  </tr>
  <tr> 
    <td height="35"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="30" align="center"> <input type="button" onClick="SelectPath();" name="Submit" value=" 确 定 "> 
          </td>
          <td align="center"> <input onClick="window.close();" type="button" name="Submit3" value=" 取 消 "></td>
        </tr>
        <tr> 
          <td height="30" colspan="2" align="center"><span class="tx">提示，鼠标右键点框内，可以新建目录，删除目录，重新命名目录</span></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
<script language="JavaScript">
var SysRootDir='<% = G_VIRTUAL_ROOT_DIR %>';
function ChangeFolder(FolderName)
{
	frames["FolderList"].location='SelectPath.asp?CurrPath='+FolderName;
}
function SelectPath()
{
	var PathInfo='',TempPath='';
	if (SysRootDir!='')
	{
		TempPath=frames["FolderList"].CurrPath;
		PathInfo=TempPath.substr(TempPath.indexOf(SysRootDir)+SysRootDir.length);
	}
	else
	{
		PathInfo=frames["FolderList"].CurrPath;
	}
	if (frames["FolderList"].SelectedFolder!='')
		window.returnValue=PathInfo+'/'+frames["FolderList"].SelectedFolder;
	else
		window.returnValue=PathInfo;
	window.close();
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>





