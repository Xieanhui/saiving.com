<% Option Explicit %>
<!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_Inc/Function.asp"-->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../../FS_InterFace/NS_Function.asp" -->
<%
Dim Conn
Dim TypeSql,RsTypeObj,LableSql,RsLableObj
MF_Default_Conn
MF_Session_TF
Dim CurrPath
CurrPath = NoSqlHack(Request("CurrPath"))
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>选择路径</title>
<style type="text/css">
<!--
.PreviewStyle {
	border: 2px outset #CCCCCC;
}
 BODY   {border: 0; margin: 0; background: buttonface; cursor: default; font-family:宋体; font-size:9pt;}
 BUTTON {width:5em}
 TABLE  {font-family:宋体; font-size:9pt}
 P      {text-align:center}
-->
</style>
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
    <td width="70%"> <iframe id="FolderList" width="100%" height="300" frameborder="1" src="SelectPath.asp?CurrPath=<% = CurrPath %>"></iframe></td>
  </tr>
  <tr> 
    <td height="35"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="30" align="center"> 
            <input type="button" onClick="SelectPath();" name="Submit" value=" 确 定 "> 
          </td>
          <td align="center"> 
            <input onClick="window.close();" type="button" name="Submit3" value=" 取 消 "></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
<script language="JavaScript">
var SysRootDir='<% = SysRootDir %>';
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





