<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<%
Dim CurrPath
CurrPath = Request.QueryString("CurrPath")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ѡ���ļ�</title>
</head>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet">
<body>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td><select onChange="ChangeFolder(this.value);" id="FolderSelectList" style="width:100%;" name="select">
        <option selected value="<% = CurrPath %>"><% = CurrPath %></option>
      </select></td>
  </tr>
  <tr> 
    <td width="70%"> <iframe id="FolderList" width="100%" height="300" frameborder="1" src="FolderList.asp?CurrPath=<% = CurrPath %>"></iframe></td>
  </tr>
  <tr> 
    <td height="35"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="30" align="center"> 
            <input type="button" onClick="SelectFile();" name="Submit" value=" ȷ �� "> 
          </td>
          <td align="center"> 
            <input onClick="window.close();" type="button" name="Submit3" value=" ȡ �� "></td>
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
	frames["FolderList"].location='FolderList.asp?CurrPath='+FolderName;
}
function SelectFile()
{
	if (frames["FolderList"].FileName!='')
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
		window.returnValue=PathInfo+'/'+frames["FolderList"].FileName;
		window.close();
	}
	else
		alert('��ѡ��ѡ���ļ�');
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>





