<%@language=vbscript codepage=936 %>
<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<!--#include file="../../FS_Inc/Cls_SysConfig.asp" -->
<%
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
Dim DirectoryRoot,str_CurrPath
Dim SysConfig
Set SysConfig = New Cls_SysConfig
DirectoryRoot = SysConfig.domain
str_CurrPath = Replace("/"&G_VIRTUAL_ROOT_DIR &"/"&G_USERFILES_DIR&"/"&Session("FS_UserNumber"),"//","/")
%>

<HTML><HEAD><title>���븽��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../../Editer/ModeWindow.css">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<style>
.Separator 
{
	BORDER-RIGHT: buttonhighlight 1px solid;
	FONT-SIZE: 0px;
	BORDER-LEFT: buttonshadow 1px solid;
	cursor: default;
	height: 50px;
	width: 1px;
	top: 10px;
}
</style>
<script language="JavaScript">
function OK(){
  var str1="";
  var strurl=url.value;
  if (strurl==""||strurl=="http://")
  {
  	alert("�������븽����ַ��");
	url.focus();
	return false;
  }
  else
  {
  str1="<a href='"+url.value+"' target=\""+document.all.target.value+"\">"+url.value+"</a>"
  window.returnValue = str1
  window.close();
  }
}
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
function setFileInfo(f_TempReturnValue) 
{
	if (f_TempReturnValue!='')
	{
		document.all.url.value='http://'+document.domain+f_TempReturnValue;
	}
}
</script>
</head>
<body bgcolor=menu topmargin=15 leftmargin=15 >
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center"> 
      <LEGEND align=left></LEGEND> <table border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td>�򿪷�ʽ��</td>
          <td><select name="target">
            <option value="_self">��ǰ����</option>
            <option value="_blank">�´���</option>
          </select>
          </td>
        </tr>
        <tr> 
          <td align="right">��ҳ��ַ��</td>
          <td><input name="url" id=url value='http://' size=30>&nbsp;
		  <input type="button" name="Button" value="ѡ�񸽼�" onClick="var TempReturnValue=OpenWindow('/<%=G_USER_DIR %>/CommPages/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,290,window);setFileInfo(TempReturnValue);" class=Anbutc></td>
        </tr>
      </table></td>
    <td width=1 align="center">
<div align="center" class="Separator"></div></td>
    <td width=100 align="center" valign="top"> 
      <input name="cmdOK" type="button" id="cmdOK" value="  ȷ��  " onClick="OK();"> 
      <br> <br>
      <input name="cmdCancel" type=button id="cmdCancel" onClick="window.close();" value='  ȡ��  '></td>
  </tr>
</table>
</body>
</html>
<script language="JavaScript">
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>
