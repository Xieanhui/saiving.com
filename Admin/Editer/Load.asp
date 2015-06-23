<%@language=vbscript codepage=936 %>
<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Cls_SysConfig.asp" -->
<%
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
Dim Conn
MF_Default_Conn
MF_Session_TF 

Dim DirectoryRoot
Dim SysConfig
Set SysConfig = New Cls_SysConfig
DirectoryRoot = SysConfig.domain
Set SysConfig = Nothing
%>

<HTML><HEAD><title>插入附件</title>
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
  	alert("请先输入附件地址！");
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
          <td>打开方式：</td>
          <td><select name="target">
            <option value="_self">当前窗口</option>
            <option value="_blank">新窗口</option>
          </select>
          </td>
        </tr>
        <tr> 
          <td align="right">网页地址：</td>
          <td><input name="url" id=url value='http://' size=30>&nbsp;
		  <input type="button" name="Button" value="选择附件" onClick="var TempReturnValue=OpenWindow('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=/<% = G_UP_FILES_DIR %>',500,290,window);setFileInfo(TempReturnValue);" class=Anbutc></td>
        </tr>
      </table></td>
    <td width=1 align="center">
<div align="center" class="Separator"></div></td>
    <td width=100 align="center" valign="top"> 
      <input name="cmdOK" type="button" id="cmdOK" value="  确定  " onClick="OK();"> 
      <br> <br>
      <input name="cmdCancel" type=button id="cmdCancel" onClick="window.close();" value='  取消  '></td>
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
