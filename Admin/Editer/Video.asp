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
<HTML><HEAD><TITLE>插入视频文件</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../../Editer/ModeWindow.css">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<script language="JavaScript">
function OK(){
  var str1="";
  var strurl=document.VideoForm.url.value;
  if (!IsExt(strurl,'avi|wmv|asf|mpg|mp3|wma'))
  {
	  alert('文件类型不对，请重新选择！');
	  document.VideoForm.url.focus();
	  return;
  }
  if (strurl==""||strurl=="http://")
  {
  	alert("请先输入视频文件地址，或者上传视频文件！");
	document.VideoForm.url.focus();
	return false;
  }
  else
  {
	str1=str1+"<embed src="+document.VideoForm.url.value+" width="+document.VideoForm.width.value+" height="+document.VideoForm.height.value+" autostart=true loop=true>"
	str1=str1+"</embed>"
	window.returnValue = str1+"$$$"+document.VideoForm.UpFileName.value;
    window.close();
  }
}

function IsExt(url,opt)
{  
	var sTemp; 
	var b=false; 
	var s=opt.toUpperCase().split("|");  
	for (var i=0;i<s.length ;i++ ) 
	{ 
		sTemp=url.substr(url.length-s[i].length-1); 
		sTemp=sTemp.toUpperCase();
		s[i]="."+s[i];
		if (s[i]==sTemp)
		{
			b=true;
			break;
		}
	}
	return b;
}

function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>
</head>
<BODY bgColor=menu topmargin=15 leftmargin=15 >
<form name="VideoForm" method="post" action="">
  <table width=100% border="0" cellpadding="0" cellspacing="2">
    <tr>
      <td> <FIELDSET align=left>
        <LEGEND align=left>视频文件参数</LEGEND>
        <TABLE border="0" cellpadding="0" cellspacing="3">
          <TR>
            <TD >地址：
              <INPUT name="url" id=url  value="http://" size=30>
              <input type="button" name="Button" value="选择视频文件" onClick="var TempReturnValue=OpenWindow('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=/<% = G_UP_FILES_DIR %>',500,290,window);if (TempReturnValue!='') document.VideoForm.url.value='<% = DirectoryRoot %>'+TempReturnValue;" class=Anbutc> 
            </td>
          </TR>
          <TR>
            <TD >宽度：
              <INPUT name="width" id=width  ONKEYPRESS="event.returnValue=IsDigit();" value=352 size=14 maxlength="4"> 
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;高度：
              <INPUT id=height ONKEYPRESS="event.returnValue=IsDigit();" value=288 size=14 maxlength="4"></TD>
          </TR>
          <TR>
            <TD align=center>支持格式为：avi、wmv、mpg、asf、mp3、wma</TD>
          </TR>
        </TABLE>
        </fieldset></td>
      <td width=80 align="center"><input name="cmdOK" type="button" id="cmdOK" value="  确定  " onClick="OK();"> 
        <br> <br> <input name="UpFileName" type="hidden" id="UpFileName2" value="None"> 
        <input name="cmdCancel" type=button id="cmdCancel" onClick="window.close();" value='  取消  '></td>
    </tr>
  </table>
</form>
</body>
</html>