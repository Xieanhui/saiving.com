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

<HTML><HEAD><TITLE>插入RealPlay文件</TITLE>
<link rel="stylesheet" type="text/css" href="../../Editer/ModeWindow.css">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<script language="JavaScript">
function OK()
{
  var str1="";
  var strurl=document.RmForm.url.value;
  if (!IsExt(strurl,'rm|ra|ram|mp3|rmvb|avi'))
  {
	  alert('文件类型不对，请重新选择！');
	  document.RmForm.url.focus();
	  return;
  }
  if (strurl==""||strurl=="http://")
  {
  	alert("请先输入RealPlay文件地址，或者上传RealPlay文件！");
	document.RmForm.url.focus();
	return false;
  }
  else
  {
		str1="<object classid=clsid:cfcdaa03-8be4-11cf-b84b-0020afbbccfa width="+document.RmForm.width.value+" height="+document.RmForm.height.value+">"
		str1=str1+"<param name=src value="+document.RmForm.url.value+">"
		str1=str1+"<param name=console value=clip1><param name=controls value=imagewindow>"
		str1=str1+"<param name=autostart value=true>"
		str1=str1+"</object><object classid=clsid:cfcdaa03-8be4-11cf-b84b-0020afbbccfa height=60 width="+document.RmForm.width.value+">"
		str1=str1+"<param name=src value="+document.RmForm.url.value+">"
		str1=str1+"<param name=controls value=controlpanel><param name=console value=clip1></object>"  
		window.returnValue = str1+"$$$"+document.RmForm.UpFileName.value;
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
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<BODY bgColor=menu topmargin=15 leftmargin=15 >
<form name="RmForm" method="post" action="">
  <table width=100% border="0" cellpadding="0" cellspacing="2">
    <tr>
      <td> <FIELDSET align=left>
        <LEGEND align=left>RealPlay文件参数</LEGEND>
        <TABLE border="0" cellpadding="0" cellspacing="3">
          <TR>
            <TD >地址： <INPUT name="url" id=url  value="http://" size=30>
              <input type="button" name="Button" value="选择文件" onClick="var TempReturnValue=OpenWindow('/<%=G_USER_DIR %>/CommPages/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,290,window);if (TempReturnValue!='') document.RmForm.url.value='<% = DirectoryRoot %>'+TempReturnValue;" class=Anbutc> 
            </td>
          </TR>
          <TR>
            <TD >宽度：
              <INPUT name="width" id=width ONKEYPRESS="event.returnValue=IsDigit();" value=500 size=7 maxlength="4" > 
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;高度：
              <INPUT name="height" id=height ONKEYPRESS="event.returnValue=IsDigit();" value=300 size=7 maxlength="4"></TD>
          </TR>
          <TR>
            <TD align=center>支持格式为：rm、ra、ram、rmvb、mp3、avi</TD>
          </TR>
        </TABLE>
        </fieldset></td>
      <td width=80 align="center"><input name="cmdOK" type="button" id="cmdOK" value="  确定  " onClick="OK();"> 
        <br> <br> <input name="cmdCancel" type=button id="cmdCancel" onClick="window.close();" value='  取消  '> 
        <input name="UpFileName" type="hidden" id="UpFileName2" value="None"></td>
    </tr>
  </table>
</form>
</body>
</html>