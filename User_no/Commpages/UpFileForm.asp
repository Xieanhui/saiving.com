<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim TypeSql,RsTypeObj,LableSql,RsLableObj
if instr(Request("Path"),Fs_User.UserNumber)=0 then:response.Write("���棺��Ҫ������˵��ļ�Ŀ¼��"):response.end:end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ϴ��ļ�</title>
</head>
<link href="../../<% = G_USER_DIR %>/images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<body onselectstart="return false;" topmargin="0" leftmargin="0">
<form name="FileForm" method="post" enctype="multipart/form-data" action="UpFileSave.asp">
 <table width="98%" align="center" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td height="26" align="center" width="33%">�������ϴ��ļ�������</td>
      
    <td width="33%"> 
      <input name="FilesNum" type="text" value="4" size="10"> 
      <input type="button" name="Submit42" value="�趨" onClick="ChooseOption();"></td>
    <td width="33%"> 
      <input name="chkAddWaterMark" type="checkbox" id="chkAddWaterMark" value="0">
      ���ˮӡ</td>
    </tr>
</table>

<div align="center">
  <table width="98%" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td> <div align="center"> 
            <table width="90%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td height="30" id="FilesList"> 
				</td>
              </tr>
            </table>
            </div>
		</td>
        <td width="40%" valign="top"><br> <fieldset style="width:100%;">
          <table width="90%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td height="20" > 
                <div align="center">ѡ���Զ�������ʽ</div></td>
            </tr>
            <!--<tr> 
              <td height="20"> 
                <div align="left"> 
                  <input name="AutoReName" type="radio" value="0" checked>
                  ���Զ�����</div></td>
            </tr>
			<tr> 
              <td height="20"> 
                <div align="left"> 
                  <input type="radio" name="AutoReName" value="1">
                  &quot; ����&quot;+�ļ���</div></td>
            </tr>-->
            <tr> 
              <td height="20"> 
                <div align="left"> 
                  <input type="radio" name="AutoReName" value="2" checked="checked">
                 2004_11_01_12_23_33</div></td>
            </tr>
            <tr> 
              <td height="20"> 
                <div align="left"> 
                  <input name="AutoReName" type="radio" value="3">
                  20041101122333</div></td>
            </tr>
          </table>
          </fieldset></td>
      </tr>
      <tr> 
        <td colspan="2"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td> <div align="center"> 
                  <input type="submit" id="BtnSubmit" onClick="PromptInfo();" name="Submit" value=" ȷ �� ">
                  <input name="Path" value="<% = Request("Path") %>" type="hidden" id="Path">
                </div></td>
              <td><div align="center"> 
                  <input type="reset" id="ResetForm" name="Submit3" value=" �� �� ">
                </div></td>
              <td><div align="center"> 
                  <input onClick="dialogArguments.location.reload();top.close();" type="button" name="Submit2" value=" �� �� ">
                </div></td>
            </tr>
          </table></td>
      </tr>
  </table>
</div>
</form>
<div id="LayerPrompt" style="position:absolute; z-index:1; left: 112px; top: 28px; background-color: #CCCCCC; layer-background-color: #CCCCCC; border: 1px none #000000; width: 254px; height: 63px; visibility: hidden;"> 
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <tr> 
      <td><div align="right">���Եȣ������ϴ��ļ�</div></td>
	  <td width="35%"><div align="left"><font id="ShowInfoArea" size="+1"></font></div></td>
    </tr>
  </table>
</div>
</body>
</html>
<%
Set Conn = Nothing
%>
<script language="JavaScript">
var ForwardShow=true;
function ShowPromptInfo()
{
	var TempStr=ShowInfoArea.innerText;
	if (ForwardShow==true)
	{
		if (TempStr.length>4) ForwardShow=false;
		ShowInfoArea.innerText=TempStr+'.';
		
	}
	else
	{
		if (TempStr.length==1) ForwardShow=true;
		ShowInfoArea.innerText=TempStr.substr(0,TempStr.length-1);
	}
}

function PromptInfo()
{
	var FilesNum = document.all.FilesNum.value;
	var obj;
	if(FilesNum=='')
		FilesNum=4;
	for(var i=1;i<=FilesNum;i++){
	   obj = eval("document.FileForm.File" + i);
	   obj.readOnly = true;
	}
	//document.FileForm.BtnSubmit.readOnly=true;
	document.FileForm.ResetForm.disabled=true;
	LayerPrompt.style.visibility='visible';
	window.setInterval('ShowPromptInfo()',600)
	return true;
}
function ChooseOption()
 {
  var FilesNum = document.all.FilesNum.value;
  if (FilesNum=='')
  	FilesNum=4;
  var i,Optionstr;
	  Optionstr = '<table width="100%" border="0" cellspacing="5" cellpadding="0">';
  for (i=1;i<=FilesNum;i++)
      {
	   Optionstr = Optionstr+'<tr><td>&nbsp;��&nbsp;��&nbsp;'+i+'</td><td>&nbsp;<input type="file" accept="html" size="20" name="File'+i+'">&nbsp;</td></tr>';
	   }
	Optionstr = Optionstr+'</table>';  
    document.all.FilesList.innerHTML = Optionstr;
  }
ChooseOption();
</script>





