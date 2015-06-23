<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<%
Dim Conn
MF_Default_Conn
MF_Session_TF
if not MF_Check_Pop_TF("MF025") then Err_Show
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>上传文件</title>
</head>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<body onselectstart="return false;" topmargin="0" leftmargin="0">
<form name="FileForm" method="post" enctype="multipart/form-data" action="UpFileSave.asp">
 <table width="98%" align="center" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td height="26" align="center" width="33%">请输入上传文件个数：</td>
      
    <td width="33%"> 
      <input name="FilesNum" type="text" value="4" size="10"> 
      <input type="button" name="Submit42" value="设定" onClick="ChooseOption();"></td>
    <td width="33%">&nbsp;</td>
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
        <td width="30%" valign="top"><br> <fieldset style="width:100%;">
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td height="20" > 
                <div align="center">自动命名</div></td>
            </tr>
            <tr> 
              <td height="20"> 
                <div align="left"> 
                  <input type="radio" name="AutoReName" value="0" checked>
                  不自动命名</div></td>
            </tr>
			<tr> 
              <td height="20"> 
                <div align="left"> 
                  <input type="radio" name="AutoReName" value="1">
                  &quot; 副件&quot;+文件名</div></td>
            </tr>
            <tr> 
              <td height="20"> 
                <div align="left"> 
                  <input type="radio" name="AutoReName" value="2">
                  &quot; 1&quot;+文件名</div></td>
            </tr>
            <tr> 
              <td height="20"> 
                <div align="left"> 
                  <input type="radio" name="AutoReName" value="3">
                  20041101122333</div></td>
            </tr>
          </table>
          </fieldset></td>
      </tr>
      <tr> 
        <td colspan="2"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td> <div align="center"> 
                  <input type="submit" id="BtnSubmit" onClick="PromptInfo();" name="Submit" value=" 确 定 ">
                  <input name="Path" value="<% = Request("Path") %>" type="hidden" id="Path">
                </div></td>
              <td><div align="center"> 
                  <input type="reset" id="ResetForm" name="Submit3" value=" 重 填 ">
                </div></td>
              <td><div align="center"> 
                  <input onClick="dialogArguments.location.reload();window.close();" type="button" name="Submit2" value=" 关 闭 ">
                </div></td>
            </tr>
          </table></td>
      </tr>
  </table>
</div>
</form>
<div id="LayerPrompt" style="position:absolute; z-index:1; left: 112px; top: 28px; background-color: #00CCFF; layer-background-color: #00CCFF; border: 1px none #000000; width: 254px; height: 63px; visibility: hidden;"> 
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <tr> 
      <td><div align="right">请稍等，正在上传文件</div></td>
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
	   Optionstr = Optionstr+'<tr><td>&nbsp;文&nbsp;件&nbsp;'+i+'</td><td>&nbsp;<input type="file" accept="html" size="20" name="File'+i+'">&nbsp;</td></tr>';
	   }
	Optionstr = Optionstr+'</table>';  
    document.all.FilesList.innerHTML = Optionstr;
  }
ChooseOption();
</script>





