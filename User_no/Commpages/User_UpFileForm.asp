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

Dim TypeSql,RsTypeObj,LableSql,RsLableObj,objPath
objPath = Request("Path")
If ReplaceExpChar(p_ShowPath) = False Then
	Response.Write "<script>alert('目录名不规范');window.Close();</script>"
	Response.End
End If
If objPath="" then
	objPath=Add_Root_Dir("/") & G_USERFILES_DIR & "/" &Fs_User.UserNumber
Else
	objPath=Add_Root_Dir("/") & G_USERFILES_DIR & "/" &Fs_User.UserNumber&"/"&objPath
End If
Set p_FSO = Server.CreateObject(G_FS_FSO)
if p_FSO.FolderExists(Server.MapPath(objPath))=false then
	Response.Write "<script>alert('目录不存在');window.Close();</script>"
	Response.End()
end if

Function ReplaceExpChar(Str)
	Dim RegEx,StrRs,S_Str,ReturnV,HaveV
	S_Str = Str & ""
	ReturnV = True
	Set RegEx = New RegExp
	RegEx.IgnoreCase = True
	RegEx.Global=True
	RegEx.Pattern = "([^a-zA-Z0-9])"
	Set StrRs = RegEx.ExeCute(S_Str)
	Set RegEx = Nothing
	For Each HaveV In StrRs
		If Instr(S_Str,HaveV) <> 0 Then
			ReturnV = False
			Exit For
		End IF	
	Next
	ReplaceExpChar = ReturnV
End Function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>上传文件</title>
</head>
<link href="../../<% = G_USER_DIR %>/images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<body onselectstart="return false;" topmargin="0" leftmargin="0">
<form name="FileForm" method="post" enctype="multipart/form-data" action="UpFileSave.asp">
 <table width="98%" align="center" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td height="26" align="center" width="33%">请输入上传文件个数：</td>
      
    <td width="33%"> 
      <input name="FilesNum" type="text" value="4" size="10"> 
      <input type="button" name="Submit42" value="设定" onClick="ChooseOption();"></td>
    <td width="33%"> 
      <input name="chkAddWaterMark" type="checkbox" id="chkAddWaterMark" value="1">
      添加水印</td>
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
                <div align="center">选择自动命名格式</div></td>
            </tr>
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
                  <input type="submit" id="BtnSubmit" onClick="PromptInfo();" name="Submit" value=" 确 定 ">
                  <input name="Path" value="<% = objPath %>" type="hidden" id="Path">
                </div></td>
              <td><div align="center"> 
                  <input type="reset" id="ResetForm" name="Submit3" value=" 重 填 ">
                </div></td>
              <td><div align="center"> 
                  <input onClick="top.close();" type="button" name="Submit2" value=" 关 闭 ">
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





