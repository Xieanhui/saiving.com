<% Option Explicit %>
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
Dim TypeSql,RsTypeObj,LableSql,RsLableObj,CurrPath
User_GetParm
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

If Fs_User.UserNumber = "" Or Isnull(Fs_User.UserNumber) Then
	Response.Write("<script>alert('目录发生错误。');</script>")
	Response.End()
Else
	CurrPath = Add_Root_Dir(G_USERFILES_DIR)&"/"&Fs_User.UserNumber
End if
If ReplaceExpChar(Replace(CurrPath,"/","")) = False Then
	Response.Write "<script>alert('目录发生错误。');window.location.href='javascript:history.back();';</script>"
	Response.End
End If	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<HTML>
<HEAD>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>选择图片</TITLE>
<link href="../../<% = G_USER_DIR %>/images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</HEAD>
<SCRIPT language="JavaScript" src="../Public.js"></SCRIPT>
<BODY leftmargin="0" topmargin="0">
<TABLE width="99%" border="0" align="center" cellpadding="1" cellspacing="0">
  <TR> 
    <TD height="25"><SELECT onChange="ChangeFolder(this.value);" id="FolderSelectList" style="width:100%;" name="select">
		<OPTION selected value="<% = CurrPath %>"><% = CurrPath %></OPTION>
      </SELECT></TD>
    <TD rowspan="2" align="center" valign="middle"><IFRAME id="PreviewArea" width="100%" height="315" frameborder="1" src="PreviewImage.asp"></IFRAME></TD>
  </TR>
  <TR> 
    <TD width="70%" align="center"> <IFRAME id="FolderList" width="100%" height="290" frameborder="1" src="FolderImageList.asp?CurrPath=<% = CurrPath %>&f_UserNumber=<% = Request.QueryString("f_UserNumber")%>"></IFRAME></TD>
  </TR>

  <TR> 
    <TD height="10" colspan="2"> 
      <TABLE width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <TR> 
          <TD width="80" height="10"> <DIV align="center">Url地址</DIV></TD>
          <TD><INPUT style="width:40%" type="text" name="UserUrl" id="UserUrl"> 
            <INPUT type="button" onClick="SetUserUrl();" name="Submit" value=" 确 定 "> 
            <INPUT type="button"  onClick="UpFileo();" name="Submit" value=" 上 传 " <%if  p_UpfileType="" or isnull(p_UpfileType) then Response.Write("disabled")%>> 
            <INPUT onClick="window.close();" type="button" name="Submit" value=" 取 消 "> 
          </TD>
        </TR>
        <TR> 
          <TD height="10" colspan="2" align="center"><span class="tx">在空白处点鼠标右键可以进行文件类操作,双击文件选择</span></TD>
        </TR>
      </TABLE></TD>
  </TR>
</TABLE>
</BODY>
</HTML>
<SCRIPT language="JavaScript">
function ChangeFolder(FolderName)
{
	var username="<%=session("FS_UserNumber")%>";
	frames["FolderList"].location='FolderImageList.asp?CurrPath='+FolderName+'&f_UserNumber='+username;
}
function UpFileo()
{
	OpenWindow('Frame.asp?FileName=UpFileForm.asp&Path='+frames["FolderList"].CurrPath,350,170,window);
	frames["FolderList"].location='FolderImageList.asp?f_UserNumber=<%=session("FS_UserNumber")%>&CurrPath='+frames["FolderList"].CurrPath;
}
function SetUserUrl()
{
	if (document.all.UserUrl.value=='') alert('请填写Url地址');
	else
	{
		window.returnValue=document.all.UserUrl.value;
		window.close();
	}
}
window.onunload=CheckReturnValue;
function CheckReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
function OpenWindow(Url,Width,Height,WindowObj)
{
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:no;');
	return ReturnStr;
}
</SCRIPT>