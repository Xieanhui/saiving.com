<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
Dim str_CurrPath
str_CurrPath = Replace("/"&G_VIRTUAL_ROOT_DIR &"/"&G_USERFILES_DIR&"/"&Session("FS_UserNumber"),"//","/")
If Request.Form("Action") = "Save" then
	if trim(Request.Form("C_Logo")) ="" then 
		strShowErr = "<li>������ͼƬ</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		Dim RsSaveIObj
		Set RsSaveIObj = server.CreateObject(G_FS_RS)
		RsSaveIObj.open "select CorpID,UserNumber,C_Logo From FS_ME_CorpUser where UserNumber = '"& Fs_User.UserNumber &"' and CorpID="&CintStr(NoSqlHack(Request.Form("ID"))),User_Conn,1,3
		RsSaveIObj("C_Logo")  = NoSqlHack(Replace(Request.Form("C_Logo"),"''",""))
		RsSaveIObj.update
		RsSaveIObj.close
		set RsSaveIObj = nothing
		set conn=nothing
		set user_conn=nothing
		set fs_user=nothing
		strShowErr = "<li>LOGO�޸ĳɹ�!</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Corp_Logo.asp")
		Response.end
	End if
Else
		Dim RsCorpObj
		Set RsCorpObj = server.CreateObject(G_FS_RS)
		RsCorpObj.open "select  CorpID,C_logo,isLockCorp From FS_ME_CorpUser where UserNumber = '"& Fs_User.UserNumber &"'",User_Conn,1,3
		if RsCorpObj.eof then
			RsCorpObj.close
			set RsCorpObj=nothing
			set conn=nothing
			set user_conn=nothing
			set fs_user=nothing
			strShowErr = "<li>�Ҳ�����ҵ����</li>"
			Call ReturnError(strShowErr,"")
		End if
		if RsCorpObj("isLockCorp") = 1 then
			set RsCorpObj=nothing
			set conn=nothing
			set user_conn=nothing
			set fs_user=nothing
			strShowErr = "<li>������ҵ���ݻ�û���ͨ��</li>"
			Call ReturnError(strShowErr,"")
		End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-��ҵLogo</title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,��Ѷ,��Ѷ��վ���ݹ���ϵͳ,��Ѷϵͳ,��Ѷ����ϵͳ,��Ѷ�̳�,��Ѷb2c,����ϵͳ,CMS,�����ռ�,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head>
<body>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td>
      <!--#include file="top.asp" -->
    </td>
  </tr>
</table>
<table width="98%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  
    <tr class="back"> 
      <td   colspan="2" class="xingmu" height="26"> <!--#include file="Top_navi.asp" --> </td>
    </tr>
    <tr class="back"> 
      <td width="18%" valign="top" class="hback"> <div align="left"> 
          <!--#include file="menu.asp" -->
        </div></td>
      <td width="82%" valign="top" class="hback"><table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td class="hback"><strong>λ�ã�</strong><a href="../">��վ��ҳ</a> &gt;&gt; 
            <a href="main.asp">��Ա��ҳ</a> &gt;&gt; ��ҵͼƬ��־</td>
        </tr>
      </table> 
      
        
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="UserForm" method="post" action="">
          <tr class="hback"> 
            <td colspan="2" class="xingmu"> <strong> </strong></td>
          </tr>
          <tr class="hback"> 
            <td width="16%" class="hback_1"><div align="center"><strong>��˾ͼƬLogo</strong></div></td>
            <td class="hback"> <input name="C_logo" type="text" id="C_logo" value="<% = RsCorpObj("C_logo")%>" size="40" maxlength="200">
              <img  src="Images/upfile.gif" width="44" height="22" onClick="OpenWindowAndSetValue('CommPages/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,320,window,document.UserForm.C_logo);" style="cursor:hand;"> 
              <div align="center"></div></td>
          </tr>
          <tr class="hback"> 
            <td colspan="2" class="hback"><div align="center"> 
                <input name="ID" type="hidden" id="ID" value="<% = RsCorpObj("CorpID")%>">
                <input name="Action" type="hidden" id="Action" value="Save">
                <input type="submit" name="Submit" value="��������"  onClick="{if(confirm('ȷ�ϱ�����?')){this.document.UserForm.submit();return true;}return false;}">
                �� 
                <input type="reset" name="Submit3" value="������д">
                �� </div></td>
          </tr>
          <tr class="hback"> 
            <td colspan="2" class="hback"> <div align="center"> </div></td>
          </tr>
        </form>
      </table>
       </td>
    </tr>
    <tr class="back"> 
      <td height="20"  colspan="2" class="xingmu"> <div align="left"> 
          <!--#include file="Copyright.asp" -->
        </div></td>
    </tr>
 
</table>
</body>
</html>
<%
	RsCorpObj.close
	set RsCorpObj = nothing
End if
Set Fs_User = Nothing
%>
<script language="JavaScript" type="text/JavaScript">
	function OpenWindowAndSetValue(Url,Width,Height,WindowObj,SetObj)
	{
		var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:no;');
		if (ReturnStr!='') SetObj.value=ReturnStr;
		return ReturnStr;
	}
</script>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->





