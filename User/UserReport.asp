<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
if NoSqlHack(request.QueryString("ToUserNumber")) = Fs_User.UserNumber then
		strShowErr = "<li>�����Լ��ٱ��Լ�</li>"
		Call ReturnError(strShowErr,"")
End if
Dim ReturnValue_Report
ReturnValue_Report = Fs_User.GetFriendName(NoSqlHack(request.QueryString("ToUserNumber")))
if Request.Form("Action") = "Save" then
	If Trim(Request.Form("F_UserName"))="" then
		strShowErr = "<li>����д���ٱ�����</li>"
		Call ReturnError(strShowErr,"")
	End if
	If Trim(Request.Form("ReportType"))="" then
		strShowErr = "<li>����дҪ�ٱ�������</li>"
		Call ReturnError(strShowErr,"")
	End if
	If Trim(Request.Form("Content"))="" then
		strShowErr = "<li>����д�ٱ�����</li>"
		Call ReturnError(strShowErr,"")
	End if
	If Len(Trim(Request.Form("Content")))>1000 then
		strShowErr = "<li>�ٱ��������ܳ���1000���ַ�</li>"
		Call ReturnError(strShowErr,"")
	End if
	Dim GetUserNumberValue,UserTFobj
	GetUserNumberValue = Fs_User.GetFriendNumber(NoSqlHack(request.Form("F_UserName")))
	Set UserTFobj = User_Conn.execute("Select UserID From FS_ME_Users Where UserNumber ='"& NoSqlHack(GetUserNumberValue) &"'")
	if UserTFobj.eof then
		strShowErr = "<li>�Ҳ�����Ҫ�ٱ����û���</li>"
		Call ReturnError(strShowErr,"")
	Else
		Dim RsRepObj
		Set RsRepObj = server.CreateObject(G_FS_RS)
		RsRepObj.open "select * From FS_ME_Report where 1=0",User_Conn,1,3
		RsRepObj.addnew
		RsRepObj("UserNumber") =  Fs_User.UserNumber
		RsRepObj("F_UserNumber") = GetUserNumberValue
		RsRepObj("addtime") = now
		RsRepObj("Content") = NoSqlHack(NoHtmlHackInput(Request.Form("Content")))
		RsRepObj("isLock") = 0
		RsRepObj("ReportType") = NoSqlHack(request.Form("ReportType"))
		RsRepObj.update
		RsRepObj.close
		set RsRepObj = nothing
		strShowErr = "<li>�ٱ��ɹ�</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End if
End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-�ٱ�</title>
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
            <a href="main.asp">��Ա��ҳ</a> &gt;&gt; �ٱ��û�</td>
        </tr>
      </table> 
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="UserForm" method="post" action="">
          <tr class="hback"> 
            <td width="16%" class="hback_1"><div align="right"><strong>�û���</strong></div></td>
            <td class="hback"> <input name="F_UserName" type="text" id="F_UserName" value="<% = ReturnValue_Report %>" size="26" maxlength="50">
              ����д�û��� 
              <div align="center"></div></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1" > <div align="right"><strong>����</strong></div></td>
            <td class="hback"><select name="ReportType" id="ReportType">
                <option selected value="">ѡ��ٱ�����</option>
                <option value="0">ƭ��</option>
                <option value="1">���</option>
                <option value="2">��������</option>
                <option value="3">�Ƿ�����</option>
                <option value="4">����</option>
              </select></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1" ><div align="right"><strong>�ٱ�����</strong></div></td>
            <td class="hback"><textarea name="Content" cols="60" rows="10" id="Content"></textarea>
              ����1000���ַ�</td>
          </tr>
          <tr class="hback"> 
            <td class="hback"><div align="center"> </div></td>
            <td class="hback"><input name="Action" type="hidden" id="Action2" value="Save"> 
              <input type="button" name="SubmitButton" value="�ύ�ٱ�"  onClick="{if(confirm('ȷ���ύ�ٱ���?\n����Ҫ�����ľٱ�����!')){this.document.UserForm.submit();return true;}return false;}">
              �� 
              <input type="reset" name="Submit3" value="������д">
              �� </td>
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
Set Fs_User = Nothing
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->





