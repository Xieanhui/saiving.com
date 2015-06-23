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
	Dim p_NickName,p_BothYear,p_picsizew,p_picsizeh
	p_NickName = NoSqlHack(Replace(Request.Form("NickName"),"''",""))
	p_BothYear = NoSqlHack(Replace(Request.Form("BothYear"),"''",""))
	p_picsizew = NoSqlHack(Replace(Request.Form("HeadPicSizew"),"''",""))
	p_picsizeh = NoSqlHack(Replace(Request.Form("HeadPicSizeh"),"''",""))
	if trim(p_NickName) ="" then 
		strShowErr = "<li>����д�û��ǳ�</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Elseif isdate(p_BothYear) = false then
		strShowErr = "<li>����д����������Ч�ġ���ȷ��ʽΪ��1980-7-7</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Elseif  isNumeric(p_picsizew) =false  then
		strShowErr = "<li>ͷ��������Ĳ�����Ч����</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Elseif isNumeric(p_picsizeh) =false then
		strShowErr = "<li>ͷ��߶�����Ĳ�����Ч����</li>>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Elseif   cint(p_picsizew)>200 then
		strShowErr = "<li>ͷ���Ȳ��ܳ���200px</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Elseif  cint(p_picsizeh)>200 then
		strShowErr = "<li>ͷ��߶Ȳ��ܳ���200px</li"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Elseif len(Request.Form("SelfIntro"))>50 then
		strShowErr = "<li>���ҽ��ܲ��ܳ���50���ַ�</li"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		Dim RsSaveIObj
		Set RsSaveIObj = server.CreateObject(G_FS_RS)
		RsSaveIObj.open "select  UserID,isLock,UserName,RealName,GroupID,Integral,UserNumber,BothYear,SelfIntro,isOpen,Certificate,CertificateCode,Vocation,HeadPic,NickName,Mobile,CloseTime,IsCorporation,isMessage,Email,sex,safeCode,UserLoginCode,HeadPicsize,OnlyLogin,UserFavor,IsMarray From FS_ME_Users where UserNumber = '"& Fs_User.UserNumber &"'",User_Conn,1,3
		RsSaveIObj("NickName") = p_NickName
		RsSaveIObj("RealName") = NoSqlHack(Replace(Request.Form("RealName"),"''",""))
		RsSaveIObj("sex") = NoSqlHack(Replace(Request.Form("sex"),"''",""))
		RsSaveIObj("Vocation")  = NoSqlHack(Replace(Request.Form("Vocation"),"''",""))
		RsSaveIObj("HeadPic")  = NoSqlHack(Replace(Request.Form("HeadPic"),"''",""))
		RsSaveIObj("HeadPicSize")  = p_picsizew&","&p_picsizeh
		RsSaveIObj("BothYear")  = p_BothYear
		RsSaveIObj("IsMarray")  = NoSqlHack(Request.Form("IsMarray"))
		RsSaveIObj("isopen")  = NoSqlHack(Request.Form("isopen"))
		RsSaveIObj("SelfIntro")  = NoSqlHack(Request.Form("SelfIntro"))
		RsSaveIObj("UserFavor")  = NoSqlHack(Request.Form("UserFavor"))
		RsSaveIObj.update
		RsSaveIObj.close
		set RsSaveIObj = nothing
		strShowErr = "<li>���������޸ĳɹ�!</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../myinfo.asp")
		Response.end
	End if
Else
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title>��ӭ�û�<%=Fs_User.UserName%>����<%=GetUserSystemTitle%>-�ҵ�����</title>
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
            <a href="main.asp">��Ա��ҳ</a> &gt;&gt; �ҵ�����</td>
        </tr>
      </table> 
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="form1" method="post" action="">
          <tr class="hback"> 
            <td width="16%" class="hback_1"><div align="center"><strong>�� �� ��</strong></div></td>
            <td width="35%" class="hback"><input name="UserName" type="text" id="UserName" value="<% = Fs_User.UserName%>" size="26" readonly></td>
            <td width="9%" class="hback_1"><div align="center"><strong>�û����</strong></div></td>
            <td width="40%" class="hback"><input name="UserNumber" type="text" id="UserNumber" value="<% = Fs_User.UserNumber%>" size="26" readonly></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center"><strong>�û��ǳ�</strong></div></td>
            <td class="hback"><input name="NickName" type="text" id="NickName" value="<% = Fs_User.NickName%>" size="26" maxlength="20"></td>
            <td class="hback_1"><div align="center"><strong>��ʵ����</strong></div></td>
            <td class="hback"><input name="RealName" type="text" id="RealName" value="<% = Fs_User.RealName%>" size="26" maxlength="20"></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center"><strong>�û��Ա�</strong></div></td>
            <td class="hback"><select name="sex" id="sex">
                <option value="0" <%if Fs_User.Sex = 0 then response.Write("selected")%>>��</option>
                <option value="1" <%if Fs_User.Sex = 1 then response.Write("selected")%>>Ů</option>
              </select></td>
            <td class="hback_1"><div align="center"><strong>�û�ְҵ</strong></div></td>
            <td class="hback"><input name="Vocation" type="text" id="Vocation" value="<% = Fs_User.Vocation%>" size="26" maxlength="30"> 
            </td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center"><strong>�û�ͷ��</strong></div></td>
            <td class="hback"><input name="HeadPic" type="text" id="HeadPic" value="<% = Fs_User.HeadPic%>" size="26" maxlength="250">
            <img  src="Images/upfile.gif" width="44" height="22" onClick="OpenWindowAndSetValue('CommPages/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,320,window,document.form1.HeadPic);" style="cursor:hand;"></td>
            <td class="hback_1"><div align="center"><strong>ͷ��ߴ�</strong></div></td>
            <td class="hback"> <%
			Dim arr_HeadPicsize
			If Not IsNull(Fs_User.HeadPicsize) then
				arr_HeadPicsize = split(Fs_User.HeadPicsize,",")
			End if
			%>
              �� 
              <input name="HeadPicsizew" type="text" value="<%If Not IsNull(Fs_User.HeadPicsize) then Response.write(arr_HeadPicsize(0))%>" size="5" maxlength="3">
              px,�� 
              <input name="HeadPicsizeh" type="text" value="<%If Not IsNull(Fs_User.HeadPicsize) then Response.write(arr_HeadPicsize(1))%>" size="5" maxlength="3">
              px</td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center"><strong>��������</strong></div></td>
            <td class="hback"><input name="BothYear" type="text" id="BothYear" value="<% = Fs_User.BothYear%>" size="26" maxlength="10">
              ��ʽ��1980-7-7</td>
            <td class="hback_1">&nbsp;</td>
            <td class="hback">&nbsp;</td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center"><strong>֤������</strong></div></td>
            <td class="hback"> <%
			if  Fs_User.PaperType = 0 then
				Response.Write("���֤")
			Elseif   Fs_User.PaperType = 1 then
				Response.Write("��ʻ֤")
			Elseif   Fs_User.PaperType = 2 then
				Response.Write("ѧ��֤")
			Elseif   Fs_User.PaperType = 3 then
				Response.Write("����֤")
			Elseif   Fs_User.PaperType = 4 then
				Response.Write("����")
			Else
				Response.Write("δ֪֤��")
			End if
			%></td>
            <td class="hback_1"><div align="center"><strong>֤������</strong></div></td>
            <td class="hback"><% = Fs_User.PaperTypeCode%></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center"><strong>�Ƿ���</strong></div></td>
            <td class="hback"><select name="IsMarray" id="IsMarray">
                <option value="0" <%if Fs_User.IsMarray = 0 then response.Write("selected")%>>����</option>
                <option value="1" <%if Fs_User.IsMarray = 1 then response.Write("selected")%>>�ѻ�</option>
                <option value="2" <%if Fs_User.IsMarray = 2 then response.Write("selected")%>>δ��</option>
              </select></td>
            <td class="hback_1"><div align="center"><strong>��������</strong></div></td>
            <td class="hback"><input type="radio" name="isopen" value="1" <%if Fs_User.isopen = 1 then response.Write("checked")%>>
              ���� 
              <input type="radio" name="isopen" value="0" <%if Fs_User.isopen = 0 then response.Write("checked")%>>
              �ر�</td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center"><strong>���ҽ���<br>
                (ǩ��)</strong></div></td>
            <td class="hback"><textarea name="SelfIntro" cols="30" rows="5" id="SelfIntro"><% = Fs_User.SelfIntro%></textarea></td>
            <td class="hback_1"><div align="center"><strong>���˰���</strong></div></td>
            <td class="hback"> <textarea name="UserFavor" cols="30" rows="5" id="UserFavor"><% = Fs_User.UserFavor%></textarea></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center"><strong>��½ IP</strong></div></td>
            <td class="hback"> <% = Fs_User.LastLoginIP%> </td>
            <td class="hback_1"><div align="center"><strong>����½</strong></div></td>
            <td class="hback"> <% = Fs_User.LastLoginTime%> </td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center"><strong>��������</strong></div></td>
            <td class="hback"> <%
			  if Fs_User.CloseTime ="3000-1-1" then
					Response.Write("û����")
			  Else
					Response.Write Fs_User.CloseTime
			  End if
			  %> </td>
            <td class="hback_1"><div align="center"><strong>ע������</strong></div></td>
            <td class="hback"> <% = Fs_User.RegTime%> </td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center"><strong>�㡡����</strong></div></td>
            <td class="hback"> <% = Fs_User.NumIntegral%> </td>
            <td class="hback_1"><div align="center"><strong>�𡡡���</strong></div></td>
            <td class="hback"> <% = Fs_User.NumFS_Money%> </td>
          </tr>
          <tr class="hback"> 
            <td colspan="4" class="hback"><div align="center"> 
                <input name="Action" type="hidden" id="Action" value="Save">
                <input type="submit" name="Submit" value="��������"   onClick="{if(confirm('ȷ������д����Ϣ��?')){this.document.form1.submit();return true;}return false;}">
                �� 
                <input type="reset" name="Submit3" value="������д">
              </div></td>
          </tr>
          <tr class="hback"> 
            <td colspan="4" class="hback"> <div align="center"> </div></td>
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
End if
Set Fs_User = Nothing
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->





