<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
If Request.Form("Action") = "Save" then
		Dim RsSaveIObj
		Set RsSaveIObj = server.CreateObject(G_FS_RS)
		RsSaveIObj.open "select  UserID,isLock,UserName,RealName,GroupID,Integral,LoginNum,RegTime, LastLoginTime,LastLoginIP,UserNumber,FS_Money,ConNumber,UserID,HomePage,BothYear,Tel,MSN,QQ,Corner,Province,City,Address,PostCode,PassQuestion,SelfIntro,isOpen,Certificate,CertificateCode,Vocation,HeadPic,NickName,Mobile,CloseTime,IsCorporation,isMessage,Email,sex,safeCode,UserLoginCode,HeadPicsize,OnlyLogin,UserFavor,IsMarray From FS_ME_Users where UserNumber = '"& Fs_User.UserNumber &"'",User_Conn,1,3
		RsSaveIObj("tel") = NoSqlHack(Replace(Request.Form("tel"),"''",""))
		RsSaveIObj("mobile") = NoSqlHack(Replace(Request.Form("mobile"),"''",""))
		RsSaveIObj("homepage") = NoSqlHack(Replace(Request.Form("homepage"),"''",""))
		RsSaveIObj("Province")  = NoSqlHack(Replace(Request.Form("Province"),"''",""))
		RsSaveIObj("city")  = NoSqlHack(Replace(Request.Form("city"),"''",""))
		RsSaveIObj("address")  =NoSqlHack(Replace(Request.Form("address"),"''",""))
		RsSaveIObj("postcode")  = NoSqlHack(Replace(Request.Form("postcode"),"''",""))
		If Request.Form("qq")<>"" Then
			RsSaveIObj("qq")  = NoSqlHack(Request.Form("qq"))
		End If
		RsSaveIObj("msn")  = NoSqlHack(Request.Form("msn"))
		RsSaveIObj.update
		RsSaveIObj.close
		set RsSaveIObj = nothing
		strShowErr = "<li>��ϵ���������޸ĳɹ�!</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../MyContact.asp")
		Response.end
Else
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-�ҵ���ϵ����</title>
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
            <a href="main.asp">��Ա��ҳ</a> &gt;&gt; ��ϵ����</td>
        </tr>
      </table> 
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="form1" method="post" action="">
          <tr class="hback"> 
            <td width="16%" class="hback_1"><div align="center"><strong>�绰</strong></div></td>
            <td width="35%" class="hback"><input name="Tel" type="text" id="Tel" value="<% = Fs_User.Tel%>" size="26" maxlength="20"></td>
            <td width="9%" class="hback_1"><div align="center"><strong>�ƶ��绰</strong></div></td>
            <td width="40%" class="hback"><input name="Mobile" type="text" id="Mobile" value="<% = Fs_User.Mobile%>" size="26" maxlength="50"></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center"><strong>�����ʼ�</strong></div></td>
            <td class="hback"><input name="Email" type="text" id="Email" value="<% = Fs_User.Email%>" size="26" maxlength="100" readonly></td>
            <td class="hback_1"><div align="center"><strong>������ҳ</strong></div></td>
            <td class="hback"><input name="homepage" type="text" id="homepage" value="<% = Fs_User.homepage%>" size="26" maxlength="200"></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center"><strong>ʡ��</strong></div></td>
            <td class="hback"><select name="Province" size=1 id="Province">
                <option value="">��ѡ�񡡡�</option>
                <option value="����" <% If Fs_User.Province ="����" then response.Write("selected")%>>����</option>
                <option value="�Ϻ�" <% If Fs_User.Province ="�Ϻ�" then response.Write("selected")%>>�Ϻ�</option>
                <option value="�Ĵ�" <% If Fs_User.Province ="�Ĵ�" then response.Write("selected")%>>�Ĵ�</option>
                <option value="���" <% If Fs_User.Province ="���" then response.Write("selected")%>>���</option>
                <option value="����" <% If Fs_User.Province ="����" then response.Write("selected")%>>����</option>
                <option value="����" <% If Fs_User.Province ="����" then response.Write("selected")%>>����</option>
                <option value="����" <% If Fs_User.Province ="����" then response.Write("selected")%>>����</option>
                <option value="�㶫" <% If Fs_User.Province ="�㶫" then response.Write("selected")%>>�㶫</option>
                <option value="����" <% If Fs_User.Province ="����" then response.Write("selected")%>>����</option>
                <option value="����" <% If Fs_User.Province ="����" then response.Write("selected")%>>����</option>
                <option value="����" <% If Fs_User.Province ="����" then response.Write("selected")%>>����</option>
                <option value="����" <% If Fs_User.Province ="����" then response.Write("selected")%>>����</option>
                <option value="�ӱ�" <% If Fs_User.Province ="�ӱ�" then response.Write("selected")%>>�ӱ�</option>
                <option value="����" <% If Fs_User.Province ="����" then response.Write("selected")%>>����</option>
                <option value="������" <% If Fs_User.Province ="������" then response.Write("selected")%>>������</option>
                <option value="����" <% If Fs_User.Province ="����" then response.Write("selected")%>>����</option>
                <option value="����" <% If Fs_User.Province ="����" then response.Write("selected")%>>����</option>
                <option value="����" <% If Fs_User.Province ="����" then response.Write("selected")%>>����</option>
                <option value="����" <% If Fs_User.Province ="����" then response.Write("selected")%>>����</option>
                <option value="����" <% If Fs_User.Province ="����" then response.Write("selected")%>>����</option>
                <option value="����" <% If Fs_User.Province ="����" then response.Write("selected")%>>����</option>
                <option value="���ɹ�" <% If Fs_User.Province ="���ɹ�" then response.Write("selected")%>>���ɹ�</option>
                <option value="����" <% If Fs_User.Province ="����" then response.Write("selected")%>>����</option>
                <option value="�ຣ" <% If Fs_User.Province ="�ຣ" then response.Write("selected")%>>�ຣ</option>
                <option value="ɽ��" <% If Fs_User.Province ="ɽ��" then response.Write("selected")%>>ɽ��</option>
                <option value="ɽ��" <% If Fs_User.Province ="ɽ��" then response.Write("selected")%>>ɽ��</option>
                <option value="����" <% If Fs_User.Province ="����" then response.Write("selected")%>>����</option>
                <option value="����" <% If Fs_User.Province ="����" then response.Write("selected")%>>����</option>
                <option value="�½�" <% If Fs_User.Province ="�½�" then response.Write("selected")%>>�½�</option>
                <option value="����" <% If Fs_User.Province ="����" then response.Write("selected")%>>����</option>
                <option value="�㽭" <% If Fs_User.Province ="�㽭" then response.Write("selected")%>>�㽭</option>
                <option value="�۰�̨" <% If Fs_User.Province ="�۰�̨" then response.Write("selected")%>>�۰�̨</option>
                <option value="����" <% If Fs_User.Province ="����" then response.Write("selected")%>>����</option>
                <option value="����" <% If Fs_User.Province ="����" then response.Write("selected")%>>����</option>
              </select></td>
            <td class="hback_1"><div align="center"><strong>����</strong></div></td>
            <td class="hback"><input name="City" type="text" id="City" value="<% = Fs_User.City%>" size="26" maxlength="20"> 
            </td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center"><strong>��ַ</strong></div></td>
            <td class="hback"><input name="address" type="text" id="address" value="<% = Fs_User.address%>" size="26" maxlength="100"> 
            </td>
            <td class="hback_1"><div align="center"><strong>��������</strong></div></td>
            <td class="hback"><input name="postcode" type="text" id="postcode" value="<% = Fs_User.postcode%>" size="26" maxlength="20"></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center"><strong>QQ</strong></div></td>
            <td class="hback"><input name="QQ" type="text" id="QQ" value="<% = Fs_User.QQ%>" size="26" maxlength="20"> 
            </td>
            <td class="hback_1"><div align="center"><strong>MSN</strong></div></td>
            <td class="hback"><input name="MSN" type="text" id="MSN" value="<% = Fs_User.MSN%>" size="26" maxlength="100"> 
            </td>
          </tr>
          <tr class="hback"> 
            <td colspan="4" class="hback"><div align="center"> 
                <input name="Action" type="hidden" id="Action" value="Save">
                <input type="submit" name="Submit" value="��������"   onClick="{if(confirm('ȷ�ϱ��������޸ĵĲ�����?')){this.document.form1.submit();return true;}return false;}">
                �� 
                <input type="reset" name="Submit3" value="������д">
              </div></td>
          </tr>
          <tr class="hback">
            <td colspan="4" class="hback"><div align="center"></div></td>
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





