<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<%
User_GetParm
if RegisterTF =false then
	strShowErr = "<li>��ʱ�ر�ע�Ṧ��</li><li>����ϵͳ������ʧ!</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl="&Request.ServerVariables("URL")&"?"&request.QueryString&"")
	Response.end
End if
if Not isnull(DefaultGroupID) then
  if DefaultGroupID = 0 then
	strShowErr = "<li>����Ա��û����Ĭ�ϻ�Ա�顣������ʱ����ע�ᣬ�������Ա��ϵ!</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl="&Request.ServerVariables("URL")&"?"&request.QueryString&"")
	Response.end
  end if
  dim rsGroup
  set rsGroup = User_Conn.execute("select GroupID,GroupName from FS_ME_Group where GroupType=1 and GroupID="&clng(DefaultGroupID))
  if rsGroup.eof then
	strShowErr = "<li>�����쳣!</li><li>����ϵͳ�ṩ�̻�ü���֧��!!</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl="&Request.ServerVariables("URL")&"?"&request.QueryString&"")
	Response.end
  end if
  rsGroup.close:set rsGroup=nothing
else
	strShowErr = "<li>����Ա��û����Ĭ�ϻ�Ա�顣������ʱ����ע�ᣬ�������Ա��ϵ!</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl="&Request.ServerVariables("URL")&"?"&request.QueryString&"")
	Response.End
end if
response.Cookies("FoosunUserCookies")("UserLogin_Style_Num")  = p_LoginStyle
If p_LoginStyle="" Or p_LoginStyle = 0 then
	response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") = "2"
End If
Dim forward
forward = Request.QueryString("forward")
forward = Server.URLEncode(forward)
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title>��Աע��step 1 of  4 step</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="keywords" content="��Ѷ,��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head>
<body oncontextmenu="return false;">
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="0">
  <tr> 
    <td><table width="100%" height="279" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
        <tr class="back"> 
          <td   colspan="2" class="xingmu" height="24">��User Register step 1.(ע��Э��)</td>
        </tr>
        <tr class="back"> 
          <td width="15%" valign="top" class="hback"><strong>��ע�Ჽ�衿</strong> <br>
            <br>
            <div align="left"> ��ͬ��ע��Э��<br>
              <br>
              ����д��Ա����<br>
              <br>
              ����д��ϵ����<br>
              <br>
              ��ע��ɹ�</div>
            </td>
          <td width="85%" valign="top" class="hback"> 
		  <%If RegisterTF = false then%>
		  	  <div align="center" class="tx"><p></p>
              <p>&nbsp;</p>
              <p>����Ա�Ѿ��ر�ע��!</p>
            </div>
			  <%Else%>
              <table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
				<tr> 
                  <td height="228"><% =  RegisterNotice %></td>
                </tr>
                <tr> 
                  <td height="39">
					<div align="center">
					<input style="CURSOR: hand" onClick="window.location.href='Reg_info.asp?forward=<%= forward %>&SubSys=<%=Request.QueryString("SubSys")%>'" type="submit" name="Submit3" value="ͬ��Э��" id="agree">
										��
					<input class="button" onClick="location.href='../'" type="button" value="��ͬ��" name="Submit12" />
										�� 
                    </div></td>
                </tr>
              </table>
			  <%End if%>
            </td>
        </tr>
        <tr class="back"> 
          <td height="26"  colspan="2" class="xingmu"> <div align="left"> 
              <!--#include file="Copyright.asp" -->
            </div></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->
<SCRIPT LANGUAGE="JavaScript">
<!--
document.getElementById('agree').disabled=true;
for (i=6; i>0; i--)
{
	window.setTimeout('change('+i+')',i*1000);
}
window.setTimeout("agree()",6000);
function agree()
{
	document.getElementById('agree').value='ͬ��Э��';
	document.getElementById('agree').disabled=false;
}
function change(Num)
{	
	Num = 6-Num;
	document.getElementById('agree').value='ͬ��Э��'+Num;
}
//-->
</SCRIPT>





