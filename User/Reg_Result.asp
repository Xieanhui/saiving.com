<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<%
User_GetParm
if RegisterTF =false then
	Response.Redirect("lib/Error.asp?ErrCodes=<li>��ʱ�ر�ע�Ṧ��</li><li>����ϵͳ������ʧ!</li>&ErrorUrl="&Request.ServerVariables("URL")&"?"&request.QueryString&"")
	Response.end
End if
if session("FS_UserName") = "" or session("FS_UserNumber") = "" or session("FS_UserPassword") ="" or session("FS_IsCorp") = "" or session("FS_UserEmail") = "" then
	Response.Redirect("lib/Error.asp?ErrCodes=<li>����Ĳ���</li><li>�벻Ҫ���ⲿ���ʱ�ҳ!</li>&ErrorUrl=/test")
	Response.end
End if
response.Cookies("FoosunUserCookies")("UserLogin_Style_Num")  = p_LoginStyle
If p_LoginStyle="" Or p_LoginStyle = 0 then
	Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num") = "3"
End if
Dim ReturnUrlpage,strReturnUrlpage
if p_NumReturnUrl = 0 then
	ReturnUrlpage = "��Ա������ҳ"
	strReturnUrlpage = "main.asp"
ElseIf p_NumReturnUrl = 1 then
	ReturnUrlpage = "��վ��ҳ"
	strReturnUrlpage = "../"
End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title>��Աע��step 4 of 4 step</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="keywords" content="��Ѷ,��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head>
<script language="JavaScript">
var intLeft = 16; 
function leavePage() {
if (0 == intLeft)
	window.location.href='<%=strReturnUrlpage%>'
else {
intLeft -= 1;
document.all.countdown.innerText = intLeft + " ";
setTimeout("leavePage()", 1000);
}
}
</script>
<body oncontextmenu="return false;" onLoad="setTimeout('leavePage()', 1000)">
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="0">
  <tr> 
    <td><table width="100%" height="279" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
        <tr class="back"> 
          <td   colspan="2" class="xingmu" height="24">��User Register step 4.(ע��ɹ�)</td>
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
            <div align="center" class="tx"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td>&nbsp;</td>
                </tr>
              </table>
            </div>
			  
            <table width="96%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
              <tr> 
                <td height="18" class="xingmu">��ϲ����<span class="tx"><% = session("FS_NickName") %></span>�����Ѿ��ڱ�վע��ɹ��� </td>
              </tr>
              <tr>
                <td height="29" class="back"><table width="100%" border="0" cellspacing="1" cellpadding="5">
                    <tr class="hback"> 
                      <td width="21%"><div align="right"><strong>�û��ţ�</strong></div></td>
                      <td width="79%"><% = session("FS_UserNumber") %>
                        ����������<span class="tx">���μ��û���ţ��Ժ�����Ҫ</span></td>
                    </tr>
                    <tr class="hback"> 
                      <td><div align="right"><strong>�û�����</strong></div></td>
                      <td><% = session("FS_UserName") %></td>
                    </tr>
                    <tr class="hback"> 
                      <td><div align="right"><strong>�����ʼ���</strong></div></td>
                      <td><% = session("FS_UserEmail") %></td>
                    </tr>
                    <tr class="hback"> 
                      <td><div align="right"><strong>���룺</strong></div></td>
                      <td>
					  <%
					  if p_isValidate =1 then
						    Response.Write("�����Ѿ����͵���������")
					   Else
					   		Response.Write""&session("TMP_UserPassword")&""		
					   End if
					   %></td>
                    </tr>
					<tr class="hback">
                      <td><div align="right"><strong>״̬��</strong></div></td>
                      <td><%
					  if session("FS_IsLock") =1 then 
						  Response.Write("δ���")
					  Else
						  Response.Write("�����")
					  End if
					  %></td>
                    </tr>
					<tr class="hback"> 
                      <td><div align="right"><strong>��Ա���ͣ�</strong></div></td>
                      <td> <%
					if  session("FS_IsCorp") =1 then
					  	Response.Write("��ҵ��Ա")
						if p_isCheckCorp = 1 then 
							Response.Write("(�ȴ����)")
						Else
							Response.Write("(�����Ѿ����)")
						End if
					Else
					  	Response.Write("���˻�Ա")
					End if
					   %></td>
                    </tr>
                    
                    <tr class="hback"> 
                      <td><div align="right"><strong>�����ʼ�״̬��</strong></div></td>
                      <td>
					  <%
					  Dim str_isSendMail
					  If str_isSendMail = false then
							Response.Write("δ���ͻ���δ���ͳɹ�")
					  Else
							Response.Write("�ʼ��Ѿ����͵����ĵ����ʼ�:"& session("FS_UserEmail") &",��ע��鿴")
					  End if
					  %>
					  </td>
                    </tr>
                    <tr class="hback"> 
                      <td><div align="right"><strong>�����Ա���ģ�</strong></div></td>
                      <td> <table width="100%" height="41" border="0" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td width="23" height="41"><span id="countdown" class="titletx"> 
                              <script language="JavaScript">document.write(intLeft);</script>
                              </span></td>
                            <td width="578">����Զ�ת����Ա����,ֱ�ӽ���<a href="<% = strReturnUrlpage %>" ><b><% = ReturnUrlpage %></b></a></td>
                          </tr>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table>
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
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->
</body>
</html>
<%
'User_Conn.close
'set User_Conn=nothing
%>





