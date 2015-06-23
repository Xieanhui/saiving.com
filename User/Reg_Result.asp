<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<%
User_GetParm
if RegisterTF =false then
	Response.Redirect("lib/Error.asp?ErrCodes=<li>暂时关闭注册功能</li><li>或者系统参数丢失!</li>&ErrorUrl="&Request.ServerVariables("URL")&"?"&request.QueryString&"")
	Response.end
End if
if session("FS_UserName") = "" or session("FS_UserNumber") = "" or session("FS_UserPassword") ="" or session("FS_IsCorp") = "" or session("FS_UserEmail") = "" then
	Response.Redirect("lib/Error.asp?ErrCodes=<li>错误的参数</li><li>请不要在外部访问本页!</li>&ErrorUrl=/test")
	Response.end
End if
response.Cookies("FoosunUserCookies")("UserLogin_Style_Num")  = p_LoginStyle
If p_LoginStyle="" Or p_LoginStyle = 0 then
	Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num") = "3"
End if
Dim ReturnUrlpage,strReturnUrlpage
if p_NumReturnUrl = 0 then
	ReturnUrlpage = "会员中心首页"
	strReturnUrlpage = "main.asp"
ElseIf p_NumReturnUrl = 1 then
	ReturnUrlpage = "网站首页"
	strReturnUrlpage = "../"
End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title>会员注册step 4 of 4 step</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="keywords" content="风讯,风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
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
          <td   colspan="2" class="xingmu" height="24">·User Register step 4.(注册成功)</td>
        </tr>
        <tr class="back"> 
          <td width="15%" valign="top" class="hback"><strong>【注册步骤】</strong> <br>
            <br>
            <div align="left"> √同意注册协议<br>
              <br>
              √填写会员资料<br>
              <br>
              √填写联系资料<br>
              <br>
              →注册成功</div>
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
                <td height="18" class="xingmu">恭喜您：<span class="tx"><% = session("FS_NickName") %></span>，您已经在本站注册成功！ </td>
              </tr>
              <tr>
                <td height="29" class="back"><table width="100%" border="0" cellspacing="1" cellpadding="5">
                    <tr class="hback"> 
                      <td width="21%"><div align="right"><strong>用户号：</strong></div></td>
                      <td width="79%"><% = session("FS_UserNumber") %>
                        　　　　　<span class="tx">请牢记用户编号，以后有需要</span></td>
                    </tr>
                    <tr class="hback"> 
                      <td><div align="right"><strong>用户名：</strong></div></td>
                      <td><% = session("FS_UserName") %></td>
                    </tr>
                    <tr class="hback"> 
                      <td><div align="right"><strong>电子邮件：</strong></div></td>
                      <td><% = session("FS_UserEmail") %></td>
                    </tr>
                    <tr class="hback"> 
                      <td><div align="right"><strong>密码：</strong></div></td>
                      <td>
					  <%
					  if p_isValidate =1 then
						    Response.Write("密码已经发送到您的邮箱")
					   Else
					   		Response.Write""&session("TMP_UserPassword")&""		
					   End if
					   %></td>
                    </tr>
					<tr class="hback">
                      <td><div align="right"><strong>状态：</strong></div></td>
                      <td><%
					  if session("FS_IsLock") =1 then 
						  Response.Write("未审核")
					  Else
						  Response.Write("已审核")
					  End if
					  %></td>
                    </tr>
					<tr class="hback"> 
                      <td><div align="right"><strong>会员类型：</strong></div></td>
                      <td> <%
					if  session("FS_IsCorp") =1 then
					  	Response.Write("企业会员")
						if p_isCheckCorp = 1 then 
							Response.Write("(等待审核)")
						Else
							Response.Write("(资料已经审核)")
						End if
					Else
					  	Response.Write("个人会员")
					End if
					   %></td>
                    </tr>
                    
                    <tr class="hback"> 
                      <td><div align="right"><strong>电子邮件状态：</strong></div></td>
                      <td>
					  <%
					  Dim str_isSendMail
					  If str_isSendMail = false then
							Response.Write("未发送或者未发送成功")
					  Else
							Response.Write("邮件已经发送到您的电子邮件:"& session("FS_UserEmail") &",请注意查看")
					  End if
					  %>
					  </td>
                    </tr>
                    <tr class="hback"> 
                      <td><div align="right"><strong>进入会员中心：</strong></div></td>
                      <td> <table width="100%" height="41" border="0" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td width="23" height="41"><span id="countdown" class="titletx"> 
                              <script language="JavaScript">document.write(intLeft);</script>
                              </span></td>
                            <td width="578">秒后自动转到会员中心,直接进入<a href="<% = strReturnUrlpage %>" ><b><% = ReturnUrlpage %></b></a></td>
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
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->
</body>
</html>
<%
'User_Conn.close
'set User_Conn=nothing
%>





