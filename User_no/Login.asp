<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

User_GetParm
Dim ThisPage
ThisPage=NoSqlHack(Request.ServerVariables("SCRIPT_NAME"))
response.Cookies("FoosunUserCookies")("UserLogin_Style_Num")  = p_LoginStyle
If p_LoginStyle="" Or p_LoginStyle = 0 then
	response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") = "1"
End if
Dim forward
forward = Request.QueryString("forward")
If forward="" Then
	forward = left(ThisPage,InStrRev(ThisPage,"/"))&"Main.asp"
End If
forward = Server.URLEncode(forward)
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>--��Ա��½</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="keywords" content="��Ѷ,��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<style>
a{text-decoration: none;} /* �������»���,��Ϊunderline */ 
a:link {color: #232323;} /* δ���ʵ����� */
a:visited {color: #232323;} /* �ѷ��ʵ����� */
a:hover{color: #FF0000;} /* ����������� */ 
a:active {color: #FF0000;} /* ����������� */
td {
	color:#232323;
	font-size:12px;
}
.input{
    FONT-FAMILY: "Verdana, ������";
    FONT-SIZE: 12px;
	COLOR:#F3F3F3;
    text-decoration: none;
    line-height: 150%;
    background:#0099CC;
	border-top-width: 1px;
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-left-width: 1px;
    border-top-style: solid;
	border-right-style: solid;
	border-bottom-style: solid;
	border-left-style: solid;
	border-top-color: #FFFFFF;
	border-right-color: #FFFFFF;
	border-bottom-color: #FFFFFF;
	border-left-color: #FFFFFF;
	padding:0px;
    margin-top: 0px;
    margin-right: 0px;
    margin-bottom: 0px;
    margin-left: 0px;
}
.input_1{
    FONT-FAMILY: "Verdana", "Arial", "Helvetica", "sans-serif";
    FONT-SIZE: 12px;
	COLOR:#006699;
    text-decoration: none;
    line-height: normal;
    background:#FFFFFF;
	border-top-width: 1px;
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-left-width: 1px;
    border-top-style: solid;
	border-right-style: solid;
	border-bottom-style: solid;
	border-left-style: solid;
	border-top-color: #80CCFF;
	border-right-color: #80CCFF;
	border-bottom-color: #80CCFF;
	border-left-color: #80CCFF;
	padding:0px;
    margin-top: 0px;
    margin-right: 0px;
    margin-bottom: 0px;
    margin-left: 0px;
}
</style>
<head>
<script language="JavaScript" type="text/javascript">
function CheckForm()
{
	if(document.myform.name.value=="")
	{
		alert("�����������û�����");
		document.myform.name.focus();
		return false;
	}
	if(document.myform.password.value == "")
	{
		alert("�������������룡");
		document.myform.password.focus();
		return false;
	}
}
</script>
<body background="images/log_bg.gif" topmargin="80" oncontextmenu="//return false;">
<table width="486" border="0" align="center" cellpadding="0" cellspacing="3" bgcolor="#00CCFF">
  <tr> 
    <td bgcolor="#FFFFFF">
	<TABLE WIDTH=486 BORDER=0 CELLPADDING=0 CELLSPACING=0>
        <TR> 
          <TD COLSPAN=2> <IMG SRC="images/i_1.gif" ALT="" WIDTH=486 HEIGHT=87 border="0" usemap="#Map" target="_blank"></TD>
        </TR>
        <TR> 
          <TD width="339" background="images/i_2.gif"> <table width="92%" height="106" border="0" align="center" cellpadding="2" cellspacing="0" class="table">
              <form action="CheckLogin.asp?forward=<%= forward %>"  method="post" name="myform" id="myform"  onsubmit="return CheckForm();">
                <tr class="back"> 
                  <td height="21" class="hback"> <div align="center">����ʽ</div></td>
                  <td class="hback"><select name="Logintype" class="input_1" id="Logintype" style="width:160px">
                      <option value="0" selected>�û���</option>%>
                      <option value="1">�û����</option>
                      <option value="2">�����ʼ�</option>
                    </select></td>
                </tr>
                <tr class="back"> 
                  <td width="23%" height="16" class="hback"> <div align="center">�û��� 
                    </div></td>
                  <td width="77%" class="hback"><input name="name" type="text" class="input_1" id="name4" style="width:160px" value="<%=Request.Cookies("FoosunUserCookie")("FS_UserName")%>"  /> 
                    <input name="AutoGet" type="checkbox" id="AutoGet" value="1" <% If Request.Cookies("FoosunUserCookie")("FS_UserName")<>"" Then Response.Write "checked" End If%>/>
                    ��ס</td>
                </tr>
                <tr class="back"> 
                  <td height="15" class="hback"> <div align="center">�ܡ���</div></td>
                  <td class="hback"><input name="password" type="password" class="input_1" id="password4" style="width:160px;FONT-SIZE:12px;" /></td>
                </tr>
                <tr class="back"> 
                  <td height="12" class="hback">&nbsp;</td>
                  <td class="hback"><a href="GetPassword.asp">�������룿</a> <a href="Register.asp">ע�����û�</a></td>
                </tr>
                <tr class="back"> 
                  <td height="21" class="hback">&nbsp;</td>
                  <td class="hback"><input class="input" type="submit" value="ȷ����½" name="Submit" /> 
                    <input class="input" onClick="javascript:location.href='../'" type="button" value="������ҳ" name="Submit1" />
                  </td>
                </tr>
              </form>
            </table></TD>
          <TD width="147"> <IMG SRC="images/i_3.gif" WIDTH=147 HEIGHT=124 ALT=""></TD>
        </TR>
        <TR> 
          <TD COLSPAN=2> <IMG SRC="images/i_4.jpg" WIDTH=486 HEIGHT=77 ALT=""></TD>
        </TR>
      </TABLE></td>
  </tr>
</table>
<map name="Map">
  <area shape="rect" coords="333,56,394,77" href="http://www.foosun.cn" alt="Foosun Inc.">
</map>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->
</body>
</html>





