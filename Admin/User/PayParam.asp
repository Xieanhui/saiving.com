<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
dim Conn,User_Conn,rs,str_c_isp,str_c_user,str_c_pass,str_c_url,str_domain,rs_param,str_c_undefined_1,strShowErr
MF_Default_Conn
MF_User_Conn
MF_Session_TF
if not MF_Check_Pop_TF("ME_Pay") then Err_Show 
if not MF_Check_Pop_TF("ME042") then Err_Show 

str_c_isp = CInt(Request("c_isp"))
if str_c_isp<0 then
	str_c_isp=0
end if
set rs_param=Conn.execute("select top 1 MF_Domain from FS_MF_Config")
str_domain=rs_param(0)
rs_param.close:set rs_param=nothing
set rs= Server.CreateObject(G_FS_RS)
rs.open "select c_isp,c_user,c_pass,c_undefined_1 From FS_ME_Pay WHERE c_isp="&str_c_isp,User_Conn,1,3
if rs.eof then
	str_c_user = ""
	str_c_pass = ""
	str_c_undefined_1 = ""
else
	str_c_user = rs("c_user")
	str_c_pass = rs("c_pass")
	str_c_undefined_1 = rs("c_undefined_1")
end if
rs.close:set rs=nothing
if Request.Form("Action")="save" then
	set rs= Server.CreateObject(G_FS_RS)
	rs.open "select * From FS_ME_Pay where c_isp="&str_c_isp,User_Conn,1,3
	if rs.eof then
		rs.addnew
	end if
	rs("c_isp")=NoSqlHack(Request.Form("c_isp"))
	rs("c_user")=NoSqlHack(Request.Form("c_user"))
	rs("c_pass")=NoSqlHack(Request.Form("c_pass"))
	rs("c_undefined_1")=NoSqlHack(Request.Form("c_undefined_1"))
	rs.update
	rs.close:set rs=nothing
	strShowErr = "<li>����ɹ�</li>"
	Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/PayParam.asp")
	Response.end
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
	<title>��������֧��</title>
	<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css" />
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" scroll="yes">
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<form name="form1" method="post" action="">
		<tr>
			<td colspan="2" class="xingmu">
				����֧������
			</td>
		</tr>
		<tr>
			<td colspan="2" class="hback">
				<a href="PayParam.asp">����֧����������</a>��<a href="Order_Pay.asp">֧����������</a>
			</td>
		</tr>
		<tr>
			<td width="19%" align="right" class="hback">
				����֧��ISP
			</td>
			<td class="hback">
				<select name="c_isp" id="c_isp">
					<option value="0" <%if str_c_isp=0 then response.Write("selected")%>>֧����</option>
					<option value="1" <%if str_c_isp=1 then response.Write("selected")%>>����֧��@��</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="hback" align="right">
				�̻���
			</td>
			<td class="hback">
				<input name="c_user" type="text" id="c_user" value="<%=str_c_user%>" /> *</td>
		</tr>
		<tr>
			<td class="hback" align="right">
				�̻�֧����Կ
			</td>
			<td class="hback">
				<input name="c_pass" type="password" id="c_pass" value="<%=str_c_pass%>" /> *</td>
		</tr>
		<tr>
			<td class="hback" align="right">
				֧���ʺ�
			</td>
			<td class="hback">
				<input name="c_undefined_1" type="text" id="c_undefined_1" value="<%=str_c_undefined_1%>" />֧��������Ҫ��д������֧��
			</td>
		</tr>
		<tr>
			<td class="hback">
				&nbsp;
			</td>
			<td class="hback">
				<input type="submit" name="Submit" value="�������" />
				<input name="Action" type="hidden" id="Action" value="save" />
			</td>
		</tr>
		</form>
	</table>
	<script type="text/javascript">
		document.getElementById('c_isp').onchange = function() {
			location.href = 'PayParam.asp?c_isp=' + this.value;
		};
	</script>
</body>
</html>
<%
Conn.close:set conn=nothing
User_Conn.close:set User_Conn=nothing
%>