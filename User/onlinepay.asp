<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<!--#include file="../FS_Inc/Md5.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
	<title>��ӭ�û�<%=Fs_User.UserName%>����<%=GetUserSystemTitle%>-��������֧��</title>
	<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css" />
</head>
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
			<td colspan="2" class="xingmu" height="26">
				<!--#include file="Top_navi.asp" -->
			</td>
		</tr>
		<tr class="back">
			<td width="18%" valign="top" class="hback">
				<div align="left">
					<!--#include file="menu.asp" -->
				</div>
			</td>
			<td width="82%" valign="top" class="hback">
				<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
					<tr class="hback">
						<td class="hback">
							<strong>λ�ã�</strong><a href="../">��վ��ҳ</a> &gt;&gt; <a href="main.asp">��Ա��ҳ</a> &gt;&gt; �������г�ֵ
						</td>
					</tr>
				</table>
				<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
					<form name="form1" method="post" action="">
					<tr class="hback">
						<td height="32" colspan="2" class="xingmu">
							Ϊ�ʻ���ֵ
						</td>
					</tr>
					<tr class="hback">
						<td width="15%" height="32" align="right" class="hback">
							���
						</td>
						<td width="60%" class="hback">
							<input name="PayMoney" type="text" id="PayMoney" value="<% = NoSqlHack(Request.QueryString("Moneys"))%>" readonly>
							�����
							<input name="OrderNumber" type="hidden" id="OrderNumber" value="<% = NoSqlHack(Request.QueryString("OrderNumber"))%>">
							<input name="OrderID" type="hidden" id="OrderID" value="<% = NoSqlHack(Request.QueryString("OrderID"))%>">
						</td>
					</tr>
					<tr class="hback">
						<td height="32" class="hback">
							<div align="right">
								֧����ʽ</div>
						</td>
						<td class="hback">
							<select name="c_isp" id="c_isp">
								<option value="0">֧����</option>
								<option value="1">����֧��@��</option>
							</select>
						</td>
					</tr>
					<tr class="hback">
						<td height="32" class="hback">
							&nbsp;
						</td>
						<td class="hback">
							<input type="button" id="btnPay" name="Submit" value="ȷ��֧����?" />
							<input type="reset" name="Submit2" value=" ���� " />
						</td>
					</tr>
					</form>
				</table>
				<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
					<tr>
						<td width="100%" class="xingmu">
							ѡ������֧����ʽ
						</td>
					</tr>
					<tr>
						<td class="hback">
							<a href="onlinepay.asp?OrderNumber=<%=NoSqlHack(Request("OrderNumber"))%>&Moneys=<%=NoSqlHack(Request("Moneys"))%>&OrderID=<%=NoSqlHack(Request("OrderID"))%>">����֧��</a> <a href="PayCenter.asp?PayStyle=MySelfAcc&OrderNumber=<%=NoSqlHack(Request("OrderNumber"))%>&Moneys=<%=NoSqlHack(Request("Moneys"))%>&OrderID=<%=NoSqlHack(Request("OrderID"))%>">�ʻ�֧��</a>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr class="back">
			<td height="20" colspan="2" class="xingmu">
				<div align="left">
					<!--#include file="Copyright.asp" -->
				</div>
			</td>
		</tr>
	</table>
	<script type="text/javascript">
		document.getElementById('btnPay').onclick = function() {
			if (confirm('��ȷ��֧����')) {
				payType = document.getElementById('c_isp').value;
				switch (payType) {
					case '0':
						location.href = 'pay/alipay/index.asp' + location.search + "&c_isp=" + payType;
						break;
					case '1':
						location.href = 'pay/cncard/index.asp' + location.search + "&c_isp=" + payType;
						break;
				}

			}
			return false;
		};
	</script>
</body>
</html>
<%
Set Fs_User = Nothing
%>