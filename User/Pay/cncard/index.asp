<% Option Explicit %>
<!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../../FS_Inc/Function.asp" -->
<!--#include file="../../lib/strlib.asp" -->
<!--#include file="../../lib/UserCheck.asp" -->
<!--#include file="../../../FS_Inc/Md5.asp" -->
<%
	Dim RsPay
	on error resume next
	if Request.QueryString("Moneys")="" Then
		strShowErr = "<li>����Ĳ���</li>"
		Response.Redirect("../../lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Else
		set RsPay = User_Conn.execute("select OrderId,IsPay,MoneyAmount From FS_ME_Order where OrderNumber='"&NoSqlHack(Request.QueryString("OrderNumber"))&"' and OrderID="&CintStr(Request.QueryString("OrderID"))&"")
		if RsPay.EOF then
			strShowErr = "<li>����Ĳ������Ҳ���������¼</li>"
			'Response.Redirect("../../lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		Else
			if CCur(RsPay("MoneyAmount")) <> CCur(Request.QueryString("Moneys")) then
				strShowErr = "<li>֧������ʵ�ʽ���</li>"
				Response.Redirect("../../lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			end if
			if RsPay("IsPay")=1 then
				strShowErr = "<li>�˶����Ѿ�֧����������֧��</li>"
				Response.Redirect("../../lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			end if
			if RsPay("MoneyAmount")=0 then
				strShowErr = "<li>�˶������Ϊ0������Ҫ֧��</li>"
				Response.Redirect("../../lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			end if
		end if
	end if
	'�õ�ISP��Ϣ
	str_c_url = "http://"&Request.ServerVariables("HTTP_HOST")&Request.ServerVariables("URL")
	str_c_url = Left(str_c_url,InstrRev(str_c_url,"/"))&"GetPayHandle.asp"
	dim rs_isp,g_rs,c_isp
	c_isp = cint(Request("c_isp"))
	set rs_isp= Server.CreateObject(G_FS_RS)
	rs_isp.open "select c_isp,c_user,c_pass,c_url,c_gurl From FS_ME_Pay WHERE c_isp = "&c_isp,User_Conn,1,3
	if rs_isp.eof then
		strShowErr = "<li>�Ҳ���ϵͳ������Ϣ������û��ͨ����֧�����ܡ�����ϵͳ����Ա��ϵ</li>>"
		Response.Redirect("../../lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		Dim str_c_isp,str_c_user,str_c_pass,str_c_url
		str_c_isp=rs_isp("c_isp")
		str_c_user=rs_isp("c_user")
		str_c_pass=rs_isp("c_pass")
	end if
	'���Ϊ�¹����û������ţ���Ѷ������ݲ������ݿ���
	'��ö�������
	Dim tmp_OrderNumber,tmp_ramcode ,RsOrderObj,tmp_PayMoney,tmp_ymd ,tmp_ymd_ 
	tmp_ramcode = NoSqlHack(Request.QueryString("OrderNumber"))
	tmp_PayMoney=NoSqlHack(Request.QueryString("Moneys"))
	tmp_ymd=now
	tmp_ymd_=year(tmp_ymd)
	if month(now)<10 then:tmp_ymd_=tmp_ymd_&"0"&month(tmp_ymd):else:tmp_ymd_=tmp_ymd_&month(tmp_ymd):end if
	if day(now)<10 then:tmp_ymd_=tmp_ymd_&"0"&day(tmp_ymd):else:tmp_ymd_=tmp_ymd_&day(tmp_ymd):end if
'---������Ϣ---
	Dim c_mid			'�̻���ţ��������̻��ɹ��󼴿ɻ�ã������������̻��ɹ����ʼ��л�ȡ�ñ��
	Dim c_order			'�̻���վ���ɵĶ����ţ������ظ�
	Dim c_name			'�̻������е��ջ�������
	Dim c_address		'�̻������е��ջ��˵�ַ
	Dim c_tel			'�̻������е��ջ��˵绰
	Dim c_post			'�̻������е��ջ����ʱ�
	Dim c_email			'�̻������е��ջ���Email
	Dim c_orderamount	'�̻������ܽ��
	Dim c_ymd			'�̻������Ĳ������ڣ���ʽΪ"yyyymmdd"����20050102
	Dim c_moneytype		'֧�����֣�0Ϊ�����
	Dim c_retflag		'�̻�����֧���ɹ����Ƿ���Ҫ�����̻�ָ�����ļ���0�����÷��� 1����Ҫ����
	Dim c_paygate		'������̻���վѡ�����������ø�ֵ������ֵ�ɲμ�������֧��@�������ӿ��ֲᡷ��¼һ�����������֧��@��ѡ�����д���Ϊ��ֵ��
	Dim c_returl		'���c_retflagΪ1ʱ����ֵ����֧���ɹ��󷵻ص��ļ���·��
	Dim c_memo1			'�̻���Ҫ��֧�����֪ͨ��ת�����̻�����һ
	Dim c_memo2			'�̻���Ҫ��֧�����֪ͨ��ת�����̻�������
	Dim c_signstr		'�̻��Զ�����Ϣ����MD5ǩ������ַ���
	Dim c_pass			'֧����Կ�����¼�̻������̨�����ʻ���Ϣ->������Ϣ->��ȫ��Ϣ�е�֧����Կ��
	Dim notifytype		'0��֪ͨͨ��ʽ/1������֪ͨ��ʽ����ֵΪ��֪ͨͨ��ʽ
	Dim c_language		'�������˹��ʿ�֧��ʱ����ʹ�ø�ֵ����������������֧��ʱ��ҳ�����֣�ֵΪ��0����ҳ����ʾΪ����/1����ҳ����ʾΪӢ��
	Dim srcStr
	c_mid		= str_c_user
	c_order		=  tmp_ramcode
	c_name		= ""
	c_address	= ""
	c_tel		= ""
	c_post		= ""
	c_email		= ""
	c_orderamount	= tmp_PayMoney
	c_ymd		= tmp_ymd_
	c_moneytype	= "0"
	c_retflag	= "1"
	c_paygate	= ""
	c_returl	= str_c_url	'�õ�ַΪ�̻���������֧�����֪ͨ��ҳ�棬���ύ�����ļ���(��Ӧ�����ļ���GetPayNotify.asp)
	if c_isp>0 then
	'��û�Ա��
		c_memo1		= c_isp
	else
		c_memo1		= ""
	end if
	c_memo2		= ""
	c_pass		= str_c_pass
	notifytype	= "0"
	c_language	= "0"
	srcStr = c_mid & c_order & c_orderamount & c_ymd & c_moneytype & c_retflag & c_returl & c_paygate & c_memo1 & c_memo2 & notifytype & c_language & c_pass
	'˵�����������ָ��֧����ʽ(c_paygate)��ֵʱ����Ҫ�����û�ѡ��֧����ʽ��Ȼ���ٸ����û�ѡ��Ľ�����������MD5���ܣ�Ҳ����˵����ʱ����ҳ��Ӧ�ò��Ϊ����ҳ�棬��Ϊ����������ɡ�
'---�Զ�����Ϣ����MD5���� 
	c_signstr	= MD5(srcStr,32)
'end if 
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
	<title>��ӭ�û�<%=Fs_User.UserName%>����<%=GetUserSystemTitle%>-��������֧��</title>
	<link href="../../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css" />
</head>
<body>
	<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
		<tr>
			<td>
				<!--#include file="../../top.asp" -->
			</td>
		</tr>
	</table>
	<table width="98%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
		<tr class="back">
			<td colspan="2" class="xingmu" height="26">
				<!--#include file="../../Top_navi.asp" -->
			</td>
		</tr>
		<tr class="back">
			<td width="18%" valign="top" class="hback">
				<div align="left">
					<!--#include file="../../menu.asp" -->
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
					<form name="form1" method="post" action="https://www.cncard.net/purchase/getorder.asp">
					<tr class="hback">
						<td class="hback">
							<input type="hidden" name="c_mid" value="<%=c_mid%>" />
							<input type="hidden" name="c_order" value="<%=c_order%>" />
							<input type="hidden" name="c_name" value="<%=c_name%>" />
							<input type="hidden" name="c_address" value="<%=c_address%>" />
							<input type="hidden" name="c_tel" value="<%=c_tel%>" />
							<input type="hidden" name="c_post" value="<%=c_post%>" />
							<input type="hidden" name="c_email" value="<%=c_email%>" />
							<input type="hidden" name="c_orderamount" value="<%=c_orderamount%>" />
							<input type="hidden" name="c_ymd" value="<%=c_ymd%>" />
							<input type="hidden" name="c_moneytype" value="<%=c_moneytype%>" />
							<input type="hidden" name="c_retflag" value="<%=c_retflag%>" />
							<input type="hidden" name="c_paygate" value="<%=c_paygate%>" />
							<input type="hidden" name="c_returl" value="<%=c_returl%>" />
							<input type="hidden" name="c_memo1" value="<%=c_memo1%>" />
							<input type="hidden" name="c_memo2" value="<%=c_memo2%>" />
							<input type="hidden" name="c_language" value="<%=c_language%>" />
							<input type="hidden" name="notifytype" value="<%=notifytype%>" />
							<input type="hidden" name="c_signstr" value="<%=c_signstr%>" />
							<input type="submit" name="submit" value="��� -> ����֧��@��" onclick="{if(confirm('��ȷ��֧����')){this.document.form1.submit();return true;}return false;}" />
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
							<a href="../../onlinepay.asp?OrderNumber=<%=NoSqlHack(Request("OrderNumber"))%>&Moneys=<%=NoSqlHack(Request("Moneys"))%>&OrderID=<%=NoSqlHack(Request("OrderID"))%>">����֧��</a> <a href="../../PayCenter.asp?PayStyle=MySelfAcc&OrderNumber=<%=NoSqlHack(Request("OrderNumber"))%>&Moneys=<%=NoSqlHack(Request("Moneys"))%>&OrderID=<%=NoSqlHack(Request("OrderID"))%>">�ʻ�֧��</a>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr class="back">
			<td height="20" colspan="2" class="xingmu">
				<div align="left">
					<!--#include file="../../Copyright.asp" -->
				</div>
			</td>
		</tr>
	</table>
</body>
</html>
<%
Set Fs_User = Nothing
%>