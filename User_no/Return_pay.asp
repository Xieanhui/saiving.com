<% 'Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
	'�õ�ISP��Ϣ
	dim rs_isp
	set rs_isp= Server.CreateObject(G_FS_RS)
	rs_isp.open "select top 1 c_isp,c_user,c_pass,c_url,c_gurl From FS_ME_Pay",User_Conn,1,3
	if rs_isp.eof then
		strShowErr = "<li>�Ҳ���ϵͳ������Ϣ������û��ͨ����֧�����ܡ�����ϵͳ����Ա��ϵ</li>>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		Dim str_c_isp,str_c_user,str_c_pass,str_c_url,str_c_gurl
		str_c_isp=rs_isp("c_isp")
		str_c_user=rs_isp("c_user")
		str_c_pass=rs_isp("c_pass")
		str_c_url=rs_isp("c_url")
		str_c_gurl=rs_isp("c_gurl")
	end if
%>
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
					<tr>
						<td class="hback">
							<%
		Dim c_mid,c_order,c_orderamount,c_ymd,c_transnum,c_succmark,c_moneytype,c_cause,c_memo1,c_signstr,srcStr,r_signstr
		c_mid			= NoSqlHack(request("c_mid"))			'�̻���ţ��������̻��ɹ��󼴿ɻ�ã������������̻��ɹ����ʼ��л�ȡ�ñ��
		c_order			= NoSqlHack(request("c_order"))		'�̻��ṩ�Ķ�����
		c_orderamount	= NoSqlHack(request("c_orderamount"))	'�̻��ṩ�Ķ����ܽ���ԪΪ��λ��С���������λ���磺13.05
		c_ymd			= NoSqlHack(request("c_ymd"))			'�̻���������Ķ����������ڣ���ʽΪ"yyyymmdd"����20050102
		c_transnum		= NoSqlHack(request("c_transnum"))		'����֧�������ṩ�ĸñʶ����Ľ�����ˮ�ţ����պ��ѯ���˶�ʹ�ã�
		c_succmark		= NoSqlHack(request("c_succmark"))		'���׳ɹ���־��Y-�ɹ� N-ʧ��			
		c_moneytype		= NoSqlHack(request("c_moneytype"))	'֧�����֣�0Ϊ�����
		c_cause			= NoSqlHack(request("c_cause"))		'�������֧��ʧ�ܣ����ֵ����ʧ��ԭ��		
		c_memo1			= NoSqlHack(request("c_memo1"))		'�̻��ṩ����Ҫ��֧�����֪ͨ��ת�����̻�����һ
		c_memo2			= NoSqlHack(request("c_memo2"))		'�̻��ṩ����Ҫ��֧�����֪ͨ��ת�����̻�������
		c_signstr		= NoSqlHack(request("c_signstr"))		'����֧�����ض�������Ϣ����MD5���ܺ���ַ��� 
		if c_mid="" or c_order="" or c_orderamount="" or c_ymd="" or c_moneytype="" or c_transnum="" or c_succmark="" or c_signstr="" then
			strShowErr = "<li>֧��ʧ��</li><li>֧����Ϣ����</li>>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		Dim c_pass	'�̻���֧����Կ����¼�̻������̨(https://www.cncard.net/admin/)���ڹ�����ҳ���ҵ���ֵ
		c_pass = str_c_pass
		srcStr = c_mid & c_order & c_orderamount & c_ymd & c_transnum & c_succmark & c_moneytype & c_memo1 & c_memo2 & c_pass
		r_signstr	= MD5(srcStr)
		if r_signstr<>c_signstr then  
			strShowErr = "<li>֧��ʧ��</li><li>ǩ����֤ʧ��</li>>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		Dim MerchantID	'�̻��Լ��ı�� 
		if str_c_user<>c_mid then 
			strShowErr = "<li>֧��ʧ��</li><li>�ύ���̻��������</li>>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if 
		dim rs_1,sql_1
		sql_1="select top 1 OrderNumber,MoneyAmount,AddTime from FS_ME_Order where OrderNumber='"& c_order &"' and UserNumber='"&Fs_User.UserNumber&"'"
		set rs_1=server.CreateObject(G_FS_RS)
		rs_1.open sql_1,User_conn
		if rs_1.eof then
			rs_1.close:set rs_1=nothing
			set conn=nothing
			set user_Conn=nothing
			strShowErr = "<li>֧��ʧ��</li><li>δ�ҵ��ö�����Ϣ</li>>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		Dim r_orderamount
		r_orderamount=rs_1("MoneyAmount")
		if  ccur(r_orderamount)<>ccur(c_orderamount) then
			strShowErr = "<li>֧��ʧ��</li><li>֧���������</li>>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		Dim r_ymd
		r_ymd=rs_1("AddTime")
		r_ymd=year(r_ymd)
		if month(rs_1("AddTime"))<10 then:r_ymd=r_ymd&"0"&month(rs_1("AddTime")):else:r_ymd=r_ymd&month(rs_1("AddTime")):end if
		if day(rs_1("AddTime"))<10 then:r_ymd=r_ymd&"0"&day(rs_1("AddTime")):else:r_ymd=r_ymd&day(rs_1("AddTime")):end if
		'��ʱ����
		'if  r_ymd<>c_ymd then 
		'	strShowErr = "<li>֧��ʧ��</li><li>����ʱ������</li>>"
		'	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		'	Response.end
		'end if
		'Dim r_memo1
		'r_memo1 = rs("ת������һ")
		'Dim r_memo2
		'r_memo2 = rs("ת��������")
		'IF r_memo1<>c_memo1 or r_memo2<>c_memo2 THEN
		'	response.write "�����ύ����"
		'	response.end
		'END IF
		if c_succmark<>"Y" and c_succmark<>"N" then
			strShowErr = "<li>֧��ʧ��</li><li>�����ύ����</li>>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		IF c_succmark="Y" THEN
			dim RsPay
			Set RsPay = User_Conn.execute("select isPay,isOnPay From FS_ME_Order where OrderNumber='"& NoSqlHack(c_order) &"' and UserNumber='"&Fs_User.UserNumber&"'")
			if CSTR(RsPay("IsPay"))="1" then
				RsPay.close:set RsPay = nothing
				strShowErr = "<li>֧��ʧ��</li><li>�����ظ�֧�����Ķ���</li>>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			else
				User_Conn.execute("Update FS_ME_Order set isLock=0,IsSuccess=1 where OrderNumber='"& c_order &"' and UserNumber='"&Fs_User.UserNumber&"'")
				dim rs_u
				set rs_u= Server.CreateObject(G_FS_RS)
				rs_u.open "select Content,M_PayDate,OrderNumber,M_state,isPay From FS_ME_Order where OrderNumber='"& c_order &"' and UserNumber='"&Fs_User.UserNumber&"'",User_Conn,1,3
				rs_u("M_PayDate")=now
				rs_u("Content")=c_memo1
				rs_u("M_state")=1
				rs_u("isPay")=1
				rs_u.update
				rs_u.close:set rs_u=nothing
				if RsPay("isOnPay") =1 then
					Call Fs_User.AddLog("����֧��",Fs_User.UserNumber,0,c_orderamount,"����֧����ý��:"&c_orderamount&"",0)
					Response.Write("<script>alert(""���׳ɹ���\n֧���ɹ�:������:"& c_order &"��\n���׽��:"&c_orderamount&"RMB"");location.href=""Order_Pay.asp"";</script>")
					Response.End
				elseif RsPay("isOnPay") =0 then
					User_Conn.execute("Update FS_ME_Users set FS_Money=FS_Money+"&c_orderamount&" where UserNumber='"&Fs_User.UserNumber&"'")
					Call Fs_User.AddLog("����֧��",Fs_User.UserNumber,0,c_orderamount,"����֧������:"&c_orderamount&"",0)
					Response.Write("<script>alert(""���׳ɹ���\n֧���ɹ�:������:"& c_order &"��\n���׽��:"&c_orderamount&"RMB�ȶ�Ľ���Ѿ����������ʻ���\n������Ƕ���֧���������ʻ�֧�����Ķ���"");location.href=""Order_Pay.asp"";</script>")
					Response.End
				else
					Response.Write("<script>alert(""δ֪����"");location.href=""Order_Pay.asp"";</script>")
					Response.End
				end if
				RsPay.close:set RsPay = nothing
			end if
		elseif c_succmark="N" then
			strShowErr = "<li>����ʧ��</li>δ֪ԭ��</li><li>�������Ķ��������в�ѯ������</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Order_Pay.asp")
			Response.end
		end if
							%>
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
</body>
</html>
<%
Set Fs_User = Nothing
%>