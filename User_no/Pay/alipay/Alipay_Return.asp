<% Option Explicit %>
<!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../../FS_Inc/Function.asp" -->
<!--#include file="../../lib/strlib.asp" -->
<!--#include file="../../lib/UserCheck.asp" -->
<!--#include file="Alipay_Lib.asp"-->
<%
Dim AlipayObj,str_c_url,state
'�õ�ISP��Ϣ
dim rs_isp
set rs_isp= Server.CreateObject(G_FS_RS)
rs_isp.open "select c_isp,c_user,c_pass,c_undefined_1 From FS_ME_Pay WHERE c_isp = 0",User_Conn,1,3
if rs_isp.eof then
	strShowErr = "<li>�Ҳ���ϵͳ������Ϣ������û��ͨ����֧�����ܡ�����ϵͳ����Ա��ϵ</li>>"
	Response.Redirect("../../lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
else
	Dim str_c_user,str_c_pass,str_c_email
	str_c_user=rs_isp("c_user")
	str_c_pass=rs_isp("c_pass")
	str_c_email=rs_isp("c_undefined_1")
end if

Set AlipayObj = New Alipay
AlipayObj.key=str_c_pass
AlipayObj.partner=str_c_user

state = AlipayObj.NotifyPage()
if state="fail" then
	strShowErr = "<li>��Ϣ��Դ����ȷ</li>"
	Response.Redirect("../../lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
else
	Dim out_trade_no,total_fee
	out_trade_no = AlipayObj.DelStr(Request("out_trade_no"))  '��ȡ������
	total_fee    = AlipayObj.DelStr(Request("total_fee"))     '��ȡ֧�����ܼ۸�
end if
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
					<tr>
						<td class="hback">
							<%
				dim rs_u
				set rs_u= Server.CreateObject(G_FS_RS)
				rs_u.open "select * From FS_ME_Order where OrderNumber='"& out_trade_no &"' and UserNumber='"&Fs_User.UserNumber&"'",User_Conn,1,3
				if rs_u.eof then
					strShowErr = "<li>����������</li>"
					Response.Redirect("../../lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
					Response.end
				else
					if rs_u("isPay")="0" then
						rs_u("M_PayDate")=now
						rs_u("M_state")=1
						rs_u("isLock")=0
						rs_u("IsSuccess")=1
						rs_u("isPay")=1
						rs_u.update
					end if
					if rs_u("isOnPay") =1 then
						Call Fs_User.AddLog("����֧��",Fs_User.UserNumber,0,total_fee,"����֧����ý��:"&total_fee&"",0)
						Response.Write("<script>alert(""���׳ɹ���\n֧���ɹ�:������:"& out_trade_no &"��\n���׽��:"&total_fee&"RMB"");location.href=""Order_Pay.asp"";</script>")
						Response.End
					elseif rs_u("isOnPay") =0 then
						User_Conn.execute("Update FS_ME_Users set FS_Money=FS_Money+"&total_fee&" where UserNumber='"&Fs_User.UserNumber&"'")
						Call Fs_User.AddLog("����֧��",Fs_User.UserNumber,0,total_fee,"����֧������:"&total_fee&"",0)
						Response.Write("<script>alert(""���׳ɹ���\n֧���ɹ�:������:"& out_trade_no &"��\n���׽��:"&total_fee&"RMB�ȶ�Ľ���Ѿ����������ʻ���\n������Ƕ���֧���������ʻ�֧�����Ķ���"");location.href=""../../Order_Pay.asp"";</script>")
						Response.End
					else
						Response.Write("<script>alert(""δ֪����"");location.href=""Order_Pay.asp"";</script>")
						Response.End
					end if
				end if
				rs_u.close:set rs_u=nothing
							%>
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






<%
key = ""      '֧������ȫ������
partner      = ""  '֧��������id

out_trade_no = DelStr(Request("out_trade_no"))  '��ȡ������
total_fee    = DelStr(Request("total_fee"))     '��ȡ֧�����ܼ۸�
'�����ȡ��������������д ���� =DelStr(Request.Form("��ȡ������"))

'**********************�ж���Ϣ�ǲ���֧��������********************
alipayNotifyURL = "http://notify.alipay.com/trade/notify_query.do?"
alipayNotifyURL = alipayNotifyURL & "partner=" & partner & "&notify_id=" & Request("notify_id")
Set Retrieval   = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
Retrieval.setOption 2, 13056
Retrieval.open "GET", alipayNotifyURL, False, "", ""
Retrieval.send()
ResponseTxt   = Retrieval.ResponseText
Set Retrieval = Nothing
'*******************************************************************

'*******��ȡ֧����GET����֪ͨ��Ϣ,�ж���Ϣ�ǲ��Ǳ��޸Ĺ�************

For Each varItem in Request.QueryString
	mystr = varItem & "=" & Request(varItem) & "^" & mystr
Next

If mystr <> "" Then
	mystr = Left(mystr,Len(mystr) - 1)
End If

mystr  = Split(mystr, "^")
Count  = UBound(mystr)
'�Բ�������

For i = Count To 0 Step - 1
	minmax       = mystr( 0 )
	minmaxSlot   = 0

	For j = 1 To i
		mark        = (mystr( j ) > minmax)

		If mark Then
			minmax     = mystr( j )
			minmaxSlot = j
		End If

	Next

	If minmaxSlot <> i Then
		temp = mystr( minmaxSlot )
		mystr( minmaxSlot ) = mystr( i )
		mystr( i ) = temp
	End If

Next

'����md5ժҪ�ַ���

For j = 0 To Count Step 1
	value = Split(mystr( j ), "=")

	If  value(1) <> "" And value(0) <> "sign" And value(0) <> "sign_type"  Then

		If j = Count Then
			md5str = md5str & mystr( j )
		Else
			md5str = md5str & mystr( j ) & "&"
		End If

	End If

Next

md5str = md5str & key
mysign = md5(md5str)
'********************************************************

If mysign = Request("sign") And ResponseTxt = "true"   Then
	Response.Write "����ɹ�ҳ��"        '�������ָ������Ҫ��ʾ������
	' �����������֧�����Ĺ�����ܣ����ڷ��ص���Ϣ���治Ҫ�������жϣ���������У��ͨ���������ֵ������������Ҫ��ȡ�����ʹ�ù����Ľ��,
	' ���ȡ������Ϣ������ֶ�discount��ֵ��ȡ����ֵ��������Ҹ����ŻݵĽ��� ԭ�������ܽ��=��Ҹ���صĽ��total_fee +|discount|.
Else
	Response.Write "��תʧ��"          '�������ָ������Ҫ��ʾ������
End If

Function DelStr(Str)

	If IsNull(Str) Or IsEmpty(Str) Then
		Str = ""
	End If

	DelStr = Replace(Str,";","")
	DelStr = Replace(DelStr,"'","")
	DelStr = Replace(DelStr,"&","")
	DelStr = Replace(DelStr," ","")
	DelStr = Replace(DelStr,"��","")
	DelStr = Replace(DelStr,"%20","")
	DelStr = Replace(DelStr,"--","")
	DelStr = Replace(DelStr,"==","")
	DelStr = Replace(DelStr,"<","")
	DelStr = Replace(DelStr,">","")
	DelStr = Replace(DelStr,"%","")
End Function

%>