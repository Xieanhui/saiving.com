<% Option Explicit %>
<!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../../FS_Inc/Function.asp" -->
<!--#include file="../../lib/strlib.asp" -->
<!--#include file="../../lib/UserCheck.asp" -->
<!--#include file="Alipay_Lib.asp"-->
<%
Dim RsPay,str_GoodsName
'on error resume next
if Request.QueryString("Moneys")="" Then
	strShowErr = "<li>����Ĳ���</li>"
	Response.Redirect("../../lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
Else
	set RsPay = User_Conn.execute("select OrderId,IsPay,MoneyAmount,Content,AddTime From FS_ME_Order where OrderNumber='"&NoSqlHack(Request.QueryString("OrderNumber"))&"' and OrderID="&CintStr(Request.QueryString("OrderID"))&"")
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
		str_GoodsName =  RsPay("Content")
	end if
end if
'�õ�ISP��Ϣ
dim rs_isp,g_rs,c_isp,str_c_email
c_isp = cint(Request("c_isp"))
set rs_isp= Server.CreateObject(G_FS_RS)
rs_isp.open "select c_isp,c_user,c_pass,c_undefined_1 From FS_ME_Pay WHERE c_isp = "&c_isp,User_Conn,1,3
if rs_isp.eof then
	strShowErr = "<li>�Ҳ���ϵͳ������Ϣ������û��ͨ����֧�����ܡ�����ϵͳ����Ա��ϵ</li>>"
	Response.Redirect("../../lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
else
	Dim str_c_isp,str_c_user,str_c_pass,str_c_url
	str_c_isp=rs_isp("c_isp")
	str_c_user=rs_isp("c_user")
	str_c_pass=rs_isp("c_pass")
	str_c_email=rs_isp("c_undefined_1")
end if

Dim tmp_ramcode,tmp_PayMoney,Subject,body,out_trade_no,price,quantity,discount,seller_email
Dim paymethod,defaultbank,AlipayObj,itemURL
tmp_ramcode = NoSqlHack(Request.QueryString("OrderNumber"))
tmp_PayMoney=NoSqlHack(Request.QueryString("Moneys"))


Subject = tmp_ramcode	'��Ʒ���ƣ�����ͻ��߹��ﳵ���̿�����Ϊ  "�����ţ�"&request("�ͻ���վ����")
body = str_GoodsName		'��Ʒ����
out_trade_no = tmp_ramcode         '��ʱ���ȡ�Ķ�����
price = tmp_PayMoney			'price��Ʒ����	0.01��50000.00 , ע����Ҫ����3,000.00���۸�֧��","��
quantity = "1"             '��Ʒ����,����߹��ﳵĬ��Ϊ1
discount = "0"             '��Ʒ�ۿ�
seller_email = str_c_email   '���ҵ�֧�����ʺ�,c2c�ͻ������Ը��Ĵ˲�����
paymethod = "directPay"      '��ֵ:bankPay(����);cartoon(��ͨ); directPay(���)
defaultbank = "directPay"     ' ����Ĭ�ϵ�����
Set AlipayObj = New Alipay
AlipayObj.key=str_c_pass
AlipayObj.partner=str_c_user
str_c_url = "http://"&Request.ServerVariables("HTTP_HOST")&Request.ServerVariables("URL")
str_c_url = Left(str_c_url,InstrRev(str_c_url,"/"))
'AlipayObj.notify_url=str_c_url&"Alipay_Notify.asp"
AlipayObj.return_url=str_c_url&"Alipay_Return.asp"
AlipayObj.show_url="http://"&Request.ServerVariables("HTTP_HOST")
itemUrl = AlipayObj.creatURL(subject,body,out_trade_no,price,quantity,seller_email,paymethod)
'Response.Write(Server.HTMLEncode(itemUrl))
Response.Redirect(itemUrl)
%>
