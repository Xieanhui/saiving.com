<% Option Explicit %>
<!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../../FS_Inc/Function.asp" -->
<!--#include file="../../lib/strlib.asp" -->
<!--#include file="../../lib/UserCheck.asp" -->
<!--#include file="Alipay_Lib.asp"-->
<%
'on error resume next
Dim AlipayObj,str_c_url,state,message
Set AlipayObj = New Alipay
message = now()&":��ʼͬ��״̬"&Request.Form
'�õ�ISP��Ϣ
dim rs_isp
set rs_isp= Server.CreateObject(G_FS_RS)
rs_isp.open "select c_isp,c_user,c_pass,c_undefined_1 From FS_ME_Pay WHERE c_isp = 0",User_Conn,1,3
if rs_isp.eof Then
	AlipayObj.Log(message&vbcrlf&"δ�ҵ�������Ϣ")
	Response.Write("fail")
	Response.end
else
	Dim str_c_user,str_c_pass,str_c_email
	str_c_user=rs_isp("c_user")
	str_c_pass=rs_isp("c_pass")
	str_c_email=rs_isp("c_undefined_1")
end if


AlipayObj.key=str_c_pass
AlipayObj.partner=str_c_user
state = AlipayObj.Notify()
if state = "success" then
	Dim out_trade_no,total_fee,receive_name,receive_address,receive_zip,receive_phone,receive_mobile
	out_trade_no	= AlipayObj.DelStr(Request("out_trade_no")) '��ȡ������
	total_fee		= AlipayObj.DelStr(Request("total_fee")) '��ȡ֧�����ܼ۸�
	receive_name    = AlipayObj.DelStr(Request("receive_name"))   '��ȡ�ջ�������
	receive_address = AlipayObj.DelStr(Request("receive_address")) '��ȡ�ջ��˵�ַ
	receive_zip     = AlipayObj.DelStr(Request("receive_zip"))   '��ȡ�ջ����ʱ�
	receive_phone   = AlipayObj.DelStr(Request("receive_phone")) '��ȡ�ջ��˵绰
	receive_mobile  = AlipayObj.DelStr(Request("receive_mobile")) '��ȡ�ջ����ֻ�
	dim rs_1,sql_1
	sql_1="select OrderNumber,MoneyAmount,AddTime from FS_ME_Order where OrderNumber='"& out_trade_no &"' and UserNumber='"&Fs_User.UserNumber&"'"
	set rs_1=server.CreateObject(G_FS_RS)
	rs_1.open sql_1,User_conn,1,3
	if rs_1.eof then
		response.Write("fail")
	else if rs_1("isPay")="0" then
		rs_1("M_PayDate")=now
		rs_1("M_state")=1
		rs_1("isPay")=1
		rs_1("isLock")=0
		rs_1("IsSuccess")=1
		rs_1.update()
	end if
	if Err then
		response.Write("fail")
		AlipayObj.Log(message&vbcrlf&"���ݿ����ʧ��")
	else
		response.Write("success")
		AlipayObj.Log(message&vbcrlf&"֧���ɹ�")
	end if
else
	response.Write("fail")
	AlipayObj.Log(message&vbcrlf&"��֤��Ϣ��Դʧ��")
end if
%>
