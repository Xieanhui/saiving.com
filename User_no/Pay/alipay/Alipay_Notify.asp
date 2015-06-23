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
message = now()&":开始同步状态"&Request.Form
'得到ISP信息
dim rs_isp
set rs_isp= Server.CreateObject(G_FS_RS)
rs_isp.open "select c_isp,c_user,c_pass,c_undefined_1 From FS_ME_Pay WHERE c_isp = 0",User_Conn,1,3
if rs_isp.eof Then
	AlipayObj.Log(message&vbcrlf&"未找到配置信息")
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
	out_trade_no	= AlipayObj.DelStr(Request("out_trade_no")) '获取定单号
	total_fee		= AlipayObj.DelStr(Request("total_fee")) '获取支付的总价格
	receive_name    = AlipayObj.DelStr(Request("receive_name"))   '获取收货人姓名
	receive_address = AlipayObj.DelStr(Request("receive_address")) '获取收货人地址
	receive_zip     = AlipayObj.DelStr(Request("receive_zip"))   '获取收货人邮编
	receive_phone   = AlipayObj.DelStr(Request("receive_phone")) '获取收货人电话
	receive_mobile  = AlipayObj.DelStr(Request("receive_mobile")) '获取收货人手机
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
		AlipayObj.Log(message&vbcrlf&"数据库操作失败")
	else
		response.Write("success")
		AlipayObj.Log(message&vbcrlf&"支付成功")
	end if
else
	response.Write("fail")
	AlipayObj.Log(message&vbcrlf&"验证消息来源失败")
end if
%>
