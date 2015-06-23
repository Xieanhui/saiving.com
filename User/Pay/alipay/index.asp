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
	strShowErr = "<li>错误的参数</li>"
	Response.Redirect("../../lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
Else
	set RsPay = User_Conn.execute("select OrderId,IsPay,MoneyAmount,Content,AddTime From FS_ME_Order where OrderNumber='"&NoSqlHack(Request.QueryString("OrderNumber"))&"' and OrderID="&CintStr(Request.QueryString("OrderID"))&"")
	if RsPay.EOF then
		strShowErr = "<li>错误的参数，找不到定单记录</li>"
		'Response.Redirect("../../lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Else
		if CCur(RsPay("MoneyAmount")) <> CCur(Request.QueryString("Moneys")) then
			strShowErr = "<li>支付金额和实际金额不符</li>"
			Response.Redirect("../../lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		if RsPay("IsPay")=1 then
			strShowErr = "<li>此定单已经支付，不能再支付</li>"
			Response.Redirect("../../lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		if RsPay("MoneyAmount")=0 then
			strShowErr = "<li>此定单金额为0，不需要支付</li>"
			Response.Redirect("../../lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		str_GoodsName =  RsPay("Content")
	end if
end if
'得到ISP信息
dim rs_isp,g_rs,c_isp,str_c_email
c_isp = cint(Request("c_isp"))
set rs_isp= Server.CreateObject(G_FS_RS)
rs_isp.open "select c_isp,c_user,c_pass,c_undefined_1 From FS_ME_Pay WHERE c_isp = "&c_isp,User_Conn,1,3
if rs_isp.eof then
	strShowErr = "<li>找不到系统配置信息，或者没开通在线支付功能。请与系统管理员联系</li>>"
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


Subject = tmp_ramcode	'商品名称，如果客户走购物车流程可以设为  "订单号："&request("客户网站订单")
body = str_GoodsName		'商品描述
out_trade_no = tmp_ramcode         '按时间获取的订单号
price = tmp_PayMoney			'price商品单价	0.01～50000.00 , 注：不要出现3,000.00，价格不支持","号
quantity = "1"             '商品数量,如果走购物车默认为1
discount = "0"             '商品折扣
seller_email = str_c_email   '卖家的支付宝帐号,c2c客户，可以更改此参数。
paymethod = "directPay"      '赋值:bankPay(网银);cartoon(卡通); directPay(余额)
defaultbank = "directPay"     ' 网银默认的银行
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
