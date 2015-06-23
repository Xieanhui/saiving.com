<% Option Explicit %>
<!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../../FS_Inc/Function.asp" -->
<!--#include file="../../lib/strlib.asp" -->
<!--#include file="../../lib/UserCheck.asp" -->
<!--#include file="Alipay_Lib.asp"-->
<%
Dim AlipayObj,str_c_url,state
'得到ISP信息
dim rs_isp
set rs_isp= Server.CreateObject(G_FS_RS)
rs_isp.open "select c_isp,c_user,c_pass,c_undefined_1 From FS_ME_Pay WHERE c_isp = 0",User_Conn,1,3
if rs_isp.eof then
	strShowErr = "<li>找不到系统配置信息，或者没开通在线支付功能。请与系统管理员联系</li>>"
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
	strShowErr = "<li>消息来源不正确</li>"
	Response.Redirect("../../lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
else
	Dim out_trade_no,total_fee
	out_trade_no = AlipayObj.DelStr(Request("out_trade_no"))  '获取定单号
	total_fee    = AlipayObj.DelStr(Request("total_fee"))     '获取支付的总价格
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
	<title>欢迎用户<%=Fs_User.UserName%>来到<%=GetUserSystemTitle%>-在线银行支付</title>
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
							<strong>位置：</strong><a href="../">网站首页</a> &gt;&gt; <a href="main.asp">会员首页</a> &gt;&gt; 在线银行冲值
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
					strShowErr = "<li>定单不存在</li>"
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
						Call Fs_User.AddLog("在线支付",Fs_User.UserNumber,0,total_fee,"在线支付获得金币:"&total_fee&"",0)
						Response.Write("<script>alert(""交易成功：\n支付成功:定单号:"& out_trade_no &"。\n交易金额:"&total_fee&"RMB"");location.href=""Order_Pay.asp"";</script>")
						Response.End
					elseif rs_u("isOnPay") =0 then
						User_Conn.execute("Update FS_ME_Users set FS_Money=FS_Money+"&total_fee&" where UserNumber='"&Fs_User.UserNumber&"'")
						Call Fs_User.AddLog("在线支付",Fs_User.UserNumber,0,total_fee,"在线支付定单:"&total_fee&"",0)
						Response.Write("<script>alert(""交易成功：\n支付成功:定单号:"& out_trade_no &"。\n交易金额:"&total_fee&"RMB等额的金币已经冲入您的帐户中\n如果您是定单支付，请用帐户支付您的定单"");location.href=""../../Order_Pay.asp"";</script>")
						Response.End
					else
						Response.Write("<script>alert(""未知错误"");location.href=""Order_Pay.asp"";</script>")
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
key = ""      '支付宝安全教研码
partner      = ""  '支付宝合作id

out_trade_no = DelStr(Request("out_trade_no"))  '获取定单号
total_fee    = DelStr(Request("total_fee"))     '获取支付的总价格
'如需获取其它参数，可填写 参数 =DelStr(Request.Form("获取参数名"))

'**********************判断消息是不是支付宝发出********************
alipayNotifyURL = "http://notify.alipay.com/trade/notify_query.do?"
alipayNotifyURL = alipayNotifyURL & "partner=" & partner & "&notify_id=" & Request("notify_id")
Set Retrieval   = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
Retrieval.setOption 2, 13056
Retrieval.open "GET", alipayNotifyURL, False, "", ""
Retrieval.send()
ResponseTxt   = Retrieval.ResponseText
Set Retrieval = Nothing
'*******************************************************************

'*******获取支付宝GET过来通知消息,判断消息是不是被修改过************

For Each varItem in Request.QueryString
	mystr = varItem & "=" & Request(varItem) & "^" & mystr
Next

If mystr <> "" Then
	mystr = Left(mystr,Len(mystr) - 1)
End If

mystr  = Split(mystr, "^")
Count  = UBound(mystr)
'对参数排序

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

'构造md5摘要字符串

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
	Response.Write "付款成功页面"        '这里可以指定你需要显示的内容
	' 如果您申请了支付宝的购物卷功能，请在返回的信息里面不要做金额的判断，否则会出现校验通不过，出现调单。如果您需要获取买家所使用购物卷的金额,
	' 请获取返回信息的这个字段discount的值，取绝对值，就是买家付款优惠的金额。即 原订单的总金额=买家付款返回的金额total_fee +|discount|.
Else
	Response.Write "跳转失败"          '这里可以指定你需要显示的内容
End If

Function DelStr(Str)

	If IsNull(Str) Or IsEmpty(Str) Then
		Str = ""
	End If

	DelStr = Replace(Str,";","")
	DelStr = Replace(DelStr,"'","")
	DelStr = Replace(DelStr,"&","")
	DelStr = Replace(DelStr," ","")
	DelStr = Replace(DelStr,"　","")
	DelStr = Replace(DelStr,"%20","")
	DelStr = Replace(DelStr,"--","")
	DelStr = Replace(DelStr,"==","")
	DelStr = Replace(DelStr,"<","")
	DelStr = Replace(DelStr,">","")
	DelStr = Replace(DelStr,"%","")
End Function

%>