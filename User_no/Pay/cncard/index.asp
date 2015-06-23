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
		strShowErr = "<li>错误的参数</li>"
		Response.Redirect("../../lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Else
		set RsPay = User_Conn.execute("select OrderId,IsPay,MoneyAmount From FS_ME_Order where OrderNumber='"&NoSqlHack(Request.QueryString("OrderNumber"))&"' and OrderID="&CintStr(Request.QueryString("OrderID"))&"")
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
		end if
	end if
	'得到ISP信息
	str_c_url = "http://"&Request.ServerVariables("HTTP_HOST")&Request.ServerVariables("URL")
	str_c_url = Left(str_c_url,InstrRev(str_c_url,"/"))&"GetPayHandle.asp"
	dim rs_isp,g_rs,c_isp
	c_isp = cint(Request("c_isp"))
	set rs_isp= Server.CreateObject(G_FS_RS)
	rs_isp.open "select c_isp,c_user,c_pass,c_url,c_gurl From FS_ME_Pay WHERE c_isp = "&c_isp,User_Conn,1,3
	if rs_isp.eof then
		strShowErr = "<li>找不到系统配置信息，或者没开通在线支付功能。请与系统管理员联系</li>>"
		Response.Redirect("../../lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		Dim str_c_isp,str_c_user,str_c_pass,str_c_url
		str_c_isp=rs_isp("c_isp")
		str_c_user=rs_isp("c_user")
		str_c_pass=rs_isp("c_pass")
	end if
	'如果为新购买用户定单号，则把定单数据插入数据库中
	'获得定单日期
	Dim tmp_OrderNumber,tmp_ramcode ,RsOrderObj,tmp_PayMoney,tmp_ymd ,tmp_ymd_ 
	tmp_ramcode = NoSqlHack(Request.QueryString("OrderNumber"))
	tmp_PayMoney=NoSqlHack(Request.QueryString("Moneys"))
	tmp_ymd=now
	tmp_ymd_=year(tmp_ymd)
	if month(now)<10 then:tmp_ymd_=tmp_ymd_&"0"&month(tmp_ymd):else:tmp_ymd_=tmp_ymd_&month(tmp_ymd):end if
	if day(now)<10 then:tmp_ymd_=tmp_ymd_&"0"&day(tmp_ymd):else:tmp_ymd_=tmp_ymd_&day(tmp_ymd):end if
'---订单信息---
	Dim c_mid			'商户编号，在申请商户成功后即可获得，可以在申请商户成功的邮件中获取该编号
	Dim c_order			'商户网站生成的订单号，不能重复
	Dim c_name			'商户订单中的收货人姓名
	Dim c_address		'商户订单中的收货人地址
	Dim c_tel			'商户订单中的收货人电话
	Dim c_post			'商户订单中的收货人邮编
	Dim c_email			'商户订单中的收货人Email
	Dim c_orderamount	'商户订单总金额
	Dim c_ymd			'商户订单的产生日期，格式为"yyyymmdd"，如20050102
	Dim c_moneytype		'支付币种，0为人民币
	Dim c_retflag		'商户订单支付成功后是否需要返回商户指定的文件，0：不用返回 1：需要返回
	Dim c_paygate		'如果在商户网站选择银行则设置该值，具体值可参见《云网支付@网技术接口手册》附录一；如果来云网支付@网选择银行此项为空值。
	Dim c_returl		'如果c_retflag为1时，该值代表支付成功后返回的文件的路径
	Dim c_memo1			'商户需要在支付结果通知中转发的商户参数一
	Dim c_memo2			'商户需要在支付结果通知中转发的商户参数二
	Dim c_signstr		'商户对订单信息进行MD5签名后的字符串
	Dim c_pass			'支付密钥，请登录商户管理后台，在帐户信息->基本信息->安全信息中的支付密钥项
	Dim notifytype		'0普通通知方式/1服务器通知方式，空值为普通通知方式
	Dim c_language		'对启用了国际卡支付时，可使用该值定义消费者在银行支付时的页面语种，值为：0银行页面显示为中文/1银行页面显示为英文
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
	c_returl	= str_c_url	'该地址为商户接收云网支付结果通知的页面，请提交完整文件名(对应范例文件：GetPayNotify.asp)
	if c_isp>0 then
	'获得会员组
		c_memo1		= c_isp
	else
		c_memo1		= ""
	end if
	c_memo2		= ""
	c_pass		= str_c_pass
	notifytype	= "0"
	c_language	= "0"
	srcStr = c_mid & c_order & c_orderamount & c_ymd & c_moneytype & c_retflag & c_returl & c_paygate & c_memo1 & c_memo2 & notifytype & c_language & c_pass
	'说明：如果您想指定支付方式(c_paygate)的值时，需要先让用户选择支付方式，然后再根据用户选择的结果在这里进行MD5加密，也就是说，此时，本页面应该拆分为两个页面，分为两个步骤完成。
'---对订单信息进行MD5加密 
	c_signstr	= MD5(srcStr,32)
'end if 
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
							<input type="submit" name="submit" value="点击 -> 云网支付@网" onclick="{if(confirm('您确认支付吗？')){this.document.form1.submit();return true;}return false;}" />
						</td>
					</tr>
					</form>
				</table>
				<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
					<tr>
						<td width="100%" class="xingmu">
							选择其他支付方式
						</td>
					</tr>
					<tr>
						<td class="hback">
							<a href="../../onlinepay.asp?OrderNumber=<%=NoSqlHack(Request("OrderNumber"))%>&Moneys=<%=NoSqlHack(Request("Moneys"))%>&OrderID=<%=NoSqlHack(Request("OrderID"))%>">在线支付</a> <a href="../../PayCenter.asp?PayStyle=MySelfAcc&OrderNumber=<%=NoSqlHack(Request("OrderNumber"))%>&Moneys=<%=NoSqlHack(Request("Moneys"))%>&OrderID=<%=NoSqlHack(Request("OrderID"))%>">帐户支付</a>
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