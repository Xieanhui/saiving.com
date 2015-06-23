<% 'Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
	'得到ISP信息
	dim rs_isp
	set rs_isp= Server.CreateObject(G_FS_RS)
	rs_isp.open "select top 1 c_isp,c_user,c_pass,c_url,c_gurl From FS_ME_Pay",User_Conn,1,3
	if rs_isp.eof then
		strShowErr = "<li>找不到系统配置信息，或者没开通在线支付功能。请与系统管理员联系</li>>"
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
	<title>欢迎用户<%=Fs_User.UserName%>来到<%=GetUserSystemTitle%>-在线银行支付</title>
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
							<strong>位置：</strong><a href="../">网站首页</a> &gt;&gt; <a href="main.asp">会员首页</a> &gt;&gt; 在线银行冲值
						</td>
					</tr>
				</table>
				<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
					<tr>
						<td class="hback">
							<%
		Dim c_mid,c_order,c_orderamount,c_ymd,c_transnum,c_succmark,c_moneytype,c_cause,c_memo1,c_signstr,srcStr,r_signstr
		c_mid			= NoSqlHack(request("c_mid"))			'商户编号，在申请商户成功后即可获得，可以在申请商户成功的邮件中获取该编号
		c_order			= NoSqlHack(request("c_order"))		'商户提供的订单号
		c_orderamount	= NoSqlHack(request("c_orderamount"))	'商户提供的订单总金额，以元为单位，小数点后保留两位，如：13.05
		c_ymd			= NoSqlHack(request("c_ymd"))			'商户传输过来的订单产生日期，格式为"yyyymmdd"，如20050102
		c_transnum		= NoSqlHack(request("c_transnum"))		'云网支付网关提供的该笔订单的交易流水号，供日后查询、核对使用；
		c_succmark		= NoSqlHack(request("c_succmark"))		'交易成功标志，Y-成功 N-失败			
		c_moneytype		= NoSqlHack(request("c_moneytype"))	'支付币种，0为人民币
		c_cause			= NoSqlHack(request("c_cause"))		'如果订单支付失败，则该值代表失败原因		
		c_memo1			= NoSqlHack(request("c_memo1"))		'商户提供的需要在支付结果通知中转发的商户参数一
		c_memo2			= NoSqlHack(request("c_memo2"))		'商户提供的需要在支付结果通知中转发的商户参数二
		c_signstr		= NoSqlHack(request("c_signstr"))		'云网支付网关对已上信息进行MD5加密后的字符串 
		if c_mid="" or c_order="" or c_orderamount="" or c_ymd="" or c_moneytype="" or c_transnum="" or c_succmark="" or c_signstr="" then
			strShowErr = "<li>支付失败</li><li>支付信息有误</li>>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		Dim c_pass	'商户的支付密钥，登录商户管理后台(https://www.cncard.net/admin/)，在管理首页可找到该值
		c_pass = str_c_pass
		srcStr = c_mid & c_order & c_orderamount & c_ymd & c_transnum & c_succmark & c_moneytype & c_memo1 & c_memo2 & c_pass
		r_signstr	= MD5(srcStr)
		if r_signstr<>c_signstr then  
			strShowErr = "<li>支付失败</li><li>签名验证失败</li>>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		Dim MerchantID	'商户自己的编号 
		if str_c_user<>c_mid then 
			strShowErr = "<li>支付失败</li><li>提交的商户编号有误</li>>"
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
			strShowErr = "<li>支付失败</li><li>未找到该订单信息</li>>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		Dim r_orderamount
		r_orderamount=rs_1("MoneyAmount")
		if  ccur(r_orderamount)<>ccur(c_orderamount) then
			strShowErr = "<li>支付失败</li><li>支付金额有误</li>>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		Dim r_ymd
		r_ymd=rs_1("AddTime")
		r_ymd=year(r_ymd)
		if month(rs_1("AddTime"))<10 then:r_ymd=r_ymd&"0"&month(rs_1("AddTime")):else:r_ymd=r_ymd&month(rs_1("AddTime")):end if
		if day(rs_1("AddTime"))<10 then:r_ymd=r_ymd&"0"&day(rs_1("AddTime")):else:r_ymd=r_ymd&day(rs_1("AddTime")):end if
		'暂时屏蔽
		'if  r_ymd<>c_ymd then 
		'	strShowErr = "<li>支付失败</li><li>订单时间有误</li>>"
		'	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		'	Response.end
		'end if
		'Dim r_memo1
		'r_memo1 = rs("转发参数一")
		'Dim r_memo2
		'r_memo2 = rs("转发参数二")
		'IF r_memo1<>c_memo1 or r_memo2<>c_memo2 THEN
		'	response.write "参数提交有误"
		'	response.end
		'END IF
		if c_succmark<>"Y" and c_succmark<>"N" then
			strShowErr = "<li>支付失败</li><li>参数提交有误</li>>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		IF c_succmark="Y" THEN
			dim RsPay
			Set RsPay = User_Conn.execute("select isPay,isOnPay From FS_ME_Order where OrderNumber='"& NoSqlHack(c_order) &"' and UserNumber='"&Fs_User.UserNumber&"'")
			if CSTR(RsPay("IsPay"))="1" then
				RsPay.close:set RsPay = nothing
				strShowErr = "<li>支付失败</li><li>不能重复支付您的定单</li>>"
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
					Call Fs_User.AddLog("在线支付",Fs_User.UserNumber,0,c_orderamount,"在线支付获得金币:"&c_orderamount&"",0)
					Response.Write("<script>alert(""交易成功：\n支付成功:定单号:"& c_order &"。\n交易金额:"&c_orderamount&"RMB"");location.href=""Order_Pay.asp"";</script>")
					Response.End
				elseif RsPay("isOnPay") =0 then
					User_Conn.execute("Update FS_ME_Users set FS_Money=FS_Money+"&c_orderamount&" where UserNumber='"&Fs_User.UserNumber&"'")
					Call Fs_User.AddLog("在线支付",Fs_User.UserNumber,0,c_orderamount,"在线支付定单:"&c_orderamount&"",0)
					Response.Write("<script>alert(""交易成功：\n支付成功:定单号:"& c_order &"。\n交易金额:"&c_orderamount&"RMB等额的金币已经冲入您的帐户中\n如果您是定单支付，请用帐户支付您的定单"");location.href=""Order_Pay.asp"";</script>")
					Response.End
				else
					Response.Write("<script>alert(""未知错误"");location.href=""Order_Pay.asp"";</script>")
					Response.End
				end if
				RsPay.close:set RsPay = nothing
			end if
		elseif c_succmark="N" then
			strShowErr = "<li>交易失败</li>未知原因</li><li>您在您的定单管理中查询定单号</li>"
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