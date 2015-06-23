<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
	<title>欢迎用户<%=Fs_User.UserName%>来到<%=GetUserSystemTitle%>-在线银行支付</title>
	<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
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
				<%
		if Request("action")="submit" then
			Call EndPay()
		Else
			Call PaySelect()
		End if
		sub PaySelect()
		if trim(Request.Form("GroupID"))<>"" then
			dim GroupMoney_,GroupTF,GroupName_,GroupID_
			set g_rs= Server.CreateObject(G_FS_RS)
			g_rs.open "select GroupID,GroupName,GroupMoney From FS_ME_Group Where GroupID="&CintStr(Request.Form("GroupID")),User_Conn,1,3
			GroupMoney_=formatNumber(g_rs("GroupMoney"),2,-1)
			GroupTF=1
			GroupID_=g_rs("GroupID")
			g_rs.close:set g_rs=nothing
		else
			GroupMoney_=10
			GroupTF=0
		end if
		
				%>
				<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
					<form name="form1" method="post" action="pay.asp?action=submit">
					<tr class="hback">
						<td width="15%" height="32" class="hback">
							<div align="right">
								您冲值的金额</div>
						</td>
						<td width="60%" class="hback">
							<input name="PayMoney" type="text" id="PayMoney" value="<%=GroupMoney_%>">
							人民币
							<input name="OrderID" type="hidden" id="OrderID" value="<% = NoSqlHack(Request.QueryString("OrderID"))%>">
						</td>
						<td width="25%" rowspan="3" class="hback">
							<a href="https://www.cncard.net/merchant.asp?pmid=1008687" target="_blank">
								<img src="../sys_images/cncard_logo.gif" alt="点击开始注册" width="184" height="40" border="0"></a>
						</td>
					</tr>
					<tr class="hback">
						<td height="32" class="hback">
							<div align="right">
								定单号</div>
						</td>
						<td class="hback">
							<input name="OrderNumber" type="text" id="OrderNumber" value="<%=year(now)&month(now)&day(now)&"-"&GetRamCode(8)%>" readonly="true">
							请记住此定单号，以方便查询
						</td>
					</tr>
					<tr class="hback">
						<td height="32" class="hback">
							<div align="right">
								您的支付方式</div>
						</td>
						<td class="hback">
							<select name="c_isp" id="c_isp">
								<option value="0">支付宝</option>
								<option value="1">云网支付@网</option>
							</select>
							<input name="GroupTF" type="hidden" value="<%=GroupTF%>">
							<input name="GroupID" type="hidden" value="<%=GroupID_%>">
						</td>
					</tr>
					<tr class="hback">
						<td height="32" class="hback">
							&nbsp;
						</td>
						<td colspan="2" class="hback">
							<input type="submit" name="Submit" value="确认支付金额">
							<input type="reset" name="Submit2" value=" 重置 ">
						</td>
					</tr>
					</form>
				</table>
				
				<%
		End sub
		Sub EndPay()
		
		if trim(Request.Form("PayMoney"))="" or  IsNumeric(trim(Request.Form("PayMoney")))=false then
				strShowErr = "<li>请填写金额</li><li>您输入的金额不合法</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
		End if
		Dim str_c_isp,moneycount,orderid
		moneycount=NoSqlHack(Request.Form("PayMoney"))
		str_c_isp = Request.Form("c_isp")
		if str_c_isp = "" then
				strShowErr = "<li>请选择支付ISP商</li>>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
		End if
		'如果为新购买用户定单号，则把定单数据插入数据库中
		Dim tmp_OrderNumber,tmp_ramcode ,RsOrderObj,tmp_PayMoney,tmp_ymd ,tmp_ymd_ 
		tmp_ramcode = NoSqlHack(Request.Form("OrderNumber"))
		tmp_PayMoney=NoSqlHack(Request.Form("PayMoney"))
		tmp_ymd=now
		tmp_ymd_=year(tmp_ymd)
		if month(now)<10 then:tmp_ymd_=tmp_ymd_&"0"&month(tmp_ymd):else:tmp_ymd_=tmp_ymd_&month(tmp_ymd):end if
		if day(now)<10 then:tmp_ymd_=tmp_ymd_&"0"&day(tmp_ymd):else:tmp_ymd_=tmp_ymd_&day(tmp_ymd):end if
		Set RsOrderObj = Server.CreateObject(G_FS_RS)
		RsOrderObj.Open "Select * From FS_ME_Order where OrderNumber='"& tmp_ramcode &"'",User_Conn,1,3
		if Not RsOrderObj.eof then
			strShowErr = "<li>定单号意外重复，请重新购买</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
			RsOrderObj.update
			RsOrderObj.close:set RsOrderObj = nothing
			set conn=nothing
			set user_conn=nothing
			set fs_user=nothing
		Else
			RsOrderObj.addnew
			RsOrderObj("OrderNumber") = tmp_ramcode
			RsOrderObj("OrderType") = 3
			RsOrderObj("MoneyAmount") = moneycount
			RsOrderObj("AddTime") = tmp_ymd
			RsOrderObj("IsSuccess") = 0
			RsOrderObj("isLock") = 0
			RsOrderObj("M_state") = 0
			if trim(Request.Form("GroupTF"))="1" then
				set g_rs= Server.CreateObject(G_FS_RS)
				g_rs.open "select GroupID,GroupName,GroupMoney From FS_ME_Group Where GroupID="&CintStr(Request.Form("GroupID")),User_Conn,1,3
				RsOrderObj("Content") ="会员续费冲金币,会员组名称:"&g_rs("GroupName")&",GroupId:"&g_rs("GroupID")&""
			else
				RsOrderObj("Content") ="会员直冲金币"
			end if
			RsOrderObj("UserNumber") = Fs_User.UserNumber
			RsOrderObj.update
			if G_IS_SQL_User_DB = "0" then
				orderid= RsOrderObj("OrderId")
			else
				Dim rssql
				set rssql = Conn.execute("SELECT ident_current('FS_ME_Order')")
				orderid = rssql(0)
				rssql.close:set rssql = nothing
			end if
			RsOrderObj.close:set RsOrderObj = nothing
			Dim payUrl
			select case str_c_isp
				case "0"
					payUrl="pay/alipay/index.asp"
					
				case "1"
					payUrl="pay/cncard/index.asp"
			end select
			payUrl=payUrl&"?c_isp="&str_c_isp&"&OrderNumber="&tmp_ramcode&"&Moneys="&moneycount&"&OrderID="&orderid
			Response.Redirect(payUrl)
		End if
		End Sub%>
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
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->
