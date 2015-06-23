<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
dim rsOrder
on error resume next
	set rsOrder = User_Conn.execute("select OrderID From FS_ME_Order where OrderNumber='"&NoSqlHack(Request("OrderNumber"))&"' and MoneyAmount="& CintStr(Request.QueryString("Moneys"))&" and UserNumber='"&Fs_User.UserNumber&"' and OrderID="&CintStr(Request("OrderID"))&"")
	if rsOrder.eof then
		response.Write "找不到记录！"
		Response.end
	end if
'开始帐户支付
if Request.Form("Action")="save" then
	'得到定单信息
	Dim rsPay
	set rsPay = User_Conn.execute("select UserNumber,MoneyAmount,isPay From FS_ME_Order where OrderID="&CintStr(Request.Form("OrderID")))
	if not rsPay.eof then
		if RsPay("isPay")=1 then
			strShowErr = "<li>定单已经支付，不能再支付!</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		if rsPay("MoneyAmount")<0 then
			strShowErr = "<li>您的定单信息有错误,请不要用非法途径获取商品!</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		if Fs_User.NumFS_Money < rsPay("MoneyAmount") then
			strShowErr = "<li>您的金币不足!</li><li><a href=""pay.asp"">点击此处为帐户冲值！</a></li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		else
			'更新会员金币
			User_Conn.execute("Update FS_ME_Users set FS_Money=FS_Money-"&rsPay("MoneyAmount")&" where UserNumber='"&Fs_User.UserNumber&"'")
			'更新定单状态
			User_Conn.execute("Update FS_ME_Order set IsSuccess=1,IsPay=1 where UserNumber='"&Fs_User.UserNumber&"' and OrderId="&CintStr(Request.Form("OrderID")))
		end if
		strShowErr = "<li>支付成功!</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Order.asp")
		Response.end
	else
		strShowErr = "<li>错误参数!</li>"
		Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Order.asp")
		Response.end
	end if
	set rsPay = nothing
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-支付中心</title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,风讯,风讯网站内容管理系统,风讯系统,风讯新闻系统,风讯商城,风讯b2c,新闻系统,CMS,域名空间,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head>
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
      <td   colspan="2" class="xingmu" height="26"> <!--#include file="Top_navi.asp" --> </td>
    </tr>
    <tr class="back"> 
      <td width="18%" valign="top" class="hback"> <div align="left"> 
          <!--#include file="menu.asp" -->
        </div></td>
      <td width="82%" valign="top" class="hback"><table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
          <tr class="hback"> 
            
          <td class="hback"><strong>位置：</strong><a href="../">网站首页</a> &gt;&gt; 
            <a href="main.asp">会员首页</a> &gt;&gt; 支付中心</td>
          </tr>
        </table>
		<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
          <tr>
            <td colspan="2" class="xingmu">以下是您的定单信息</td>
          </tr>
          <tr class="hback">
            <td width="20%"><div align="right">定单编号</div></td>
            <td width="80%"><%=NoSqlHack(Request("OrderNumber"))%></td>
          </tr>
          <tr class="hback">
            <td><div align="right">产生金额</div></td>
            <td><%=NoSqlHack(Request("Moneys"))%>RMB</td>
          </tr>
          <tr class="hback">
            <td colspan="2"><div align="center" class="tx">请您在办理完成相关支付手续后凭定单号与我们客服人员联系</div></td>
          </tr>
        </table>
		<%
		select Case NoSqlHack(Request("PayStyle"))
			Case "PostOrBank"
				Call PostOrBank()
			Case "MySelfAcc"
				Call MySelfAcc()
			Case "Card"
				Call Card()
			Case Else
				Call MySelfAcc()
		end Select
		Sub PostOrBank()
		%>
        <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
          <tr>
            <td colspan="2" class="xingmu">您选择是的给我们电汇或者邮局汇款，请按照以下方式给我们汇款</td>
          </tr>
          <tr class="hback">
            <td width="20%" height="27" class="hback_1"><div align="right">邮局汇款资料</div></td>
            <td width="80%">
			<%
			dim rsMall
			if IsExist_SubSys("MS") Then
			 set rsMall = Conn.execute("select top 1 Address,Content,PostCode,UserName From FS_MS_PayMethod")
			 if not rsMall.eof then
				 Response.Write "地址："& rsMall("Address") &"<br />"
				 Response.Write "收件人："& rsMall("UserName") &"<br />"
				 Response.Write "邮政编码："& rsMall("PostCode") &"<br />"
			 else
				 Response.Write "如果您没开通商城子系统，请手工再此输入"
			 end if
			 rsMall.close:set rsMall = nothing
			%>
			<%else%>
			如果您没开通商城子系统，请手工再此输入
			<%end if%>			</td>
          </tr>
          <tr class="hback">
            <td height="28" class="hback_1"><div align="right">银行电汇资料</div></td>
            <td>
			<%
			if IsExist_SubSys("MS") Then
			 set rsMall = Conn.execute("select top 1 Other From FS_MS_PayMethod")
			 if not rsMall.eof then
				 Response.Write ""& rsMall("Other") &"<br />"
			 else
				 Response.Write "如果您没开通商城子系统，请手工再此输入"
			 end if
			 rsMall.close:set rsMall = nothing
			%>
			<%else%>
			如果您没开通商城子系统，请手工再此输入
			<%end if%>
			</td>
          </tr>
          <tr class="hback">
            <td height="28" class="hback_1"><div align="right">配送须知</div></td>
            <td>
			<%
			if IsExist_SubSys("MS") Then
			 set rsMall = Conn.execute("select top 1 Content From FS_MS_PayMethod")
			 if not rsMall.eof then
				 Response.Write ""& rsMall("Content") &"<br />"
			 else
				 Response.Write "如果您没开通商城子系统，请手工再此输入"
			 end if
			 rsMall.close:set rsMall = nothing
			%>
			<%else%>
			如果您没开通商城子系统，请手工再此输入
			<%end if%>
			</td>
          </tr>
        </table>
		<%end sub%>
		<%sub MySelfAcc()%>
		<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
          <tr>
            <td width="100%" class="xingmu">您选择的是帐户支付</td>
          </tr>
          <tr>
            <form name="form1" method="post" action="">
              <td class="hback"><input type="button" name="Submit" value="您需要支付的金币为：<%=NoSqlHack(Request("Moneys"))%>。您确认支付吗。支付成功后将扣除您的相应金币" onClick="{if(confirm('您确定支付吗?')){this.document.form1.submit();return true;}return false;}">
              <input name="Action" type="hidden" id="Action" value="save">
              <input name="OrderNumber" type="hidden" id="OrderNumber" value="<%=NoSqlHack(Request("OrderNumber"))%>">
              <input name="Moneys" type="hidden" id="Moneys" value="<%=NoSqlHack(Request("Moneys"))%>">
              <input name="OrderID" type="hidden" id="OrderID" value="<%=NoSqlHack(Request("OrderID"))%>"></td>
            </form>
          </tr>
        </table>
		<%end sub%>
		<%sub Card()%>
		<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
          <tr>
            <td width="100%" class="xingmu">点卡支付</td>
          </tr>
          <tr>
            <td class="hback">点卡定单支付功能暂时没开通。如果您需要使用点卡支付，请先购买点卡冲值入您的帐户后，用帐户购买。<a href="card.asp">[点卡冲值]</a> </td>
          </tr>
        </table>
		<%end sub%>
	    <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
          <tr>
            <td width="100%" class="xingmu">选择其他支付方式</td>
          </tr>
          <tr>
            <td class="hback"><a href="onlinepay.asp?OrderNumber=<%=NoSqlHack(Request("OrderNumber"))%>&Moneys=<%=NoSqlHack(Request("Moneys"))%>&OrderID=<%=NoSqlHack(Request("OrderID"))%>">在线支付</a> 　<a href="PayCenter.asp?PayStyle=MySelfAcc&OrderNumber=<%=NoSqlHack(Request("OrderNumber"))%>&Moneys=<%=NoSqlHack(Request("Moneys"))%>&OrderID=<%=NoSqlHack(Request("OrderID"))%>">帐户支付</a> </td>
          </tr>
        </table></td>
    </tr>
    <tr class="back"> 
      <td height="20"  colspan="2" class="xingmu"> <div align="left"> 
          <!--#include file="Copyright.asp" -->
        </div></td>
    </tr>
</table>
</body>
</html>
<%
Set Fs_User = Nothing
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->





