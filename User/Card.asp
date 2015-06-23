<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
'插入权限
User_GetParm
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-点卡冲值</title>
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
            <a href="main.asp">会员首页</a> &gt;&gt; 点卡冲值</td>
          </tr>
        </table>
        <%if Request("action")="submit" then
				Call CardSubmit()
			 Elseif Request("action")="save" then
			 	Call savedata()
			 Else
			 	Call InPutCard()
			 End if
		Sub savedata()
				Dim RsCardsaveObj,RsCardsaveSQL,p_strCardNumbersave,p_strCardPasswordsave,p_strCardPointsave,p_strCardMoneysave
				p_strCardNumbersave = NoSqlHack(Replace(Request.Form("CardNumber"),"''",""))
				p_strCardPasswordsave = NoSqlHack(Replace(Request.Form("CardPasswords"),"''",""))
				p_strCardPointsave = NoSqlHack(Replace(Request.Form("CardPoint"),"''",""))
				p_strCardMoneysave = NoSqlHack(Replace(Request.Form("CardMoney"),"''",""))
				RsCardsaveSQL = "select  CardID,CardNumber,CardPasswords,CardMoney,CardPoint,CardDateNumber,CardOverDueTime,IsUse,UserNumber,UserTime,AddTime,isBuy From FS_ME_Card where CardNumber='"& NoSqlHack(p_strCardNumbersave) &"' and CardPasswords = '"& NoSqlHack(p_strCardPasswordsave) &"'"
				Set RsCardsaveObj = server.CreateObject(G_FS_RS)
				RsCardsaveObj.Open RsCardsaveSQL,User_Conn,1,3
				if RsCardsaveObj.eof then 
					strShowErr = "<li>找不到卡号和密码</li>"
					Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
					Response.end
				Else
					if RsCardsaveObj("isUse") = 1 Or Trim(RsCardsaveObj("UserNumber")) <>""  or trim(RsCardsaveObj("UserTime"))<>"" then 
						strShowErr = "<li>此点卡已经被使用</li>"
						Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
						Response.end
					Else
						'更新点卡 
						Dim RsUpdateCardRSQL,RsUpdateCardRObj
						RsUpdateCardRSQL = "select  isuse,UserTime,isBuy,UserNumber From FS_ME_Card where CardNumber='"& NoSqlHack(p_strCardNumbersave) &"'  and CardPasswords = '"& NoSqlHack(p_strCardPasswordsave) &"'" 
						Set RsUpdateCardRObj = server.CreateObject(G_FS_RS)
						RsUpdateCardRObj.Open RsUpdateCardRSQL,User_Conn,1,3
						RsUpdateCardRObj("isuse") = 1
						RsUpdateCardRObj("UserNumber") = Fs_User.UserNumber
						RsUpdateCardRObj("UserTime") = now
						RsUpdateCardRObj("isBuy") = 1
						RsUpdateCardRObj.update
						RsUpdateCardRObj.close:set RsUpdateCardRObj = nothing
						'插入数据，以增加会员的金币或者点数或者天数
						'取消天数的冲值，暂时
						'**********************************************************
						User_conn.execute("Update FS_ME_Users set FS_Money=FS_Money+"& Clng(p_strCardMoneysave) &",Integral=Integral+"&clng(p_strCardPointsave)&" where UserNumber='"&Fs_User.UserNumber&"'")
						'**********************************************************
						dim RsUpdateOrderSQL,RsUpdateOrderObj,tmp_order
						tmp_order = year(now)&month(now)&day(now)&"-"&GetRamCode(10)
						RsUpdateOrderSQL = "select  * From FS_ME_Order where 1=0"
						Set RsUpdateOrderObj = server.CreateObject(G_FS_RS)
						RsUpdateOrderObj.Open RsUpdateOrderSQL,User_Conn,1,3
						RsUpdateOrderObj.addnew
						RsUpdateOrderObj("OrderNumber") = NoSqlHack(tmp_order)
						RsUpdateOrderObj("AddTime") = now
						RsUpdateOrderObj("IsSuccess") = 1
						RsUpdateOrderObj("UserNumber") = Fs_User.UserNumber
						RsUpdateOrderObj("OrderType") = 2
						RsUpdateOrderObj("M_PayStyle") = 4
						RsUpdateOrderObj("M_PayDate") = now 
						RsUpdateOrderObj("Content") = "冲值点卡，点卡号："& NoSqlHack(p_strCardNumbersave) &""
						RsUpdateOrderObj("isLock") = 0
						RsUpdateOrderObj.update
						RsUpdateOrderObj.close:set RsUpdateOrderObj = nothing
						''--------------------人才卡的冲值
						'dim LeftCount,LeftDateNumber ,aptmprs,addsql
						'LeftCount = 0:LeftDateNumber=0
						'set aptmprs = Conn.execute("select LeftCount from FS_AP_UserList where UserNumber='"&Fs_User.UserNumber&"' and GroupLevel=1")
						'if not aptmprs.eof then 
						'	LeftCount = aptmprs(0)
						'	''得到送点数标准
						'	
						'	dim InitCount,sysaprs
						'	InitCount = 0
						'	set sysaprs = Conn.execute("select top 1 InitCount from FS_AP_SysPara")
						'	if not sysaprs.eof then if not isnull(sysaprs(0)) then InitCount = clng(sysaprs(0))
						'	Conn.execute("update FS_AP_UserList set LeftCount=LeftCount+"&InitCount+clng(p_strCardPointsave)&" where UserNumber='"&Fs_User.UserNumber&"'")	
						'else
						'	aptmprs.close
						'	set aptmprs = Conn.execute("select BeginDate,EndDate from FS_AP_UserList where UserNumber='"&Fs_User.UserNumber&"' and GroupLevel=2")	
						'	if not aptmprs.eof then  
						'		LeftDateNumber = datediff("day",date(),aptmprs(1))+RsCardsaveObj("CardDateNumber")
						'		if G_IS_SQL_User_DB=1 then
						'			addsql = "dateadd(day,EndDate,"&RsCardsaveObj("CardDateNumber")&")"
						'		else
						'			addsql = "dateadd('day',EndDate,"&RsCardsaveObj("CardDateNumber")&")"
						'		end if
						'		Conn.execute("update FS_AP_UserList set EndDate="&addsql&" where UserNumber='"&Fs_User.UserNumber&"'")	
						'	else
						'	''vip会员
						'	''可以永久使用 不受时间和点数的限制	
						'	end if		
						'	aptmprs.close				
						'end if
						''冲值记录
						'Conn.execute("insert into FS_AP_Payment (UserNumber,PayDate,PayMoney,PayCount,LeftCount,LeftDateNumber) values ('"&Fs_User.UserNumber&"','"&now()&"',"&Clng(p_strCardMoneysave)&","&clng(p_strCardPointsave)&","&clng(LeftCount)+clng(p_strCardPointsave)&","&LeftDateNumber&") ")
						''-----------------------
						Call Fs_User.AddLog("点卡冲值",Fs_User.UserNumber,p_strCardMoneysave,p_strCardPointsave,"点卡冲值",0)
													
						strShowErr = "<li>冲值成功！定单记录为"& tmp_order &"</li>"
						Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Myaccount.asp")
						Response.end
					End if
				End if
				RsCardsaveObj.close:set RsCardsaveObj=nothing
		End sub
		sub InPutCard()
		%>
        <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="form1" method="post" action="Card.asp?action=submit">
          <tr class="hback"> 
            <td width="24%" height="32" class="hback"> <div align="right">卡号</div></td>
            <td width="76%" class="hback"><input name="CardNumber" type="text" id="CardNumber" size="40"> 
            </td>
          </tr>
          <tr class="hback"> 
            <td height="32" class="hback"><div align="right">密码</div></td>
            <td class="hback"><input name="CardPassword" type="password" id="CardPassword" size="40"></td>
          </tr>
          <tr class="hback"> 
            <td height="32" class="hback">&nbsp;</td>
            <td class="hback"><input type="submit" name="Submit" value=" 冲 值 ">
              　 
              <input type="reset" name="Submit2" value=" 重 置 "></td>
          </tr>
          <tr class="hback">
            <td height="32" class="hback">&nbsp;</td>
            <td class="hback"><a href="Pay.asp"><strong>在线银行冲值</strong></a></td>
          </tr>
        </form>
      </table>
		<%
		End sub
		Sub CardSubmit()
			if trim(Request.Form("CardNumber"))="" or  trim(Request.Form("CardPassword"))=""then
					strShowErr = "<li>请输入卡号</li><li>请输入密码</li>"
					Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
					Response.end
			End if
			Dim RsQCardObj,RsCardSQL,p_strCardNumber,p_strCardPassword
			Dim p_CardID,p_CardNumber,p_CardPassord,p_CardMoney,p_CardDateNumber
			Dim p_CardOverDueTime,p_IsUse,p_UserNumber,p_UserTime,p_AddTime,p_isBuy,p_CardPoint
			p_strCardNumber = NoSqlHack(Replace(Request.Form("CardNumber"),"''",""))
			p_strCardPassword = NoSqlHack(Replace(Request.Form("CardPassword"),"''",""))
			p_strCardPassword = Encrypt(p_strCardPassword) ''加密
			RsCardSQL = "select  CardID,CardNumber,CardPasswords,CardMoney,CardPoint,CardDateNumber,CardOverDueTime,IsUse,UserNumber,UserTime,AddTime,isBuy From FS_ME_Card where CardNumber='"& NoSqlHack(p_strCardNumber) &"' and CardPasswords = '"& NoSqlHack(p_strCardPassword) &"'"
			Set RsQCardObj = server.CreateObject(G_FS_RS)
			RsQCardObj.Open RsCardSQL,User_Conn,1,3
			if RsQCardObj.eof then
					strShowErr = "<li>无效的点卡</li><li>点卡或者密码错误</li>"
					Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
					Response.end
			 Else
					if RsQCardObj("isUse") = 1 Or Trim(RsQCardObj("UserNumber")) <>""  or trim(RsQCardObj("UserTime"))<>"" then 
						strShowErr = "<li>点卡已经被使用</li>"
						Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
						Response.end
					Elseif RsQCardObj("isUse") = 0 Or Trim(RsQCardObj("UserNumber")) =""  or trim(RsQCardObj("UserTime"))="" then
						'定义日期
						Dim strCard_month,strCard_day 
						if Len(Cstr(month(RsQCardObj("CardOverDueTime"))))<2  then  
							strCard_month = "0"&month(RsQCardObj("CardOverDueTime"))
						Else
							strCard_month = month(RsQCardObj("CardOverDueTime"))
						End  if 
						if Len(Cstr(day(RsQCardObj("CardOverDueTime"))))<2 then
							strCard_day = "0"&day(RsQCardObj("CardOverDueTime"))
						Else
							strCard_day = day(RsQCardObj("CardOverDueTime"))
						End  if
						if clng(right(year(RsQCardObj("CardOverDueTime")),2)&strCard_month&strCard_day)< clng(strTodaydate)  then
							strShowErr = "<li>点卡已经过期</li>"
							Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
							Response.end
						Else
							p_CardID  = RsQCardObj("CardID")
							p_CardNumber  = RsQCardObj("CardNumber")
							p_CardPassord  = RsQCardObj("CardPasswords")
							if trim(RsQCardObj("CardMoney")) <> "" then
								p_CardMoney  = RsQCardObj("CardMoney")
							Else
								p_CardMoney  = 0
							End if
							if trim(RsQCardObj("CardPoint")) <> "" then
								p_CardPoint  = RsQCardObj("CardPoint")
							Else
								p_CardPoint  = 0
							End if
							if trim(RsQCardObj("CardDateNumber")) <> "" then
								p_CardDateNumber  = RsQCardObj("CardDateNumber")
							Else
								p_CardDateNumber  = 0
							End if
							p_CardOverDueTime  = RsQCardObj("CardOverDueTime")
							p_IsUse  = RsQCardObj("isUse")
							p_UserNumber  = RsQCardObj("UserNumber")
							p_UserTime  = RsQCardObj("UserTime")
							p_AddTime  = RsQCardObj("AddTime")
							p_isBuy  = RsQCardObj("isBuy")
						End if
					End if
			 End if
			 RsQCardObj.close:Set  RsQCardObj = nothing
		%>
        <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="form1" method="post" action="Card.asp?action=save">
          <tr class="hback"> 
            <td width="24%" height="22" class="hback"> <div align="right">点卡面值</div></td>
            <td width="76%" class="hback">
			<font style="font-size:20px;color:red;"><i><b><% = p_CardMoney %></b></i></font>&nbsp; <%=p_MoneyName%>
              <input name="CardMoney" type="hidden" id="CardMoney" value="<% = p_CardMoney%>">
              <input name="CardNumber" type="hidden" id="CardNumber" value="<% = p_CardNumber%>">
              <input name="CardPasswords" type="hidden" id="CardPasswords" value="<% = p_CardPassord%>">
            </td>
          </tr>
          <tr class="hback"> 
            <td height="22" class="hback"> <div align="right">点数</div></td>
            <td class="hback"><font style="font-size:20px;color:red;"><i><b>
<% = p_CardPoint %></b></i></font>&nbsp; 点 
              <input name="CardPoint" type="hidden" id="CardPoint" value="<% = p_CardPoint%>"></td>
          </tr>
          <tr class="hback"> 
            <td height="3" class="hback"><div align="right">天数</div></td>
            <td class="hback"><font style="font-size:20px;color:red;"><i><b>
<% = p_CardDateNumber %></b></i></font>&nbsp; 天 
              <input name="CardDateNumber" type="hidden" id="CardDateNumber" value="<% = p_CardDateNumber%>"></td>
          </tr>
          <tr class="hback"> 
            <td height="3" class="hback"><div align="right">过期日期</div></td>
            <td class="hback">
<% = p_CardOverDueTime %>
              <input name="CardOverDueTime" type="hidden" id="CardOverDueTime" value="<% = p_CardOverDueTime %>"></td>
          </tr>
          <tr class="hback"> 
            <td height="32" class="hback">&nbsp;</td>
            <td class="hback"><input type="submit" name="Submit" value="确认冲值"> 
            </td>
          </tr>
        </form>
      </table>
		<%End Sub%>
      </td>
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





