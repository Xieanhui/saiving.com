<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
Dim obj_buy_rs,obj_buySQL_,tmp_tf_
If Request.QueryString("Action_1") = "Update" Then
	If Request("id") = "" then
		strShowErr = "<li>请选择一个商品更新</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Else
		if Not isnumeric(Request.Form("ProductNum")) then
			strShowErr = "<li>请输入一个有效数字</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
	    User_Conn.execute("Update FS_ME_BuyBag Set BuyNumber = "& CintStr(Request.Form("ProductNum")) &"  where BuyID ="&CintStr(Request.Form("ID"))&" and UserNumber='"&Fs_User.UserNumber&"'")
		Response.Redirect "BuyBag.asp"
		Response.end
	End If 
End if
If Request.QueryString("Action") = "Del" Then
	If Request("Buyid") = "" then
		strShowErr = "<li>请选择一个商品更新</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Else
	    User_Conn.execute("Delete From FS_ME_BuyBag  where BuyID ="&CintStr(Request.QueryString("BuyID"))&" and UserNumber='"&Fs_User.UserNumber&"'")
		Response.Redirect "BuyBag.asp"
		Response.end
	End If 
End if
if Request.Form("clear")="clearall" then
	User_Conn.execute("Delete From FS_ME_BuyBag  where UserNumber='"&Fs_User.UserNumber&"'")
	Response.Redirect "BuyBag.asp"
	Response.end
End if
if request.Form("action")="makeorder" then

	Dim productIDS,OrderRs,BagRs,OrderDetail,OrderNumber,ExpressCompany
	productIDS=FormatIntArr(DelHeadAndEndDot(request.Form("productIDS")))
	Set OrderRs=Server.CreateObject(G_FS_RS)
	Set BagRs=Server.CreateObject(G_FS_RS)
	Set OrderDetail=Server.CreateObject(G_FS_RS)
	OrderRs.open "Select * From FS_ME_Order where 1=2",User_Conn,1,3
	BagRs.open "Select mid,BuyType,AddTime,UserNumber,BuyMoney,BuyNumber from FS_ME_BuyBag where MID in("&productIDS&")",User_Conn,1,1
	if not BagRs.eof then
		OrderRs.addnew
		Fs_User.Name=session("FS_UserName")
		OrderNumber=GetRamCode(6)&"-"&right(year(now),2)&month(now)&day(now)&hour(now)&minute(now)
		'OrderRs("MoneyAmount")=NoSqlHack(Request.Form("UserName"))
		OrderRs("OrderNumber")=OrderNumber
		OrderRs("OrderType")=1
		OrderRs("AddTime")=Now()
		OrderRs("M_UserName")=NoSqlHack(Request.Form("UserName"))
		OrderRs("UserNumber")=session("FS_UserNumber")
		OrderRs("M_City")=NoSqlHack(Request.Form("M_City"))
		OrderRs("M_Province")=NoSqlHack(Request.Form("M_Province"))
		OrderRs("M_Address")=NoSqlHack(Request.Form("M_Address"))
		OrderRs("M_Tel")=NoSqlHack(Request.Form("M_Tel"))
		OrderRs("M_PostCode")=NoSqlHack(Request.Form("M_PostCode"))
		OrderRs("M_Mobile")=NoSqlHack(Request.Form("Mobile"))
		OrderRs("isPay")=0
		if Request.Form("M_PayStyle")="0" then
			OrderRs("isOnPay")=1
		else
			OrderRs("isOnPay")=0
		end if
		ExpressCompany=trim(request.Form("ExpressCompany"))
		if ExpressCompany="" then
			ExpressCompany=0
		End if
		OrderRs("M_ExpressCompany")=NoSqlHack(ExpressCompany)
		OrderRs("M_Sex")=NoSqlHack(request.Form("Sex"))
		OrderRs("M_Type")=NoSqlHack(request.Form("M_Type"))
		OrderRs("M_PayStyle")=NoSqlHack(Request.Form("M_PayStyle"))
		OrderRs("M_state")=0
		OrderRs("isLock")=1
		OrderRs("IsSuccess")=0
		OrderRs("LackDeal")=NoSqlHack(request.Form("LackDeal"))
		OrderRs.update
	End if
	OrderRs.close
	dim t_totle
	t_totle = 0
	while not BagRs.eof
		User_Conn.execute("Insert into FS_ME_Order_Detail (OrderNumber,ProductID,ProductNumber,M_state,Moneys) values('"&NoSqlHack(OrderNumber)&"',"&BagRs("mid")&","&BagRs("BuyNumber")&",0,"&Clng(BagRs("BuyMoney"))*Clng(BagRs("BuyNumber"))&")")
		t_totle = t_totle + Clng(BagRs("BuyMoney"))*Clng(BagRs("BuyNumber"))
		bagRs.movenext
	Wend
	User_Conn.execute("Update FS_ME_Order set MoneyAmount="& CintStr(t_totle) &" where OrderNumber='"& NoSqlHack(OrderNumber) &"' and UserNumber='"& session("FS_UserNumber") &"'")
	User_Conn.execute("Delete From FS_ME_BuyBag where MID in("& FormatIntArr(productIDS)&")")
	if err.number=0 then
		Response.Redirect("lib/success.asp?ErrCodes=<li>订购成功</li><li>但此定单还没进行支付，请到定单管理中支付.</li><li>本次操作的定单编号："& OrderNumber &"</li><li><a href=../order.asp><span class=tx>返回定单管理页面</span></a></font></li>&ErrorURL=../buybag.asp")
	End if
End if
Set obj_buy_rs = Server.CreateObject(G_FS_RS)
obj_buySQL_ = "Select BuyID,MID,BuyType,AddTime,UserNumber,BuyMoney,BuyNumber,Content From FS_ME_BuyBag where UserNumber='"& NoSQLHack(Replace(Replace(Fs_User.UserNumber,"'",""),Chr(39),""))&"'"
obj_buy_rs.Open obj_buySQL_,User_Conn,1,3
Dim GetBuyCount
GetBuyCount = obj_buy_rs.RecordCount
If GetBuyCount = 0 then
	tmp_tf_ = 1 
End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title>欢迎用户<%=Fs_User.UserName%>来到<%=GetUserSystemTitle%>-购物车</title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,风讯,风讯网站内容管理系统,风讯系统,风讯新闻系统,风讯商城,风讯b2c,新闻系统,CMS,域名空间,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="javascript" src="../FS_Inc/prototype.js"></script>
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
            <a href="main.asp">会员首页</a> &gt;&gt; 我的购物车</td>
        </tr>
      </table> 
      <table width="98%" border=0 align="center" cellPadding=1 cellSpacing=1 background="" class="table">
        <tbody>
          <tr  class="xingmu" align="center"> 
            <td height="26"  class="xingmu">商品名称</td>
            <td height="26"  class="xingmu">类型</td>
            <td  class="xingmu" align="center">商品单价(RMB)</td>
            <td width="99"  class="xingmu">更新数量</td>
            <td width="67"  class="xingmu">小计</td>
            <td width="39"  class="xingmu">有货</td>
            <td width="70"  class="xingmu">删除</td>
          </tr>
          <%
			dim tmp_product_rs,tmp_produts_title,tmp_NewPrice,Nowmoney,tmp_OldPrice,tmp__tf,tmp_Stockpile,sum_tmp,tmp_stro,tmp_up,tmp_href,empty_tf
			sum_tmp = 0
			empty_tf=False
			If obj_buy_rs.Eof Then 
				empty_tf=true
			End if
			Do while Not obj_buy_rs.Eof 
				dim tmp_type_,tmp_ptype_
				tmp_type_ = obj_buy_rs("BuyType")
				productIDS=obj_buy_rs("mid")&","&productIDS
				select case tmp_type_
						case 0
							set tmp_product_rs = User_conn.execute("select GroupID,GroupMoney,GroupName from FS_ME_Group where GroupID="&CintStr(obj_buy_rs("MID"))&"")
							tmp_ptype_ ="会员权限"
							if tmp_product_rs.eof then
								tmp_produts_title = "<span class=""tx"">此权限已经被管理员删除</span>"
								tmp_NewPrice = "--"
								tmp_OldPrice = "--"
								tmp_Stockpile = ""
								tmp__tf =1
								tmp_href=""
							else
								tmp_produts_title = tmp_product_rs("GroupName")
								tmp_NewPrice = tmp_product_rs("GroupMoney")
								tmp_OldPrice = "0"
								tmp__tf = 0
								tmp_Stockpile =1000000
								tmp_href=""
							End if
							tmp_product_rs.close:set tmp_product_rs=nothing
						case 1
							tmp_ptype_ ="商品"
							set tmp_product_rs = conn.execute("select id,ProductTitle,NewPrice,OldPrice,Stockpile,Mail_Money from FS_MS_Products where id="&obj_buy_rs("MID")&" and ReycleTF=0")
							if tmp_product_rs.eof then
								tmp_produts_title = "<span class=""tx"">此商品已经被管理员删除</span>"
								tmp_NewPrice = "--"
								tmp_OldPrice = "--"
								tmp_Stockpile = ""
								tmp__tf =1
								tmp_href=""
							else
								tmp_produts_title = tmp_product_rs("ProductTitle")
								If tmp_product_rs("Mail_Money") = "" Or Isnull(tmp_product_rs("Mail_Money")) Then
								    Nowmoney = tmp_product_rs("NewPrice")
								Else
								    Nowmoney = tmp_product_rs("Mail_Money") + tmp_product_rs("NewPrice")
								End if
								tmp_NewPrice = formatCurrency(Nowmoney)
								tmp_OldPrice = "<strike>" & formatCurrency(tmp_product_rs("OldPrice")) &"</strike>"
								tmp__tf = 0
								tmp_Stockpile =tmp_product_rs("Stockpile")
								tmp_href=""
							End if
							tmp_product_rs.close:set tmp_product_rs = nothing
						case 2
							set tmp_product_rs = User_conn.execute("select CardID,CardNumber,CardMoney,isBuy from FS_ME_Card where CardID="&obj_buy_rs("MID")&" and isBuy=0")
							tmp_ptype_ ="点卡"
							if tmp_product_rs.eof then
								tmp_produts_title = "<span class=""tx"">点卡已经被管理员删除</span>"
								tmp_NewPrice = "--"
								tmp_OldPrice = "--"
								tmp_Stockpile = ""
								tmp__tf =1
								tmp_href=""
							else
								tmp_produts_title = tmp_product_rs("CardNumber")
								tmp_NewPrice = tmp_product_rs("CardMoney")
								tmp_OldPrice = "0"
								tmp__tf = 0
								tmp_Stockpile =1000000
								tmp_href=""
							End if
							tmp_product_rs.close:set tmp_product_rs=nothing
			 end select
			 if tmp_type_=1 then
				 if tmp_Stockpile>obj_buy_rs("BuyNumber") or tmp_Stockpile=obj_buy_rs("BuyNumber") then
					if tmp__tf = 1 then
						 tmp_stro="--"
					else
						 tmp_stro="有"
					End if
					 tmp_up = 1
				 else
					 tmp_stro="<span class=""tx"">无货</span>"
					 tmp_up = 0
				 end if
			Else
					 tmp_up = 1
			End if
		 %>
          <tr class="hback" align="center"> 
            <form method=POST action="BuyBag.asp?Action_1=Update&ID=<% =obj_buy_rs("Buyid")%>" name=BuyForm>
              <td width="256" height="26" align="left">·<% = tmp_produts_title%></td>
              <td width="61" align="left"><div align="center"> 
                  <% = 	tmp_ptype_%>
                </div></td>
              <td><input name="Nowmoney" type="hidden" id="Nowmoney" value="<%=tmp_NewPrice%>">
                含运费：
              <% = tmp_NewPrice %></td>
              <td><input name="ProductNum" type="text" id="ProductNum" value="<% = obj_buy_rs("BuyNumber")%>" size="5"> 
                <input type="submit" name="Submit" value="更新"> <input name="id" type="hidden" id="id2" value="<% = obj_buy_rs("Buyid")%>">              </td>
              <td> <div align="right"> 
                  <%
				if tmp__tf =0   then
					Response.Write  formatnumber(obj_buy_rs("BuyNumber")*tmp_NewPrice,2,-1)
				Else
					Response.Write  "--"
				End if
				%>
                </div></td>
              <td><%=tmp_stro%> <input name="tmp_tf" type="hidden" id="tmp_tf" value="<% = tmp_up %>"></td>
              <td><div align="center"><a href="BuyBag.asp?Action=Del&Buyid=<% = obj_buy_rs("BuyID")%>" onClick="{if(confirm('确定要删除吗？')){return true;}return false;}">删除</a></div></td>
            </form>
          </tr>
          <%
			if tmp__tf =0 then
				'if tmp_Stockpile>obj_buy_rs("BuyNumber") or tmp_Stockpile=obj_buy_rs("BuyNumber")then
					sum_tmp = sum_tmp + obj_buy_rs("BuyNumber")*tmp_NewPrice
				'End if
			End if
				obj_buy_rs.MoveNext
			Loop
			%>
        <form action="BuyBag.asp" method="post" name="BuyForm_1" onClick="">
          <tr class="hback"> 
            <td height="26" colspan="7"><div align="right" class="tx"> 
                <input type="button" name="Submit3" value="清空购物车"  onClick="document.BuyForm_1.clear.value='clearall';{if(confirm('确定清空购物车吗？')){this.document.BuyForm_1.submit();return true;}return false;}">
                <input name="clear" type="hidden" id="clear" value="">
                <input name="Action" type="hidden" id="Action">
                <input type="button" name="btnRefresh" value="刷新购物车" id="btnRefresh3" class="button" onClick="location.reload()" >
                <input class="button" type="button" value=" 继续购物 " name="b3"  onclick="history.back()">
                <input name="productIDS" value="<%=productIDS%>" type="hidden"/>
                <%
				  If tmp_tf_ = 1 then
				  %>
                <input type="submit" name="Submit2" value="库存不够或者没有商品记录"  class="button" disabled>
				<%ElseIf empty_tf then%>
				<input type="button" name="Submit2" value="去收银台"  class="button" disabled>
                <%Else%>
                <input type="button" name="Submit2" value="去收银台"  class="button" onClick="buyIt()">
                <%End if%>
                　<font style="font-family:宋体">共计金额：<%=formatnumber(sum_tmp,2,-1)%> RMB</font></div></td>
          </tr>
          <tr class="hback">
            <td height="26" colspan="7"><div id="ExpressPane"></div></td>
          </tr>
        </form>
      </table>
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
<script type="text/javascript">
	var buyIt = function() {
		var myAjax = new Ajax.Request("mall/ChoiceExpress.asp?rnd=" + Math.random(), {
			method: 'get',
			onComplete: function(OriginalRequest) {
				$('ExpressPane').innerHTML = OriginalRequest.responseText;
			}
		}
		);
	};
	var makeOrder = function() {
		if (confirm("确认进行支付操作？")) {
			$("Action").value = "makeorder"
			$("BuyForm_1").submit();
		}
	};
</script>
<%
obj_buy_rs.close:set obj_buy_rs = nothing
Set Fs_User = Nothing
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->