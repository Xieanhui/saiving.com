<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_InterFace/Dynamic_Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
Dim obj_buy_rs,obj_buySQL_,tmp_tf_,Dynamic_HTML
Dynamic_HTML = Get_Dynamic_Refresh_Content(G_MALL_CART_TEMPLET,"","MF",0,"")

If Request("Action") = "Update" Then
	If Request("id") = "" then
		strShowErr = "<li>请选择一个商品更新</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Else
		if Not isnumeric(Request("ProductNum")) then
			strShowErr = "<li>请输入一个有效数字</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
	    User_Conn.execute("Update FS_ME_BuyBag Set BuyNumber = "& CintStr(Request("ProductNum")) &"  where BuyID ="&CintStr(Request("ID"))&" and UserNumber='"&Fs_User.UserNumber&"'")
		Response.Redirect "Cart.asp"
		Response.end
	End If
End if

If Request("Action") = "Del" Then
	If Request("Buyid") = "" then
		strShowErr = "<li>请选择一个商品更新</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Else
	    User_Conn.execute("Delete From FS_ME_BuyBag  where BuyID ="&CintStr(Request.QueryString("BuyID"))&" and UserNumber='"&Fs_User.UserNumber&"'")
		Response.Redirect "Cart.asp"
		Response.end
	End If 
End if
if Request("Action")="clearall" then
	User_Conn.execute("Delete From FS_ME_BuyBag  where UserNumber='"&Fs_User.UserNumber&"'")
	Response.Redirect "Cart.asp"
	Response.end
End if
if Request("Action")="makeorder" then
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
		Response.Redirect("lib/success.asp?ErrCodes=<li>订购成功</li><li>但此定单还没进行支付，请到定单管理中支付.</li><li>本次操作的定单编号："& OrderNumber &"</li><li><a href=../order.asp><span class=tx>返回定单管理页面</span></a></font></li>&ErrorURL=../Cart.asp")
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

Dim CartHTML
CartHTML="<form action=""Cart.asp"" method=""post"" id=""CartForm""><table width=""98%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1""><thead><tr><th>商品名称</th><th>类型</th><th>商品单价(RMB)</th><th>更新数量</th><th>小计</th><th>有货</th><th>删除</th></tr></thead><tbody>"
dim tmp_product_rs,tmp_produts_title,tmp_NewPrice,Nowmoney,tmp_OldPrice,tmp__tf,tmp_Stockpile,sum_tmp,tmp_stro,tmp_up,tmp_href,empty_tf
sum_tmp = 0
empty_tf=False
If obj_buy_rs.Eof Then 
	empty_tf=true
End if
Dim CartListHTML
CartListHTML=""
while Not obj_buy_rs.Eof 
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
	CartListHTML=CartListHTML&"<tr><td align=""left"">&middot;"&tmp_produts_title &"</td><td align=""center"">"&tmp_ptype_&"</td><td><input name=""Nowmoney"" type=""hidden"" id=""Nowmoney"" value="""&tmp_NewPrice&""" />含运费："&tmp_NewPrice&"</td><td><input name=""ProductNum"" type=""text"" id=""ProductNum"" value="""&obj_buy_rs("BuyNumber")&""" size=""5"" /><input type=""button"" name=""updateNum"" onclick=""location.href='Cart.asp?Action=Update&Id="&obj_buy_rs("Buyid")&"&ProductNum='+$('ProductNum').value;"" value=""更新"" /></td><td align=""right"">"
	if tmp__tf =0 then
		CartListHTML=CartListHTML&formatnumber(obj_buy_rs("BuyNumber")*tmp_NewPrice,2,-1)
		sum_tmp = sum_tmp + obj_buy_rs("BuyNumber")*tmp_NewPrice
	Else
		CartListHTML=CartListHTML&"--"
	End if
	CartListHTML=CartListHTML&"</td><td>"&tmp_stro&"<input name=""tmp_tf"" type=""hidden"" id=""tmp_tf"" value="""&tmp_up&"""></td><td align=""center""><a href=""Cart.asp?Action=Del&Buyid="&obj_buy_rs("BuyID")&""" onclick=""{if(confirm('确定要删除吗？')){return true;}return false;}"">删除</a></td></tr>"
	obj_buy_rs.MoveNext
Wend
CartHTML=CartHTML&CartListHTML
CartHTML=CartHTML&"<tr><td colspan=""7"" align=""right""><input type=""button"" name=""btnClearAll"" value=""清空购物车"" onclick=""{if(confirm('确定清空购物车吗？')){this.form.action='Cart.asp?Action=clearall';this.form.submit();return true;}return false;}"" /><input type=""button"" name=""btnRefresh"" value=""刷新购物车"" id=""btnRefresh"" onclick=""location.reload()"" /><input type=""button"" value="" 继续购物 "" name=""b3"" onclick=""history.back()"" /><input name=""productIDS"" value="""&productIDS&""" type=""hidden"" />"

If tmp_tf_ = 1 then
	CartHTML=CartHTML&"<input type=""button"" name=""btnGO"" value=""库存不够或者没有商品记录"" disabled=""disabled"" />"
ElseIf empty_tf then
	CartHTML=CartHTML&"<input type=""button"" name=""btnGO"" value=""去收银台"" disabled=""disabled"" />"
Else
	CartHTML=CartHTML&"<input type=""button"" name=""btnGO"" value=""去收银台"" onclick=""buyIt()"" />"
End if
CartHTML=CartHTML&"<font style=""font-family: 宋体"">共计金额："&formatnumber(sum_tmp,2,-1)&" RMB</font></td></tr><tr><td height=""26"" colspan=""7""><div id=""ExpressPane""></div></td></tr></tbody></table></form><script type=""text/javascript"">var buyIt = function() { var myAjax = new Ajax.Request('mall/ChoiceExpress.asp?rnd=' + Math.random(), { method: 'get', onComplete: function(OriginalRequest) { $('ExpressPane').innerHTML = OriginalRequest.responseText; } }); }; var makeOrder = function() { if (confirm('确认进行支付操作？')) { $('CartForm').action = 'Cart.asp?Action=makeorder'; $('CartForm').submit(); } };</script>"


obj_buy_rs.close:set obj_buy_rs = nothing
Set Fs_User = Nothing

Dynamic_HTML = Replace(Dynamic_HTML,"{FS:Mall_Cart_Content}",CartHTML)
Response.Write(Dynamic_HTML)
Response.End()
%>