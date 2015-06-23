<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%

Dim Rs,productID
Dim productRs,price,nowString
price=0
on error resume next
productID=CintStr(request.QueryString("pid"))

if not isnumeric(productID) or productID="" then
	Response.Write("<script>history.back();</script>")
End if
Set productRs=Conn.execute("Select NewPrice,Mail_Money from FS_MS_Products where ID="& CintStr(productID))
if not productRs.eof then
	if productRs("NewPrice")="" then
		price = 0
	Else
	    If productRs("Mail_Money") = "" Or Isnull(productRs("Mail_Money")) Then
		    price = productRs("NewPrice")
	    Else
		    price = productRs("NewPrice") + productRs("Mail_Money")
	    End if
	End if
Else
	Response.Redirect("lib/error.asp?ErrCodes=<li>没有找到商品</li>&ErrorUrl="&Request.ServerVariables("HTTP_REFERER"))
End if
if G_IS_SQL_User_DB=0 then
	nowString="now()"
else
	nowString="getdate()"
end if
User_Conn.execute("Insert into FS_ME_BuyBag (mid,BuyType,AddTime,UserNumber,BuyMoney,BuyNumber) values("&productid&",1,"&nowString&",'"&session("FS_Usernumber")&"',"&price&",1)")
if err.number=0 then
	Response.Redirect("Cart.asp")
Else
	Response.Redirect("lib/error.asp?ErrCodes=<li>"&err.description&"</li>")
End if
%>
<%
Set Conn = Nothing
Set Fs_User = Nothing
%>





