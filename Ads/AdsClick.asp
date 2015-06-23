<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp"-->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_InterFace/NS_Function.asp" -->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,int_ID,o_Ad_Rs,AdTxtID,AdTxtRs,tempLinkUrl,ClickObj
MF_Default_Conn

int_ID=NoSqlHack(Request.QueryString("Location"))
If int_ID="" or isnull(int_ID) Then
	AdTxtID=NoSqlHack(Request.QueryString("AdTxtID")) 
	If AdTxtID<>"" Or Not IsNull(AdTxtID) Then
		AdTxtID=Clng(AdTxtID)
		Conn.execute("Update FS_AD_Info Set AdClickNum=AdClickNum+1 where AdID In(Select AdID From  FS_AD_TxtInfo Where Ad_TxtID="&CintStr(AdTxtID)&")")
		Set AdTxtRs=Conn.execute("Select AdID,LinkUrl From FS_AD_TxtInfo Where Ad_TxtID="&CintStr(AdTxtID))
		If Not AdTxtRs.EOf Then
			Conn.execute("Insert into FS_AD_Source(AdID,AdIPAdress,VisitTime) values("&AdTxtRs("AdID")&",'"&NoSqlHack(Request.ServerVariables("REMOTE_ADDR"))&"','"&Now()&"')")
			tempLinkUrl=AdTxtRs("LinkUrl")
			Set AdTxtRs=Nothing
			Response.Redirect(tempLinkUrl)
		Else
			Set AdTxtRs=Nothing
		End If
	Else
		Response.write("参数错误")
	End If
Else
	If Isnumeric(int_ID)=false then
		Response.write("参数错误")
	Else
		int_ID=Clng(int_ID)
	End if
	Set o_Ad_Rs=Conn.execute("Select AdLinkAdress From FS_AD_Info Where AdID="&CintStr(int_ID)&"")
	If Not o_Ad_Rs.Eof THen
		Conn.execute("Update FS_AD_Info Set AdClickNum=AdClickNum+1 where AdID="&CintStr(int_ID)&"")
		Conn.execute("Insert into FS_AD_Source(AdID,AdIPAdress,VisitTime) values("&CintStr(int_ID)&",'"&NoSqlHack(Request.ServerVariables("REMOTE_ADDR"))&"','"&Now()&"')")
		Response.Redirect(o_Ad_Rs("AdLinkAdress"))
	Else
		Response.write("参数错误")
	End If
	
	Set ClickObj = Conn.Execute("select AdLinkAdress from FS_AD_Info where AdID="&CintStr(int_ID)&"")
	If Not ClickObj.eof then
		 Response.Redirect(ClickObj("AdLinkAdress"))
	End if
	Set ClickObj=Nothing
	Set o_Ad_Rs=Nothing
	Set Conn=Nothing
End if
%><!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->





