<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
Dim action,favorRs
Dim FID,FavoriteTypeDesc,FavoriteType,UserNumber,AddTime,FavoClassID
FID=CintStr(request.querystring("id"))
FavoriteTypeDesc=NoSqlHack(request.querystring("type"))
'0为新闻，1为下载，2为企业会员（人才系统有用），3为供求信息，4为商品，5为房产信息,6招聘,7日志
Select Case LCase(FavoriteTypeDesc)
	Case "ns" FavoriteType=0
	Case "ds" FavoriteType=1
	Case "corp" FavoriteType=2
	Case "sd" FavoriteType=3
	Case "ms" FavoriteType=4
	Case "hs" FavoriteType=5
	Case "ap" FavoriteType=6
	Case "log" FavoriteType=7
End select
UserNumber=Session("FS_UserNumber")
AddTime=DateValue(Now())
If Trim(FID)="" or not isnumeric(FID) Then
	Response.Redirect("lib/error.asp?ErrCodes=<li>错误的参数，ID必须为数字</li>")
End If
If Trim(FavoriteTypeDesc)="" Then
	Response.Redirect("lib/error.asp?ErrCodes=<li>错误的参数，类型为空</li>")
End If
Set favorRs=Server.CreateObject(G_FS_RS)
favorRs.open "select * from FS_ME_Favorite where 1=0",User_Conn,1,3
favorRs.addNew
favorRs("FID")=FID
favorRs("FavoriteType")=NoSqlHack(FavoriteType)
favorRs("UserNumber")=NoSqlHack(UserNumber)
favorRs("AddTime")=AddTime
favorRs("FavoClassID")=0
favorRs.update
favorRs.close
If Err.number<>0 Then 
	Response.Redirect("lib/error.asp?ErrCodes=<li>发生错误</li>")
Else
	Response.Redirect("lib/success.asp?ErrCodes=<li>添加成功</li>")
End if
%>
<%
Set Fs_User = Nothing
%>





