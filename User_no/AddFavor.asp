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
'0Ϊ���ţ�1Ϊ���أ�2Ϊ��ҵ��Ա���˲�ϵͳ���ã���3Ϊ������Ϣ��4Ϊ��Ʒ��5Ϊ������Ϣ,6��Ƹ,7��־
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
	Response.Redirect("lib/error.asp?ErrCodes=<li>����Ĳ�����ID����Ϊ����</li>")
End If
If Trim(FavoriteTypeDesc)="" Then
	Response.Redirect("lib/error.asp?ErrCodes=<li>����Ĳ���������Ϊ��</li>")
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
	Response.Redirect("lib/error.asp?ErrCodes=<li>��������</li>")
Else
	Response.Redirect("lib/success.asp?ErrCodes=<li>��ӳɹ�</li>")
End if
%>
<%
Set Fs_User = Nothing
%>





