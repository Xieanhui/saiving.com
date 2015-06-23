<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="lib/cls_main.asp" -->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn
MF_Default_Conn
MF_Session_TF
if not MF_Check_Pop_TF("NS047") then Err_Show
dim Fs_news
set Fs_news = new Cls_News
Dim Action,NewsID,C_NewsID,i,strShowErr
Action=NoSqlHack(Request.QueryString("Action"))
NewsID=NoSqlHack(Request.QueryString("NewsID"))
IF Action<>"" then
	If Action="signDel" Then
	IF NewsID="" Then
			strShowErr="<li>你必须选择一项再删除</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
	End If
	'进行单条记录的删除
		Conn.execute("Delete From FS_NS_News_Unrgl where UnregulatedMain='"&NewsID&"'")
		Response.write "<script>alert('操作成功');location.href='DefineNews_Manage.asp';</script>"
	End If
	'进行批量删除
	If Action="del" Then
		C_NewsID=NoSqlHack(Replace(Request.form("C_NewsID")," ",""))
		If C_NewsID="" Then
			strShowErr="<li>你必须选择一项再删除</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End If
		C_NewsID=split(C_NewsID,",")
		For i=LBound(C_NewsID) to UBound(C_NewsID)
			Conn.execute("Delete From FS_NS_News_Unrgl where UnregulatedMain='"&C_NewsID(i)&"'")
			Response.write "<script>alert('操作成功');location.href='DefineNews_Manage.asp';</script>"
		Next
	End If
Else
	strShowErr="<li>你必须选择一项再删除</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
End if
Set Conn=nothing
%>





