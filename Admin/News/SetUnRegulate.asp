<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="lib/cls_main.asp" -->
<!--#include file="NF_News_Function.asp"-->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn
MF_Default_Conn
MF_Session_TF
'session�ж�
'********************
Call MF_Get_Session(0,1,1)
if not MF_Check_Pop_TF("NS_UnRl") then Err_Show
dim Fs_news
set Fs_news = new Cls_News
'����һ�������Ӳ���������
Dim NewsID,MainNewsID,NF_MainNews,NewsTitle,i,TempTitle,UpdateRs,TempRs,OldNewsID,Rows,StrSql
if trim(Request.QueryString("Action"))="Add" or trim(Request.QueryString("Action"))="Edit" then
	NewsID = Replace(Request.Form("NewsID")," ","")
	If trim(Request.QueryString("Action"))="Edit" Then
		if not MF_Check_Pop_TF("NS046") then Err_Show
		MainNewsID=NoSqlHack(Request.QueryString("MainNewsID"))
		Conn.execute("Delete From [FS_NS_News_Unrgl] where UnregulatedMain='"&MainNewsID&"'")
	Else
		if not MF_Check_Pop_TF("NS045") then Err_Show
		MainNewsID = Fs_News.GetRamCode(15)
	End If
	
	If NewsID="" Then
		Response.write "<script>alert('����Ĳ���,û��ѡ������');history.back();</script>" : Response.End
	End If
	If MainNewsID="" Then
		Response.write "<script>alert('����������Ų���');history.back();</script>" : Response.End
	End If
	'ȡ�ӱ����������
	NewsTitle=split(NewsID,",")
	For i = LBound(NewsTitle) To UBound(NewsTitle)
		TempTitle=Request.Form("NewsTitle"&NewsTitle(i))
		Rows=Request.form("Row"&NewsTitle(i))
		'Response.Write("��:"&MainNewsID&"ID:"&NewsTitle(i)&"����:"&TempTitle&"����:"&Rows&"<br>")
		StrSql="Insert Into FS_NS_News_Unrgl (UnregulatedMain,MainUnregNewsID,UnregNewsName,[Rows]) values('"&NoSqlHack(MainNewsID)&"','"&NoSqlHack(NewsTitle(i))&"','"&NoSqlHack(TempTitle)&"',"&NoSqlHack(Rows)&")"
		Conn.Execute(StrSql)
	Next
	Response.write "<script>alert('�����ɹ�');location.href='DefineNews_Manage.asp';</script>"
end if
Set Conn=nothing
%>





