<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<%
''Code By Ken For Fs_FreeLabel
session.CodePage="936"
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.Charset="GB2312"
Server.ScriptTimeOut=9999999

Dim Conn,strShowErr
MF_Default_Conn
MF_Session_TF 
If Not MF_Check_Pop_TF("MF_sPublic") Then Err_Show


Dim Sql_Str,Rs
Sql_Str = Trim(Request.QueryString("Act"))
If Sql_Str = "" Then
	Response.Write "��������sql���Ϊ��"
	Response.End
End If
If Left(LCase(Sql_Str),10) <> "select top" Then
	Response.Write "��������sql��䲻��ȷ"
	Response.End
End If

On Error Resume Next
Set Rs = Server.CreateObject(G_FS_RS)
Rs.Open Sql_Str,Conn,1,1

If Err.Number <> 0 Then
	Response.Write "������������ԭ��" & Err.Description
	Response.End
Else
	Response.Write "SQl�����ȷ���鵽��¼" & Clng(Rs.RecordCount) & "��"
	Response.End
End If

Rs.CLose : Set Rs = Nothing
Conn.Close : Set Conn = NOthing		
%>
	





