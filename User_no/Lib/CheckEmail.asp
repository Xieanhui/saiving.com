<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="Cls_user.asp" -->
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<%
Dim Conn,User_Conn
MF_Default_Conn
MF_User_Conn
Dim ReturnValue,Email
Email=Replace(replace(NoSqlHack(Request("Email")),"'","''"),Chr(39),"")
If Email="" then 
	Response.Write("����д�õ����ʼ�")
	Response.end
End if
If len(Email)<6 then 
	Response.Write("�����ʼ�����д��ȷ")
	Response.end
End if
if Email<> "" then
if Instr(Email,"@")=0 or Instr(Email,".")=0then
	Response.Write("�����ʼ�����д��ȷ")
	Response.end
end if
end if
dim Fs_User
Set Fs_User = New Cls_User
ReturnValue = Fs_User.checkEmail(Email)
if ReturnValue then
	Response.Write(""& Email &" ���ʼ�����ע��")
	Response.end
Else
		Response.Write(""& Email &" �ʼ��Ѿ���ע��")
		Response.end
End if
set Fs_User=nothing
%>
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">






