<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%> 
<% Option Explicit %>
<%Session.CodePage=936%> 
<!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_Inc/Function.asp" -->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../lib/strlib.asp" -->
<!--#include file="../../lib/UserCheck.asp" -->
<%
Response.Charset="GB2312"
Dim resumeRs
MF_Default_Conn
if session("FS_UserNumber")<>"" then
	Set resumeRs=Conn.execute("select bid,UserNumber,Uname,Sex,PictureExt,Birthday,CertificateClass,CertificateNo,CurrentWage,CurrencyType,WorkAge,Province,City,HomeTel,CompanyTel,Mobile,Email,QQ,isPublic,click,lastTime from FS_AP_Resume_BaseInfo where UserNumber='"&session("FS_UserNumber")&"'")
	if not resumeRs.eof then 
		Dim ispublic
		if resumeRs("ispublic")="0" then 
			ispublic="����"
		else
			ispublic="<font color='red'>������</font>"
		End if
		Response.Write("<table border='0' width='100%' class='table' align='center'>"&vbcrlf)
		Response.Write("<tr height='20'>"&vbcrlf)
		Response.Write("<td class='xingmu' align='center'>��������</td>"&vbcrlf)
		Response.Write("<td class='xingmu'align='center'>���������</td>"&vbcrlf)
		Response.Write("<td class='xingmu'align='center'>����޸�ʱ��</td>"&vbcrlf)
		Response.Write("<td class='xingmu'align='center'>����</td>"&vbcrlf)
		Response.Write("</tr>"&vbcrlf)
		Response.Write("<tr  height='30' id='"&resumeRs("bid")&"'>"&vbcrlf)
		Response.Write("<td class='hback'align='center'>"&ispublic&"</td>"&vbcrlf)
		Response.Write("<td class='hback'align='center'>"&resumeRs("click")&"</td>"&vbcrlf)
		Response.Write("<td class='hback'align='center'>"&resumeRs("lastTime")&"</td>"&vbcrlf)
		Response.Write("<td class='hback'align='center'><a href='#' onClick=""getResumeForm('resume_container','baseinfo','"&resumeRs("bid")&"','edit')"">�޸�</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href='#' onClick=""Delete('baseinfo','"&resumeRs("bid")&"')"">ɾ��</a></td>"&vbcrlf)
		Response.Write("</tr>"&vbcrlf)	
		Response.Write("</table>"&vbcrlf)
	else
		Response.Write("<table border='0' width='100%' class='table' align='center'>"&vbcrlf)
		Response.Write("<tr height='30'>"&vbcrlf)
		Response.Write("<td>�㻹û�м��������ڿ�ʼ<a href='#' onClick=""getResumeForm('resume_container','baseinfo','','add')"">������ļ���</a></td>"&vbcrlf)
		Response.Write("</tr>"&vbcrlf)
		Response.Write("</table>"&vbcrlf)
	End if
End if
Conn.close
Set Conn=nothing
set resumeRs=nothing
Response.End()
%>







