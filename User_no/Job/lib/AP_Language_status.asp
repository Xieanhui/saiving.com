<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%> 
<% Option Explicit %>
<%Session.CodePage=936%> <!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_Inc/Function.asp" -->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../lib/strlib.asp" -->
<!--#include file="../../lib/UserCheck.asp" -->
<%
Response.Charset="GB2312"
Dim resumeRs
MF_Default_Conn
if session("FS_UserNumber")<>"" then
	Set resumeRs=Conn.execute("select bid,Language,Degree from FS_AP_Resume_Language where UserNumber='"&session("FS_UserNumber")&"'")
	if not resumeRs.eof then 
			Response.Write("<table border='0' width='100%' class='table' align='center'>"&vbcrlf)
			Response.Write("<tr height='20'>"&vbcrlf)
			Response.Write("<td class='xingmu' align='center'>����</td>"&vbcrlf)
			Response.Write("<td class='xingmu' align='center'>�ȼ�</td>"&vbcrlf)
			Response.Write("<td class='xingmu' align='center'>����</td>"&vbcrlf)
			Response.Write("</tr>"&vbcrlf)
			while not resumeRs.eof
				Response.Write("<tr height='30'>"&vbcrlf)
				Response.Write("<td class='hback' align='center'>"&resumeRs("Language")&"</td>"&vbcrlf)
				Response.Write("<td class='hback' align='center'>"&resumeRs("Degree")&"</td>"&vbcrlf)
				Response.Write("<td class='hback' align='center'><a href='#' onClick=""getResumeForm('resume_container','language','"&resumeRs("bid")&"','edit')"">�޸�/�鿴</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href='#' onClick=""Delete('language','"&resumeRs("bid")&"')"">ɾ��</a></td>"&vbcrlf)
				Response.Write("</tr>"&vbcrlf)	
				resumeRs.movenext
			wend
			Response.Write("</table>"&vbcrlf)
	else
		Set resumeRs=Conn.execute("select UserNumber,Uname,Sex,PictureExt,Birthday,CertificateClass,CertificateNo,CurrentWage,CurrencyType,WorkAge,Province,City,HomeTel,CompanyTel,Mobile,Email,QQ,isPublic,click,lastTime from FS_AP_Resume_BaseInfo where UserNumber='"&session("FS_UserNumber")&"'")
		Response.Write("<table border='0' width='100%' class='table' align='center'>"&vbcrlf)
		Response.Write("<tr height='30'>"&vbcrlf)
		if resumeRs.eof then
			Response.Write("<td>�㻹û�м�����<a href='#' onClick=""getResumeForm('resume_container','baseinfo','')"">������ļ���</a></td>"&vbcrlf)
		Else
			Response.Write("<td>�㻹û����д�������������ڴ������������������</td>"&vbcrlf)
		End if
		Response.Write("</tr>"&vbcrlf)
		Response.Write("</table>"&vbcrlf)
	End if
End if
Conn.close
Set Conn=nothing
set resumeRs=nothing
Response.End()
%>






