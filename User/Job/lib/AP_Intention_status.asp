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
	Set resumeRs=Conn.execute("select bid,WorkType,Salary,SelfAppraise from FS_AP_Resume_Intention where UserNumber='"&session("FS_UserNumber")&"'")
	if not resumeRs.eof then 
		Dim WorkType,Salary
		select case resumeRs("WorkType")
			case "2" WorkType="��ְ"
			case "3" WorkType="ʵϰ"
			case "4" WorkType="ȫְ/��ְ"
			case else WorkType="ȫְ"
		End select
		select case resumeRs("Salary")
			case "1" Salary="1500����"
			case "2" Salary="1500-1999"
			case "3" Salary="2000-2999"
			case "4" Salary="3000-4499"
			case "5" Salary="4500-5999"
			case "6" Salary="6000-7999"
			case "7" Salary="8000-9999"
			case "8" Salary="10000-14999"
			case "9" Salary="15000-19999"
			case "10" Salary="20000-29999"
			case "11" Salary="30000-49999"
			case else Salary="50000������"

		End select
		Response.Write("<table border='0' width='100%' class='table' align='center'>"&vbcrlf)
		Response.Write("<tr height='20'>"&vbcrlf)
		Response.Write("<td class='xingmu' align='center'>��������</td>"&vbcrlf)
		Response.Write("<td class='xingmu' align='center'>����н��</td>"&vbcrlf)
		Response.Write("<td class='xingmu' align='center'>����</td>"&vbcrlf)
		Response.Write("</tr>"&vbcrlf)
		Response.Write("<tr  height='30'>"&vbcrlf)
		Response.Write("<td class='hback' align='center'>"&WorkType&"</td>"&vbcrlf)
		Response.Write("<td class='hback' align='center'>"&Salary&"</td>"&vbcrlf)
		Response.Write("<td class='hback' align='center'><a href='#' onClick=""getResumeForm('resume_container','intention','"&resumeRs("bid")&"','edit')"">�޸�/�鿴</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href='#' onClick=""Delete('intention','"&resumeRs("bid")&"')"">ɾ��</a></td>"&vbcrlf)
		Response.Write("</tr>"&vbcrlf)	
		Response.Write("</table>"&vbcrlf)
	else
		Set resumeRs=Conn.execute("select UserNumber,Uname,Sex,PictureExt,Birthday,CertificateClass,CertificateNo,CurrentWage,CurrencyType,WorkAge,Province,City,HomeTel,CompanyTel,Mobile,Email,QQ,isPublic,click,lastTime from FS_AP_Resume_BaseInfo where UserNumber='"&session("FS_UserNumber")&"'")
		Response.Write("<table border='0' width='100%' class='table' align='center'>"&vbcrlf)
		Response.Write("<tr height='30'>"&vbcrlf)
		if resumeRs.eof then
			Response.Write("<td>�㻹û�м�����<a href='#' onClick=""getResumeForm('resume_container','baseinfo')"">������ļ���</a></td>"&vbcrlf)
		Else
			Response.Write("<td>�㻹û����д��ְ�������ڴ��������ְ����</td>"&vbcrlf)
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






