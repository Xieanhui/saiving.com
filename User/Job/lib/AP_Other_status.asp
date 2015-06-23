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
	Set resumeRs=Conn.execute("select bid,Title,Content from FS_AP_Resume_Other where UserNumber='"&session("FS_UserNumber")&"'")
	if not resumeRs.eof then 
			Response.Write("<table border='0' width='100%' class='table' align='center'>"&vbcrlf)
			Response.Write("<tr height='20'>"&vbcrlf)
			Response.Write("<td class='xingmu' align='center' width='10%'>ID</td>"&vbcrlf)
			Response.Write("<td class='xingmu'>主题标题</td>"&vbcrlf)
			Response.Write("<td class='xingmu' align='center' width='20%'>操作</td>"&vbcrlf)
			Response.Write("</tr>"&vbcrlf)
			while not resumeRs.eof
				Response.Write("<tr height='30'>"&vbcrlf)
				Response.Write("<td class='hback'align='center'>"&resumeRs("BID")&"</td>"&vbcrlf)
				Response.Write("<td class='hback'>"&resumeRs("Title")&"</td>"&vbcrlf)
				Response.Write("<td class='hback'align='center'  width='20%'><a href='#' onClick=""getResumeForm('resume_container','other','"&resumeRs("bid")&"','edit')"">修改/查看</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href='#' onClick=""Delete('other','"&resumeRs("bid")&"')"">删除</a></td>"&vbcrlf)
				Response.Write("</tr>"&vbcrlf)	
			resumeRs.movenext
		wend
		Response.Write("</table>"&vbcrlf)
	else
		Set resumeRs=Conn.execute("select UserNumber,Uname,Sex,PictureExt,Birthday,CertificateClass,CertificateNo,CurrentWage,CurrencyType,WorkAge,Province,City,HomeTel,CompanyTel,Mobile,Email,QQ,isPublic,click,lastTime from FS_AP_Resume_BaseInfo where UserNumber='"&session("FS_UserNumber")&"'")
		Response.Write("<table border='0' width='100%' class='table' align='center'>"&vbcrlf)
		Response.Write("<tr height='30'>"&vbcrlf)
		if resumeRs.eof then
			Response.Write("<td>你还没有简历！<a href='#' onClick=""getResumeForm('resume_container','baseinfo','')"">创建你的简历</a></td>"&vbcrlf)
		Else
			Response.Write("<td>你还没有附加信息！现在创建</td>"&vbcrlf)
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





