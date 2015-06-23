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
	Set resumeRs=Conn.execute("select bid,FetchDate,Certificate,Score from FS_AP_Resume_Certificate where UserNumber='"&session("FS_UserNumber")&"'")
	if not resumeRs.eof then 
			Response.Write("<table border='0' width='100%' class='table' align='center'>"&vbcrlf)
			Response.Write("<tr height='20'>"&vbcrlf)
			Response.Write("<td class='xingmu' align='center'>获证时间</td>"&vbcrlf)
			Response.Write("<td class='xingmu' align='center'>证书名称</td>"&vbcrlf)
			Response.Write("<td class='xingmu' align='center'>成绩（等级）</td>"&vbcrlf)
			Response.Write("<td class='xingmu' align='center'>操作</td>"&vbcrlf)
			Response.Write("</tr>"&vbcrlf)
			while not resumeRs.eof
				Response.Write("<tr height='30'>"&vbcrlf)
				Response.Write("<td class='hback' align='center'>"&resumeRs("FetchDate")&"</td>"&vbcrlf)
				Response.Write("<td class='hback' align='center'>"&resumeRs("Certificate")&"</td>"&vbcrlf)
				Response.Write("<td class='hback' align='center'>"&resumeRs("Score")&"</td>"&vbcrlf)
				Response.Write("<td class='hback' align='center'><a href='#' onClick=""getResumeForm('resume_container','certificate','"&resumeRs("bid")&"','edit')"">修改/查看</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href='#' onClick=""Delete('certificate','"&resumeRs("bid")&"')"">删除</a></td>"&vbcrlf)
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
			Response.Write("<td>你还没有填写你获得的证书！现在创建你的证书列表</td>"&vbcrlf)
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






