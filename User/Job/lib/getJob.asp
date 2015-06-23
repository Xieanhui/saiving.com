<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%> 
<% Option Explicit %>
<%Session.CodePage=936%> 
<!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_Inc/Function.asp" -->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../lib/strlib.asp" -->
<!--#include file="../../lib/UserCheck.asp" -->
<%
response.Charset="GB2312"
Dim tid,jobRs
tid=NoSqlHack(trim(request("tid")))
Set jobRs=Conn.execute("Select Job  from FS_AP_Job where TID="&CintStr(tid))
response.Write("<select name=""sel_job"" onChange=""setValue(this,$('hid_job'))"">"&vbcrlf)
response.Write("<option value="""">«Î—°‘Ò––“µ</option>"&vbcrlf)
while not jobRs.eof
	response.Write("<option value="""&jobRs("Job")&""">"&jobRs("Job")&"</option>"&vbcrlf)
	jobRs.movenext
wend
response.Write("</select>"&vbcrlf)
response.End()
%>





