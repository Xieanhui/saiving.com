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
Dim pid,cityRs
pid=NoSqlHack(trim(request("pid")))
Set cityRs=Conn.execute("Select CID,City from FS_AP_City where PID="&CintStr(pid))
response.Write("<select name=""sel_city"" onChange=""setValue(this,$('hid_city'))"">"&vbcrlf)
response.Write("<option value="""">«Î—°‘Ò≥« –</option>"&vbcrlf)
while not cityRs.eof
	response.Write("<option value="""&cityRs("City")&""">"&cityRs("City")&"</option>"&vbcrlf)
	cityRs.movenext
wend
response.Write("</select>"&vbcrlf)
Set Conn=nothing
response.End()
%>