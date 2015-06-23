<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
Dim debateid,SelectName,result,User_Conn,debateGroupRs
MF_Default_Conn
MF_User_Conn
MF_Session_TF
debateid=NoSqlHack(Request.QueryString("debateid"))
SelectName=NoSqlHack(Request.QueryString("name"))
if debateid<>"" then
	Set debateGroupRs=server.CreateObject(G_FS_RS)
	debateGroupRs.open "select DebateID,Title,ParentID from FS_ME_GroupDebate where ParentID="&CintStr(debateid),User_Conn,1,3
	if SelectName="Topic1" and  not debateGroupRs.eof then
		result="<select name='"&SelectName&"' id='"&SelectName&"' onchange=""getDebate(this,'Topic2')"">"
		result=result&"<option value=''>请选择二级主题</option>"
	elseif not debateGroupRs.eof then
		result="<select name='"&SelectName&"' id='"&SelectName&"'>"
		result=result&"<option value=''>请选择三级主题</option>"
	end if
	while not debateGroupRs.eof
		result=result&"<option value='"&debateGroupRs("DebateID")&"'>"&debateGroupRs("Title")&"</option>"&Chr(10)&chr(13)
		debateGroupRs.movenext
	wend
	if result<>"" then
		result=result&"</select>"&Chr(10)&Chr(13)
	end if
	Response.Charset="GB2312"
	Response.Write(result)
	debateGroupRs.close
	Set debateGroupRs=nothing
	User_Conn.close
	Set User_Conn=nothing
end if
%>







