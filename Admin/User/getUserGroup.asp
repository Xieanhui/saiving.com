<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
Dim conn,typeId,result,User_Conn,GroupParaRs,i
MF_Default_Conn
MF_User_Conn
MF_Session_TF
if CintStr(Request.QueryString("id"))=1 or CintStr(Request.QueryString("id"))=0 then
	Set GroupParaRs=server.CreateObject(G_FS_RS)
	If G_IS_SQL_User_DB=0 then
		GroupParaRs.open "select GroupID,GroupName,GroupPopList from FS_ME_Group where GroupType="&CintStr(Request.QueryString("ID"))&" order by GroupPopList",User_Conn,1,3
	Else
		GroupParaRs.open "select GroupID,GroupName,GroupPopList from FS_ME_Group where GroupType="&CintStr(Request.QueryString("ID")),User_Conn,1,3
	End if
	if Request.QueryString("id")=1 then 
		if Request.QueryString("page")="UserGroup" then
			result="<select name='GroupIndex' id='GroupIndex' onchange='getGroupParam(this)'>"&Chr(10)&Chr(13)&"<option value='user'>所有个人会员组</option>"
		elseif Request.QueryString("page")="News" then
			result="<select name='GroupIndex' id='GroupIndex'>"&Chr(10)&Chr(13)&"<option value='user'>所有个人会员组</option>"		
		end if
	elseif Request.QueryString("id")=0 then 
		if Request.QueryString("page")="UserGroup" then
			result="<select name='GroupIndex' id='GroupIndex' onchange='getGroupParam(this)'>"&Chr(10)&Chr(13)&"<option value='corp'>所有企业会员组</option>"
		elseif Request.QueryString("page")="News" then
			result="<select name='GroupIndex' id='GroupIndex'>"&Chr(10)&Chr(13)&"<option value='corp'>所有企业会员组</option>"
		end if
	end if
	while not GroupParaRs.eof
		result=result&"<option value='"&GroupParaRs("Groupid")&"'>"&GroupParaRs("GroupName")&"</option>"&Chr(10)&chr(13)
		GroupParaRs.movenext
	wend
	result=result&"</select>"&Chr(10)&Chr(13)
	Response.Charset="GB2312"
	Response.Write(result)
	GroupParaRs.close
	Set GroupParaRs=nothing
	User_Conn.close
	Set User_Conn=nothing
End if
%>







