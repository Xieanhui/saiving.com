<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Func_Page.asp"-->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
Dim groupid,debateid,action,DebateRs,parentID
action=request.QueryString("act")
groupid=NoSqlHack(request.QueryString("classid"))
debateid=NoSqlHack(request.QueryString("debateid"))
if action="delete" then
	Set DebateRs=User_Conn.execute("Select parentID from FS_ME_GroupDebate where Debateid="&CintStr(debateID))
	if not DebateRs.eof then
		parentID=DebateRs("parentID")
	Else
		parentID=0
	ENd if
	if parentID=0 then
		User_Conn.execute("Delete From FS_ME_GroupDebate where UserNumber = '" & Fs_User.UserNumber & "' And DebateID="&CintStr(debateID))
		User_Conn.execute("Delete From FS_ME_GroupDebate where UserNumber = '" & Fs_User.UserNumber & "' And DebateID in (Select DebateID from FS_ME_GroupDebate where parentID="&CintStr(debateID)&")")
	Else
		User_Conn.execute("Delete From FS_ME_GroupDebate where UserNumber = '" & Fs_User.UserNumber & "' And DebateID="&CintStr(debateID))
	End if
ENd if
User_Conn.close
Set User_Conn=nothing
Set Fs_User = Nothing
if err.number=0 then 
	if Cint(parentID)<>0 then
		Response.Redirect("lib/success.asp?ErrCodes=<li>修改成功</li>&ErrorURL=../Debate_unit.asp?gdid="&groupid&"***DebateID="&parentID)
	Else
		Response.Redirect("lib/success.asp?ErrCodes=<li>修改成功</li>&ErrorURL=../Group_unit.asp?gdid="&groupid)
	End if
	Response.End()
else
	Response.Redirect("lib/error.asp?ErrCodes=<li>"&err.description&"</li>")
	Response.End()
end if

%>





