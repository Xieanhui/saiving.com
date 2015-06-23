<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp"-->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_InterFace/ns_Function.asp"-->
<!--#include file="../FS_Inc/Function.asp" -->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn
MF_Default_Conn
MF_User_Conn
dim ClassID,NoteRs,Topic,Content,Face,IsUser,isTop,IsAdmin
Dim Configobj,PageS,sql,MSTitle,ShowIP,s_IsUser,s_UserMember
Set Configobj=server.CreateObject (G_FS_RS)
sql="select ID,IsUser,UserMember From FS_WS_Config"
configobj.open sql,Conn,1,1
if not configobj.eof then
	s_IsUser = configobj("IsUser")
	s_UserMember = configobj("UserMember")
	if s_UserMember="" or not isnumeric(s_UserMember) then
		s_UserMember = 0
	else
		s_UserMember = NoSqlHack(s_UserMember)
	end if
end if
set configobj=nothing
if s_IsUser <> "0" then
	if session("FS_UserName")="" then
		response.Write"未开放匿名发布帖！"
		response.end
	end if
end if
Set NoteRs=Server.CreateObject(G_FS_RS)
if NoSqlHack(Request("Act"))="Add" then
	ClassID= NoSQLHack(Request.form("ClassID"))
	Topic= NoSQLHack(Request.form("Topic"))
	IsTop= NoSQLHack(Request.form("Style"))
	Content= NoHtmlHackInput(NoSQLHack(Request.form("Content")))
	IsAdmin= NoSQLHack(NoSQLHack(Request.form("IsAdmin")))
	Face= NoSQLHack(NoSQLHack(Request.Form("FaceNum")))
	if IsAdmin="" then
		IsAdmin="0"
	end if
	if ClassID="" then
		Response.write("<script>alert('参数出错!');</script>")
		response.end
	end if
	if Topic="" then
		Response.write("<script>alert('标题不能为空');</script>")
		Response.end
	end if
	if isTop="" then
		Response.write("<script>alert('错误参数');</script>")
		Response.end
	end if
	if Content="" then
		Response.write("<script>alert('内容不能为空');</script>")
		response.end
	end if
	NoteRs.open "Select * from FS_WS_BBS where id=0",Conn,1,3
	NoteRs.Addnew
	NoteRs("ClassID")=ClassID
	if session("FS_UserName")<>"" then
		NoteRs("User")=session("FS_UserName")
	else
		NoteRs("User")="游客"
	end if
	NoteRs("Topic")=Topic
	NoteRs("State")=conn.execute("select top 1 isaut from FS_WS_Config")(0)
	NoteRs("Body")=Content
	NoteRs("AddDate")=now()
	NoteRs("IsTop")=IsTop
	NoteRs("Style")="普通"
	NoteRs("IsAdmin")=IsAdmin
	if session("FS_UserName")<>"" then
		NoteRs("LastUpdateUser")=session("FS_UserName")
	else
		NoteRs("LastUpdateUser")="游客"
	end if
	NoteRs("Face")=Face
	NoteRs("IP")=NoSqlHack(Request.ServerVariables("REMOTE_ADDR"))
	NoteRs.update
	Set NoteRs=nothing
	'更新会员积分
	if session("FS_UserName")<>"" then
		User_Conn.execute("Update FS_ME_Users set Integral=Integral+"& s_UserMember &" where UserName='"& session("FS_UserName")&"'")
		if s_UserMember<>0 then
			dim f_AddlogObj
			Set f_AddlogObj = server.CreateObject(G_FS_RS)
			f_AddlogObj.open "select  * From FS_ME_Log where 1=0",User_Conn,1,3
			f_AddlogObj.addnew
			f_AddlogObj("LogType")="其他"
			f_AddlogObj("UserNumber")=GetFriendNumber(session("FS_UserName"))
			f_AddlogObj("points")=s_UserMember
			f_AddlogObj("moneys")=0
			f_AddlogObj("LogTime")=Now
			f_AddlogObj("LogContent")="发表帖子增加积分"
			f_AddlogObj("Logstyle")=0
			f_AddlogObj.update
			f_AddlogObj.close
			set f_AddlogObj = nothing
		end if 
	end if
	Response.write "<script>history.go(-2);</script>"
end if
Function GetFriendNumber(f_strNumber)
	Dim RsGetFriendNumber
	Set RsGetFriendNumber = User_Conn.Execute("Select UserNumber From FS_ME_Users Where UserName = '"& NoSQLHack(f_strNumber) &"'")
	If  Not RsGetFriendNumber.eof  Then 
		GetFriendNumber = RsGetFriendNumber("UserNumber")
	End If 
	set RsGetFriendNumber = nothing
End Function 
Set Conn=nothing
%>






